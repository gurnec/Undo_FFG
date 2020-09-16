# nrbf.py - .NET Remoting Binary Format reading library for Python 3.6+
# Copyright (C) 2017, 2020 Christopher Gurnee
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions
# are met:
#
# 1. Redistributions of source code must retain the above copyright
# notice, this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright
# notice, this list of conditions and the following disclaimer in the
# documentation and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
# "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
# LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
# A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
# HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
# SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
# LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
# DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
# THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

__all__ = ['read_stream', 'serialization', 'JSONEncoder']

import enum, itertools, json, re, sys
from array       import array as Array, typecodes
from collections import OrderedDict, namedtuple
from copy        import deepcopy
from datetime    import datetime, timedelta, timezone
from decimal     import Decimal
from keyword     import iskeyword
from struct      import Struct, calcsize, pack
from contextlib  import suppress
from namedlist   import namedlist


# Decorator which adds an enum value (an int or a length-one bytes) and its associated
# reader function to a registration dict (either PrimitiveType_readers or RecordType_readers)
def _register_reader(type_dict, enum_value):
    if isinstance(enum_value, int):
        enum_value = bytes([enum_value])  # convert from int to length-one bytes
    else:
        assert isinstance(enum_value, bytes) and len(enum_value) == 1
    def decorator(read_func):
        type_dict[enum_value] = read_func
        return read_func
    return decorator


class serialization:
    def __init__(self, streamfile, can_overwrite_member = False):
        '''
        :param streamfile: a file-like object in .NET Remoting Binary Format
        :param can_overwrite_member: iff True, overwrite_member() may be called
        '''
        assert streamfile.readable()
        if can_overwrite_member:
            assert streamfile.writable()
            assert streamfile.seekable()
        self._file               = streamfile
        self._Class_by_id        = {}    # see _read_ClassInfo()
        self._objects_by_id      = {}    # all referenceable objects indexed by ObjectId
        self._root_id            = None  # the id of the single root object
        self._member_references  = []    # all seen NRBF MemberReferences, see _read_MemberReference()
        self._collection_infos   = []    # _CollectionInfo for each seen .NET collection
        self._add_overwrite_info = can_overwrite_member
        # If can_overwrite_member is True, the below is a dict indexed by an object's python id();
        # each value contains multiple _OverwriteInfo objects, and has the same layout as the object
        self._overwrite_infos_by_pyid = {} if can_overwrite_member else None

    _CollectionInfo = namedtuple('_CollectionInfo', 'obj id')      # original collection object and ObjectId
    _OverwriteInfo  = namedtuple('_OverwriteInfo',  'pos format')  # file.tell() and struct format string

    # Below are a set of "readers" and other support types, one per defined structure in
    # revisions 10.0-12.0 of the ".NET Remoting: Binary Format Data Structure" specification
    # in approximately the same order of each structure's definition. Each one reads in
    # the structure's elements, and depending on the reader type may return some or all
    # of the read values. This first section implements the Primitive Types, each one
    # returning the closest Python analogous type to the .NET Primitive Type read.
    ######## Reference: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-nrbf ########

    # A dict of all PrimitiveType readers indexed by a length-one PrimitiveTypeEnumeration bytes object
    _PrimitiveType_readers = {}

    @_register_reader(_PrimitiveType_readers, 3)
    def _read_Char(self):
        utf8_bytes = bytearray()
        while True:
            utf8_bytes += self._file.read(1)
            try:
                return utf8_bytes.decode('utf-8')
            except UnicodeDecodeError:
                if len(utf8_bytes) >= 4:
                    raise

    @_register_reader(_PrimitiveType_readers, 12)
    def _read_TimeSpan(self):
        return timedelta(microseconds= self._read_Int64() / 10)  # units of 100 nanoseconds

    @_register_reader(_PrimitiveType_readers, 13)
    def _read_DateTime(self):
        ticks = self._read_UInt64()
        kind  = ticks >> 62     # the 2 most significant bits store the "kind"
        ticks &= (1 << 62) - 1  # all but the above
        if ticks >= 1 << 61:    # if negative, reinterpret
            ticks -= 1 << 62    # as 62-bit two's complement
        time = datetime(1, 1, 1)
        with suppress(OverflowError):
            time += timedelta(microseconds= ticks / 10)
        if kind == 1:
            time = time.replace(tzinfo=timezone.utc)
        elif kind == 2:
            with suppress(OSError):
                time = time.astimezone()  # kind 2 is the local time zone
        return time

    @_register_reader(_PrimitiveType_readers, 18)
    def _read_LengthPrefixedString(self):
        length = 0
        for bit_range in range(0, 5*7, 7):  # each bit range is 7 bits long, with at most 5 of them
            byte    = self._read_Byte()
            length += (byte & ((1 << 7) - 1)) << bit_range  # highest bit masked out, then shifted
            if byte & (1 << 7) == 0:        # the "is there more?" (highest/most significant) bit
                break
        else:
            raise OverflowError('LengthPrefixedString overflow')
        return self._file.read(length).decode('utf-8')

    @_register_reader(_PrimitiveType_readers, 5)
    def _read_Decimal(self):
        return Decimal(self._read_LengthPrefixedString())

    @_register_reader(_PrimitiveType_readers, 17)
    def _read_Null(self):
        return None

    # Creates the rest of the PrimitiveType readers (those based on struct.unpack);
    # this only works when it's called after the class has been fully defined.
    _array_format_by_primitive_type  = {}  # array-format-specs  indexed by PrimitiveTypeEnumeration length-one bytes
    _struct_format_by_primitive_type = {}  # struct-format-specs indexed by PrimitiveTypeEnumeration length-one bytes
    @classmethod
    def _create_PrimitiveType_readers(cls):
        for enum_value, name, format in (
                ( 1, '_read_Boolean',  '?'),
                ( 2, '_read_Byte',     'B'),
                ( 6, '_read_Double',  '<d'),
                ( 7, '_read_Int16',   '<h'),
                ( 8, '_read_Int32',   '<l'),
                ( 9, '_read_Int64',   '<q'),
                (10, '_read_SByte',    'b'),
                (11, '_read_Single',  '<f'),
                (14, '_read_UInt16',  '<H'),
                (15, '_read_UInt32',  '<L'),
                (16, '_read_UInt64',  '<Q')):
            enum_value = bytes([enum_value])  # convert from int to length-one bytes
            if format[-1] in typecodes:  # if the format can also be used to create an array type
                cls._array_format_by_primitive_type[enum_value] = format[-1]  # (the [-1] skips the '<' flag)
            cls._struct_format_by_primitive_type[enum_value] = format
            struct = Struct(format)
            length = calcsize(format)
            reader = lambda self, s=struct, l=length: s.unpack(self._file.read(l))[0]
            setattr(cls, name, _register_reader(cls._PrimitiveType_readers, enum_value)(reader))

    # The ClassTypeInfo structure is read and ignored
    def _read_ClassTypeInfo(self):
        self._read_LengthPrefixedString()  # TypeName
        self._read_Int32()                 # LibraryId

    # @enum.unique
    # class _MessageFlags(enum.Flag):
    #     NoArgs                 = 0x00000001
    #     ArgsInline             = 0x00000002
    #     ArgsIsArray            = 0x00000004
    #     ArgsInArray            = 0x00000008
    #     NoContext              = 0x00000010
    #     ContextInline          = 0x00000020
    #     ContextInArray         = 0x00000040
    #     MethodSignatureInArray = 0x00000080
    #     PropertiesInArray      = 0x00000100
    #     NoReturnValue          = 0x00000200
    #     ReturnValueVoid        = 0x00000400
    #     ReturnValueInline      = 0x00000800
    #     ReturnValueInArray     = 0x00001000
    #     ExceptionInArray       = 0x00002000
    #     GenericMethod          = 0x00008000

    def _read_ValueWithCode(self):
        return self._PrimitiveType_readers[self._file.read(1)](self)

    _read_StringValueWithCode = _read_ValueWithCode

    def _read_ArrayOfValueWithCode(self):
        return [self._read_ValueWithCode() for i in range(self._read_Int32())]


    # A dict of all RecordType readers indexed by a length-one RecordTypeEnumeration bytes object
    _RecordType_readers = {}

    @_register_reader(_RecordType_readers, 21)
    def _read_BinaryMethodCall(self):
        raise NotImplementedError('BinaryMethodCall')
        # flags = self._MessageFlags(self._read_Int32())  # MessageEnum
        # self._read_StringValueWithCode()                # MethodName
        # self._read_StringValueWithCode()                # TypeName
        # if flags & self._MessageFlags.ContextInline:
        #     self._read_StringValueWithCode()            # CallContext
        # if flags & self._MessageFlags.ArgsInline:
        #     self._read_ArrayOfValueWithCode()           # Args

    @_register_reader(_RecordType_readers, 22)
    def _read_BinaryMethodReturn(self):
        raise NotImplementedError('BinaryMethodReturn')
        # flags = self._MessageFlags(self._read_Int32())  # MessageEnum
        # if flags & self._MessageFlags.ReturnValueInline:
        #     self._read_ValueWithCode()                  # ReturnValue
        # if flags & self._MessageFlags.ContextInline:
        #     self._read_StringValueWithCode()            # CallContext
        # if flags & self._MessageFlags.ArgsInline:
        #     self._read_ArrayOfValueWithCode()           # Args


    ######## Classes ########

    # Reads a ClassInfo structure, creates a new Python class with the members specified by the
    # ClassInfo, adds it to self._Class_by_id indexed by the ObjectId, and returns the class and id.
    def _read_ClassInfo(self):
        object_id    = self._read_Int32()
        class_name   = self._read_LengthPrefixedString()
        member_count = self._read_Int32()
        member_names = [self._read_LengthPrefixedString() for i in range(member_count)]
        unique_members = set()
        for member_num, member_name in enumerate(member_names):
            member_name = make_unique(sanitize_identifier(member_name), unique_members)
            unique_members.add(member_name)
            member_names[member_num] = member_name
        Class = namedlist(sanitize_identifier(class_name), member_names, default=None)
        Class._primitive_types = {}  # filled in below by _read_MemberTypeInfo()
        # Check to see if there is a converter method which can convert this type from a .NET
        # Collection (e.g. an ArrayList or Generic.List) to a native python type, and store it.
        if class_name.startswith('System.Collections.'):
            short_name = class_name[19:].split('`', 1)[0]  # strip 'System.Collections.' and any type params after backtick
            short_name = short_name.replace('.', '_').lower()  # e.g. 'arraylist' or 'generic_list'
            converter = getattr(self, f'_convert_{short_name}', None)
            if callable(converter):
                Class._convert_collection = converter
        self._Class_by_id[object_id] = Class
        return Class, object_id

    # Readers for the AdditionalInfos member of MemberTypeInfo indexed by
    # BinaryTypeEnumeration ints; created after this class is fully defined
    _AdditionalInfo_readers = ()

    def _read_MemberTypeInfo(self, Class):
        binary_types = [self._read_Byte() for m in Class._fields]  # BinaryTypeEnums
        for member_num, binary_type in enumerate(binary_types):    # AdditionalInfos
            additional_info = self._AdditionalInfo_readers[binary_type](self)
            if binary_type == 0:  # (0 == BinaryTypeEnumeration.Primitive)
                Class._primitive_types[member_num] = additional_info  # save any PrimitiveType for later

    @_register_reader(_RecordType_readers, 5)
    def _read_ClassWithMembersAndTypes(self):
        Class, object_id = self._read_ClassInfo()
        self._read_MemberTypeInfo(Class)
        self._read_Int32()  # LibraryId is ignored
        return self._read_members_into(Class(), object_id, Class._primitive_types.get)

    @_register_reader(_RecordType_readers, 3)
    def _read_ClassWithMembers(self):
        Class, object_id = self._read_ClassInfo()
        self._read_Int32()  # LibraryId is ignored
        return self._read_members_into(Class(), object_id, Class._primitive_types.get)

    @_register_reader(_RecordType_readers, 4)
    def _read_SystemClassWithMembersAndTypes(self):
        Class, object_id = self._read_ClassInfo()
        self._read_MemberTypeInfo(Class)
        return self._read_members_into(Class(), object_id, Class._primitive_types.get)

    @_register_reader(_RecordType_readers, 2)
    def _read_SystemClassWithMembers(self):
        Class, object_id = self._read_ClassInfo()
        return self._read_members_into(Class(), object_id, Class._primitive_types.get)

    @_register_reader(_RecordType_readers, 1)
    def _read_ClassWithId(self):
        object_id = self._read_Int32()
        Class = self._Class_by_id[self._read_Int32()]  # MetadataId
        return self._read_members_into(Class(), object_id, Class._primitive_types.get)

    # If primitive_type is not None, read the specified primitive_type with one of the
    # PrimitiveType_readers. Otherwise read the next RecordType in the streamfile with one
    # of the RecordType_readers. If overwrite_infos is not None, adds an _OverwriteInfo
    # object to overwrite_infos[overwrite_index] for each value read that is an overwritable
    # primitive. Finally, returns the value read.
    def _read_Record_or_Primitive(self, primitive_type, overwrite_infos = None, overwrite_index = None):
        if primitive_type:
            if overwrite_infos:
                format = self._struct_format_by_primitive_type.get(primitive_type)
                if format:
                    overwrite_infos[overwrite_index] = self._OverwriteInfo(self._file.tell(), format)
            return self._PrimitiveType_readers[primitive_type](self)
        else:
            record_type = self._file.read(1)
            # If overwrite_infos is not None and record_type == MemberPrimitiveTyped, parse it ourselves--
            # read in the PrimitiveTypeEnum and call ourselves to finish parsing and add the _OverwriteInfo
            if overwrite_infos and record_type == b'\x08':
                return self._read_Record_or_Primitive(self._file.read(1), overwrite_infos, overwrite_index)
            return self._RecordType_readers[record_type](self)

    # Represents an NRBF MemberReference, see _read_MemberReference()
    _MemberReference = namedlist('_MemberReference', 'id parent index_in_parent resolved', default=None)

    # Reads list elements or members into the 'obj' pre-allocated list or namedlist instance
    def _read_members_into(self, obj, object_id, members_primitive_type):
        assert callable(members_primitive_type)  # when called with the member_num, returns any respective primitive type
        if self._add_overwrite_info:
            # create the object which will store the _OverwriteInfo objects for each overwritable member in 'obj'
            overwrite_infos = [None] * len(obj) if isinstance(obj, list) else obj.__class__()
        else:
            overwrite_infos = None
        member_num = 0
        while member_num < len(obj):
            primitive_type = members_primitive_type(member_num)
            val = self._read_Record_or_Primitive(primitive_type, overwrite_infos, member_num)
            if isinstance(val, self._BinaryLibrary):  # a BinaryLibrary can precede the actual member, it's ignored
                val = self._read_Record_or_Primitive(primitive_type, overwrite_infos, member_num)
            if isinstance(val, self._ObjectNullMultiple):  # represents one or more empty members
                member_num += val.count
                continue
            if isinstance(val, self._MemberReference):     # see _read_MemberReference()
                val.parent          = obj
                val.index_in_parent = member_num
            obj[member_num] = val
            member_num += 1
        if overwrite_infos and any(overwrite_infos):
            self._overwrite_infos_by_pyid[id(obj)] = overwrite_infos
        # Convertible collections are added to _objects_by_id after their conversion in read_stream()
        if hasattr(obj.__class__, '_convert_collection'):
            self._collection_infos.append(self._CollectionInfo(obj, object_id))
        else:
            self._objects_by_id[object_id] = obj
        return obj


    ######## Arrays ########

    @enum.unique
    class _BinaryArrayTypeEnumeration(enum.Enum):
        Single            = 0
        Jagged            = 1
        Rectangular       = 2
        SingleOffset      = 3
        JaggedOffset      = 4
        RectangularOffset = 5
        def is_offset(self):
            return self in (self.SingleOffset, self.JaggedOffset, self.RectangularOffset)
        def is_rectangular(self):
            return self in (self.Rectangular, self.RectangularOffset)

    def _read_ArrayInfo(self):
        return (self._read_Int32(),  # ObjectId
                self._read_Int32())  # Length

    # This might not be able to use _read_members_into() since BinaryArrays are much more complex than
    # other types, thus it has no choice but to reimplement much of what _read_members_into() does
    @_register_reader(_RecordType_readers, 7)
    def _read_BinaryArray(self):
        object_id  = self._read_Int32()
        array_type = self._BinaryArrayTypeEnumeration(self._read_Byte())
        rank       = self._read_Int32()
        lengths    = [self._read_Int32() for i in range(rank)]
        if array_type.is_offset():
            [self._read_Int32() for i in range(rank)]  # LowerBounds are not implemented; they're ignored
        binary_type     = self._read_Byte()
        additional_info = self._AdditionalInfo_readers[binary_type](self)
        primitive_type  = additional_info if binary_type == 0 else None  # (0 == BinaryTypeEnumeration.Primitive)
        # If the BinaryArray is multidimensional, the complex code branch is required:
        if array_type.is_rectangular():
            array_format = self._array_format_by_primitive_type.get(primitive_type) if primitive_type else None
            if array_format:  # if not None, a Python Array can be used for the *last* dimension...
                array_length = lengths.pop()  # ...which is removed from the lengths list here
            array = multidimensional_array(lengths)  # preallocate a list of lists (it's not a Python Array)
            skip  = 0
            for indexes in itertools.product(*[range(l) for l in lengths]):  # iterates through all of the indexes
                if array_format:
                    val = self._read_Array_native_elements(array_format, array_length)  # read in the last index in one call
                else:
                    if skip > 0:
                        skip -= 1
                        continue
                    val = self._read_Record_or_Primitive(primitive_type)
                    if isinstance(val, self._BinaryLibrary):   # a BinaryLibrary can precede the actual array element, it's ignored
                        val = self._read_Record_or_Primitive(primitive_type)
                    if isinstance(val, self._ObjectNullMultiple):  # represents one or more empty elements
                        skip = val.count - 1  # counts this iteration which we're skipping right now
                        continue
                    if isinstance(val, self._MemberReference):     # see _read_MemberReference()
                        list_indexes = ''.join(f'[{i}]' for i in indexes[:-1])
                        val.parent          = eval(f'array{list_indexes}')  # the parent list, i.e. all but the last index
                        val.index_in_parent = indexes[-1]                   # the last index
                list_indexes = ''.join(f'[{i}]' for i in indexes)
                exec(f'array{list_indexes} = val')
            self._objects_by_id[object_id] = array
            return array
        else:  # else it's not multidimensional, the standard code branch can be used:
            assert len(lengths) == 1
            return self._read_Array_elements(lengths[0], object_id, primitive_type)

    @_register_reader(_RecordType_readers, 16)
    def _read_ArraySingleObject(self):
        object_id, length = self._read_ArrayInfo()
        return self._read_Array_elements(length, object_id)

    @_register_reader(_RecordType_readers, 15)
    def _read_ArraySinglePrimitive(self):
        object_id, length = self._read_ArrayInfo()
        primitive_type    = self._file.read(1)
        return self._read_Array_elements(length, object_id, primitive_type)

    _read_ArraySingleString = _register_reader(_RecordType_readers, 17)(_read_ArraySingleObject)

    def _read_Array_elements(self, length, object_id, primitive_type = None):
        array_format = self._array_format_by_primitive_type.get(primitive_type) if primitive_type else None
        if array_format:  # if not None, a Python Array can be used
            array = self._read_Array_native_elements(array_format, length)
            self._objects_by_id[object_id] = array
            return array
        return self._read_members_into([None] * length, object_id, lambda i: primitive_type)

    def _read_Array_native_elements(self, format, length):
        if self._add_overwrite_info:
            initial_pos = self._file.tell()
        array_struct = Struct(f'<{length}{format}')
        array = Array(format, array_struct.unpack(self._file.read(array_struct.size)))
        if self._add_overwrite_info:
            format = '<' + format  # always little-endian
            self._overwrite_infos_by_pyid[id(array)] = [ self._OverwriteInfo(pos, format)
                for pos in range(initial_pos, initial_pos + array_struct.size, calcsize(format)) ]
        return array


    _read_MemberPrimitiveTyped = _register_reader(_RecordType_readers, 8)(_read_ValueWithCode)

    # Wherever an ObjectId is encountered above, it's added to the self._objects_by_id dict
    # along with the newly-created object. A MemberReference contains an IdRef which refers to
    # an object in this dict, however because MemberReferences can appear before the object to
    # which they refer, they can't be resolved until the very end. Instead, a _MemberReference
    # Python object is temporarily created which stores enough information (its parent and
    # its index inside the parent) to eventually replace it with the actual object referenced
    # once all of the objects have been read from the stream in read_stream().
    @_register_reader(_RecordType_readers, 9)
    def _read_MemberReference(self):
        member_ref = self._MemberReference(self._read_Int32())  # IdRef
        self._member_references.append(member_ref)  # (parent and index_in_parent should be set by the caller)
        return member_ref

    _read_ObjectNull = _register_reader(_RecordType_readers, 10)(_read_Null)

    _ObjectNullMultiple = namedtuple('_ObjectNullMultiple', 'count')

    @_register_reader(_RecordType_readers, 14)
    def _read_ObjectNullMultiple(self):
        return self._ObjectNullMultiple(self._read_Int32())

    @_register_reader(_RecordType_readers, 13)
    def _read_ObjectNullMultiple256(self):
        return self._ObjectNullMultiple(self._read_Byte())

    @_register_reader(_RecordType_readers, 6)
    def _read_BinaryObjectString(self):
        object_id = self._read_Int32()
        string    = self._read_LengthPrefixedString()
        self._objects_by_id[object_id] = string
        return string

    @_register_reader(_RecordType_readers, 0)
    def _read_SerializationHeaderRecord(self):
        self._root_id = self._read_Int32()
        self._read_Int32()  # HeaderId is ignored
        major_version = self._read_Int32()
        minor_version = self._read_Int32()
        if major_version != 1:
            raise NotImplementedError(f'SerializationHeaderRecord.MajorVersion == {major_version}')
        if minor_version != 0:
            raise NotImplementedError(f'SerializationHeaderRecord.MinorVersion == {minor_version}')
        if self._root_id == 0:
            raise NotImplementedError('SerializationHeaderRecord.RootId == 0')

    class _BinaryLibrary: pass
    @_register_reader(_RecordType_readers, 12)
    def _read_BinaryLibrary(self):
        self._read_Int32()                 # MinorVersion and
        self._read_LengthPrefixedString()  # LibraryName are ignored
        return self._BinaryLibrary()

    class _MessageEnd: pass
    @_register_reader(_RecordType_readers, 11)
    def _read_MessageEnd(self):
        return self._MessageEnd()


    def read_header(self):
        '''Reads just the SerializationHeaderRecord from the streamfile.
        It's not necessary to call this, however it may be called before read_stream() if desired.
        Otherwise, read_stream() will raise an exception if the SerializationHeaderRecord isn't found.

        :return: True if the streamfile is in a supported .NET Remoting Binary Format
        '''
        assert self._root_id is None, 'read_header() has not already been called'
        try:
            self._read_Record_or_Primitive(primitive_type=False)
        except Exception:
            return False
        return self._root_id is not None

    def read_stream(self):
        '''Read the streamfile in .NET Remoting Binary Format and extract its root object

        :return: the root object contained in the stream
        '''
        if self._root_id is None and not self.read_header():
            raise RuntimeError('SerializationHeaderRecord not found (probably not an NRBF file)')
        obj = None
        while not isinstance(obj, self._MessageEnd):
            obj = self._read_Record_or_Primitive(primitive_type=False)
        self._Class_by_id.clear()

        # Resolve MemberReferences, ignoring failures (refs to convertible collections can't yet be resolved)
        for reference in self._member_references:
            self._resolve_reference(reference, ignore_unresolvable=True)

        # Convert collections to native Python types, and add them to _objects_by_id so they can be referenced
        # (references to non-collections must already have been resolved as done just above)
        for collection_info in self._collection_infos:
            collection = collection_info.obj
            converted  = collection.__class__._convert_collection(collection)  # calls one of the converters below
            self._objects_by_id[collection_info.id] = converted
        self._collection_infos.clear()

        # Resolve all the remaining member references (formerly pointing to collections)
        for reference in self._member_references:
            if not reference.resolved:
                self._resolve_reference(reference)
        self._member_references.clear()

        obj = self._objects_by_id[self._root_id]
        self._objects_by_id.clear()
        self._root_id = None
        return obj

    # Convert a .NET Collections.Generic.HashSet into a Python set
    @staticmethod
    def _convert_generic_hashset(collection):
        converted = set()
        for item in collection.Elements:
            try:
                assert item not in converted
                converted.add(item)
            except (AssertionError, TypeError):  # not all .NET HashSets can be converted to Python sets;
                return collection                # if the conversion fails, just proceed w/the original object
        return converted

    # Do the work of _convert_hashtable() and _convert_generic_dictionary()
    # converting a .NET key-value Collection into a Python dict
    def _do_convertto_dict(self, collection, collection_iter):
        converted       = {}
        overwrite_infos = {} if self._add_overwrite_info else None
        added_overwrite = False
        for key, value, values_overwrite_info in collection_iter(collection):
            try:
                assert key not in converted
                converted[key] = value
                if values_overwrite_info:
                    overwrite_infos[key] = values_overwrite_info
                    added_overwrite = True
            except (AssertionError, TypeError):  # not all .NET key-value Collections can be converted to Python dicts;
                return collection                # if the conversion fails, just proceed w/the original object
        # If any dict value is a _MemberReference, fix its parent and index_in_parent
        for key, value in converted.items():
            if isinstance(value, self._MemberReference):
                value.parent          = converted
                value.index_in_parent = key
        if added_overwrite:
            self._overwrite_infos_by_pyid[id(converted)] = overwrite_infos
        return converted

    # Convert a .NET Collections.HashTable into a Python dict
    def _convert_hashtable(self, collection):
        return self._do_convertto_dict(collection, self._hashtable_iter)
    #
    # Iterate over each key, value, and value's _OverwriteInfo in a .NET Collections.HashTable
    def _hashtable_iter(self, collection):
        values     = collection.Values
        len_values = len(values)
        overwrite_infos = self._overwrite_infos_by_pyid.get(id(values)) if self._add_overwrite_info else None
        for i, key in enumerate(collection.Keys):
            if i < len_values:
                yield key, values[i], overwrite_infos[i] if overwrite_infos else None
            else:
                yield key, None, None

    # Convert a .NET Collections.Generic.Dictionary into a Python dict
    def _convert_generic_dictionary(self, collection):
        return self._do_convertto_dict(collection, self._generic_dictionary_iter)
    #
    # Iterate over each key, value, and _OverwriteInfo in a .NET Collections.Generic.Dictionary
    def _generic_dictionary_iter(self, collection):
        for item in collection.KeyValuePairs:
            overwrite_infos = self._overwrite_infos_by_pyid.get(id(item)) if self._add_overwrite_info else None
            yield item.key, item.value, overwrite_infos.value if overwrite_infos else None

    # Convert a .NET Collections.ArrayList or .Generic.List into a Python list
    def _do_convertto_list(self, collection):
        converted = collection.items[:collection.size]
        if self._add_overwrite_info:
            overwrite_infos = self._overwrite_infos_by_pyid.get(id(collection.items))
            if overwrite_infos:
                self._overwrite_infos_by_pyid[id(converted)] = overwrite_infos
        # If any list element is a _MemberReference, fix its parent
        for element in converted:
            if isinstance(element, self._MemberReference):
                element.parent = converted  # (the index_in_parent remains the same)
        return converted
    #
    _convert_arraylist    = _do_convertto_list
    _convert_generic_list = _do_convertto_list

    # Replace a _MemberReference with its final referenced object, optionally ignoring
    # any unresolvable reference (which should point to a not-yet-converted collection)
    def _resolve_reference(self, reference, ignore_unresolvable = False):
        replacement = self._objects_by_id.get(reference.id)
        if replacement is not None:
            reference.parent[reference.index_in_parent] = replacement
            reference.resolved = True
        elif not ignore_unresolvable:
            raise RuntimeError(f'Unresolvable MemberReference with ObjectId {reference.id}')


    def overwrite_member(self, obj, member, value):
        '''Overwrite an object's member with a new value in the streamfile (does not change obj).

        :param obj: an object in the hierarchy returned by read_stream()
        :param member: a writable member (attribute) name, index (for lists/arrays), or key (for dicts) in obj
        :param value: the new value to be written to the streamfile
        '''
        assert self._add_overwrite_info, 'serialization object must have been constructed with can_overwrite_member == True'
        overwrite_infos = self._overwrite_infos_by_pyid[id(obj)]
        if isinstance(member, int) or isinstance(obj, dict):
            overwrite_info = overwrite_infos[member]
        else:
            overwrite_info = getattr(overwrite_infos, member)
        value   = pack(overwrite_info.format, value)
        old_pos = self._file.tell()
        try:
            self._file.seek(overwrite_info.pos)
            self._file.write(value)
        finally:
            self._file.seek(old_pos)

    def is_member_writable(self, obj, member):
        '''Returns True if the object's member can be overwritten.
        Can (but isn't guaranteed to) raise an exception if the member doesn't exist.

        :param obj: an object in the hierarchy returned by read_stream()
        :param member: a member (attribute) name, index (for lists/arrays), or key (for dicts) in obj
        '''
        assert self._add_overwrite_info, 'serialization object must have been constructed with can_overwrite_member == True'
        overwrite_infos = self._overwrite_infos_by_pyid.get(id(obj))
        if overwrite_infos is None:
            return False
        if isinstance(obj, dict):
            return overwrite_infos.get(member)  is not None
        if isinstance(member, int):
            return overwrite_infos[member]      is not None
        return getattr(overwrite_infos, member) is not None


# Finish setting up the serialization class
serialization._create_PrimitiveType_readers()
#
serialization._AdditionalInfo_readers = (
    lambda self: self._file.read(1),           # 0 Primitive
    serialization._read_Null,                  # 1 String
    serialization._read_Null,                  # 2 Object
    serialization._read_LengthPrefixedString,  # 3 SystemClass
    serialization._read_ClassTypeInfo,         # 4 Class
    serialization._read_Null,                  # 5 ObjectArray
    serialization._read_Null,                  # 6 StringArray
    lambda self: self._file.read(1),           # 7 PrimitiveArray
)

# Now that we're done adding PrimitiveType and RecordType readers, ensure they're
# all present (except PrimitiveType 4 and RecordTypes 18-20 which aren't defined)
assert all(bytes([i]) in serialization._PrimitiveType_readers if i != 4            else True for i in range(1, 19))
assert all(bytes([i]) in serialization._RecordType_readers    if not 18 <= i <= 20 else True for i in range(23))


def read_stream(streamfile):
    '''Read a file in .NET Remoting Binary Format and extract its root object

    :param streamfile: a file-like object
    :return: the root object contained in the stream
    '''
    return serialization(streamfile).read_stream()

# A JSONEncoder which can convert an object returned by read_stream() into json
# (can't handle circular references; primarily intended for debugging purposes)
class JSONEncoder(json.JSONEncoder):
    def default(self, o):
        if hasattr(o, '_asdict'):
            d = OrderedDict(_class_name=o.__class__.__name__)  # prepend the class name
            d.update(o._asdict())
            return d
        if isinstance(o, Array):
            return o.tolist()
        if isinstance(o, set):
            return list(o)
        if isinstance(o, (datetime, timedelta)):
            return str(o)
        if isinstance(o, Decimal):
            return repr(o)
        return super().default(o)


# Returns a version of the identifier suitable to pass to namedlist
_identifier_re = re.compile('[^a-z0-9_]', flags=re.IGNORECASE)
def sanitize_identifier(identifier):
    identifier = _identifier_re.sub('_', identifier).lstrip('0123456789_')
    if not identifier:
        return 'invalid_identifier'
    if iskeyword(identifier):
        identifier += '_'
    assert identifier.isidentifier()
    return identifier

# Returns a version of the name which isn't present in the unique_set
def make_unique(name, unique_set):
    if name not in unique_set:
        return name
    for append in itertools.count(2):
        replacement = name + str(append)
        if replacement not in unique_set:
            return replacement

# Pre-allocates a "multidimensional array", i.e. a list of lists of Nones
def multidimensional_array(lengths):
    if not lengths:
        return None
    arrays = [None] * lengths[-1]
    for length in reversed(lengths[:-1]):
        arrays = [deepcopy(arrays) for i in range(length)]
    return arrays


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit(f'Usage: {sys.argv[0]} streamfile')
    with open(sys.argv[1], 'rb') as streamfile:
        while True:
            json.dump(read_stream(streamfile), sys.stdout, cls=JSONEncoder, indent=4)
            print()
            if streamfile.peek(1) == b'':
                break
