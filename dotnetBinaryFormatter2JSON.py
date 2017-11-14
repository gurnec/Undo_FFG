# https://github.com/agix/NetBinaryFormatterParser/blob/996c20c9a4a51d1647ff6c3fec70ea37680a3dfd/dotnetBinaryFormatter2JSON.py
# Adapted from the link above to Python 3. All credit goes to the original author, GitHub user "agix".
# Added Unicode support and fixed 3 parsing bugs: ObjectNullMultiple records inside Arrays, records in BinaryArrays, & Nulls

import struct

theClass = {}

def date(s, options):
    d = popp(s, options[0])
    return d.hex()

def unpack_value(s, options):
    return struct.unpack(options[1], popp(s, options[0]))[0]

def Char(s, options=None):
    b = bytearray()
    while True:
        b += popp(s, 1)
        try:
            return b.decode('utf-8')
        except:
            pass

def LengthPrefixedString(s, options=None):
    Length = 0
    shift = 0
    c = True
    while c:
        byte = popp(s, 1)[0]
        if byte&128:
            byte^=128
        else:
            c = False
        Length += byte<<shift
        shift+=7
    return popp(s, Length).decode('utf-8')


PrimitiveTypeEnumeration = {
1:['Boolean',unpack_value, [1, '<b']],
2:['Byte', unpack_value, [1, '<B']],
3:['Char', Char, None],
5:['Decimal', LengthPrefixedString, None],
6:['Double', unpack_value, [8, '<d']],
7:['Int16', unpack_value, [2, '<h']],
8:['Int32', unpack_value, [4, '<i']],
9:['Int64', unpack_value, [8, '<q']],
10:['SByte', unpack_value, [1, '<b']],
11:['Single', unpack_value, [4, '<f']],
12:['TimeSpan', date, [8]],
13:['DateTime', date, [8]],
14:['UInt16', unpack_value, [2, '<H']],
15:['UInt32', unpack_value, [4, '<I']],
16:['UInt64', unpack_value, [8, '<Q']],
17:['Null', lambda s,o: None, None],
18:['String', LengthPrefixedString, None]
}

def parse_value(object, s):
    if object[0]=="Primitive":
        for p in PrimitiveTypeEnumeration:
            if PrimitiveTypeEnumeration[p][0] == object[1]:
                value = PrimitiveTypeEnumeration[p][1](s, PrimitiveTypeEnumeration[p][2])
    else:
        value = parse_object(s)
    return value

def parse_values(objectID, s):
    values = []
    myClass = theClass[objectID]
    a = 0
    for object in myClass[0]:
        values.append((myClass[1][a]+' : '+object[0], parse_value(object, s)))
        a+=1
    return values




BinaryArrayTypeEnumeration = {
0:['Single'],
1:['Jagged'],
2:['Rectangular'],
3:['SingleOffset'],
4:['JaggedOffset'],
5:['RectangularOffset']
}


def Primitive(s):
    return PrimitiveTypeEnumeration[popp(s, 1)[0]]

def SystemClass(s):
    return LengthPrefixedString(s)

def none(s):
    pass

def ClassTypeInfo(s):
    classTypeInfo = {}
    TypeName = LengthPrefixedString(s)
    LibraryId = struct.unpack('<I', popp(s, 4))[0]
    classTypeInfo['TypeName'] = TypeName
    classTypeInfo['LibraryId'] = LibraryId
    return classTypeInfo


BinaryTypeEnumeration = {
0:['Primitive', Primitive],
1:['String', none],
2:['Object', none],
3:['SystemClass', SystemClass],
4:['Class', ClassTypeInfo],
5:['ObjectArray', none],
6:['StringArray', none],
7:['PrimitiveArray', Primitive]
}


def SerializedStreamHeader(s):
    serializedStreamHeader = {}
    (RootId, HeaderId, MajorVersion, MinorVersion) = struct.unpack('<IIII', popp(s, 16))
    serializedStreamHeader['RootId'] = RootId
    serializedStreamHeader['HeaderId'] = HeaderId
    serializedStreamHeader['MajorVersion'] = MajorVersion
    serializedStreamHeader['MinorVersion'] = MinorVersion
    return serializedStreamHeader

def BinaryLibrary(s):
    binaryLibrary = {}
    LibraryId = struct.unpack('<I', popp(s, 4))[0]
    LibraryName = LengthPrefixedString(s)
    binaryLibrary['LibraryId'] = LibraryId
    binaryLibrary['LibraryName'] = LibraryName
    return binaryLibrary

def ClassInfo(s):
    classInfo = {}
    ObjectId = struct.unpack('<I', popp(s, 4))[0]
    Name = LengthPrefixedString(s)
    MemberCount = struct.unpack('<I', popp(s, 4))[0]
    MemberNames = []
    for i in range(MemberCount):
        MemberNames.append(LengthPrefixedString(s))
    classInfo['ObjectId'] = ObjectId
    classInfo['Name'] = Name
    classInfo['MemberCount'] = MemberCount
    classInfo['MemberNames'] = MemberNames
    return classInfo

def MemberTypeInfo(s, c):
    memberTypeInfo = {}
    BinaryTypeEnums = []
    binaryTypeEnums = []
    AdditionalInfos = []
    for i in range(c):
        binaryTypeEnum = BinaryTypeEnumeration[popp(s, 1)[0]]
        binaryTypeEnums.append(binaryTypeEnum)
        BinaryTypeEnums.append(binaryTypeEnum[0])
    for i in binaryTypeEnums:
        if i[0] == 'Primitive' or i[0] == 'PrimitiveArray':
            AdditionalInfos.append((i[0],i[1](s)[0]))
        else:
            AdditionalInfos.append((i[0],i[1](s)))
    memberTypeInfo['BinaryTypeEnums'] = BinaryTypeEnums
    memberTypeInfo['AdditionalInfos'] = AdditionalInfos
    return memberTypeInfo

def ClassWithMembersAndTypes(s):
    classWithMembersAndTypes = {}
    Members = ClassInfo(s)
    MemberCount = Members['MemberCount']
    MemberTypeI = MemberTypeInfo(s, MemberCount)
    LibraryId = struct.unpack('<I', popp(s, 4))[0]
    theClass[Members['ObjectId']] = (MemberTypeI['AdditionalInfos'], Members['MemberNames'])
    classWithMembersAndTypes['ClassInfo'] = Members
    classWithMembersAndTypes['MemberTypeInfo'] = MemberTypeI
    classWithMembersAndTypes['LibraryId'] = LibraryId
    classWithMembersAndTypes['Values'] = parse_values(Members['ObjectId'], s)
    return classWithMembersAndTypes

def BinaryObjectString(s):
    binaryObjectString = {}
    ObjectId = struct.unpack('<I', popp(s, 4))[0]
    Value = LengthPrefixedString(s)
    binaryObjectString['ObjectId'] = ObjectId
    binaryObjectString['Value'] = Value
    return binaryObjectString

def ClassWithId(s):
    classWithId = {}
    ObjectId = struct.unpack('<I', popp(s, 4))[0]
    MetadataId = struct.unpack('<I', popp(s, 4))[0]
    classWithId['ObjectId'] = ObjectId
    classWithId['MetadataId'] = MetadataId
    classWithId['Values'] = parse_values(MetadataId, s)
    return classWithId

def MemberReference(s):
    memberReference = {}
    IdRef = struct.unpack('<I', popp(s, 4))[0]
    memberReference['IdRef'] = IdRef
    return memberReference

def SystemClassWithMembersAndTypes(s):
    systemClassWithMembersAndTypes = {}
    Members = ClassInfo(s)
    MemberCount = Members['MemberCount']
    MemberTypeI = MemberTypeInfo(s, MemberCount)
    theClass[Members['ObjectId']] = (MemberTypeI['AdditionalInfos'], Members['MemberNames'])
    systemClassWithMembersAndTypes['ClassInfo'] = Members
    systemClassWithMembersAndTypes['MemberTypeInfo'] = MemberTypeI
    systemClassWithMembersAndTypes['Values'] = parse_values(Members['ObjectId'], s)
    return systemClassWithMembersAndTypes


def BinaryArray(s):
    binaryArray = {}
    ObjectId = struct.unpack('<I', popp(s, 4))[0]
    BinaryArrayTypeEnum = BinaryArrayTypeEnumeration[popp(s, 1)[0]][0]
    Rank = struct.unpack('<I', popp(s, 4))[0]
    Lengths = []
    total_length = 1
    LowerBounds = []
    for i in range(Rank):
        length = struct.unpack('<I', popp(s, 4))[0]
        Lengths.append(length)
        total_length *= length
    if 'Offset' in BinaryArrayTypeEnum:
        for i in range(Rank):
            LowerBounds.append(struct.unpack('<I', popp(s, 4))[0])
        binaryArray['LowerBounds'] = LowerBounds
    TypeEnum = BinaryTypeEnumeration[popp(s, 1)[0]]
    AdditionalTypeInfo = TypeEnum[1](s)
    binaryArray['ObjectId'] = ObjectId
    binaryArray['BinaryArrayTypeEnum'] = BinaryArrayTypeEnum
    binaryArray['Rank'] = Rank
    binaryArray['Lengths'] = Lengths
    binaryArray['TypeEnum'] = TypeEnum[0]
    binaryArray['AdditionalTypeInfo'] = AdditionalTypeInfo
    binaryArray['Values'] = []
    i = 0
    while i < total_length:
        value = parse_object(s)
        binaryArray['Values'].append(value)
        if value[0].startswith('ObjectNullMultiple'):
            i += value[1]['NullCount']
        else:
            i += 1
    return binaryArray

def ObjectNull(s):
    pass

def ObjectNullMultiple256(s):
    objectNullMultiple256 = {}
    NullCount = popp(s, 1)[0]
    objectNullMultiple256['NullCount'] = NullCount
    return objectNullMultiple256

def ObjectNullMultiple(s):
    objectNullMultiple = {}
    NullCount = struct.unpack('<I', popp(s, 4))[0]
    objectNullMultiple['NullCount'] = NullCount
    return objectNullMultiple

def ClassWithMembers(s):
    classWithMembers = {}
    Members = ClassInfo(s)
    LibraryId = struct.unpack('<I', popp(s, 4))[0]
    classWithMembers['ClassInfo'] = Members
    classWithMembers['LibraryId'] = LibraryId
    return classWithMembers

def SystemClassWithMembers(s):
    systemClassWithMembers = {}
    Members = ClassInfo(s)
    systemClassWithMembers['ClassInfo'] = Members
    return systemClassWithMembers

def MemberPrimitiveTyped(s):
    memberPrimitiveTyped = {}
    primitive = Primitive(s)
    value = primitive[1](s, primitive[2])
    memberPrimitiveTyped['PrimitiveTypeEnum'] = primitive[0]
    memberPrimitiveTyped['Value'] = value
    return memberPrimitiveTyped

def ArraySingleObject(s):
    arraySingleObject = {}
    ObjectId = struct.unpack('<I', popp(s, 4))[0]
    Length = struct.unpack('<I', popp(s, 4))[0]
    arraySingleObject['ObjectId'] = ObjectId
    arraySingleObject['Length'] = Length
    arraySingleObject['Values'] = []
    i = 0
    while i < Length:
        value = parse_object(s)
        arraySingleObject['Values'].append(value)
        if value[0].startswith('ObjectNullMultiple'):
            i += value[1]['NullCount']
        else:
            i += 1
    return arraySingleObject

def ArraySinglePrimitive(s):
    arraySinglePrimitive = {}
    ObjectId = struct.unpack('<I', popp(s, 4))[0]
    Length = struct.unpack('<I', popp(s, 4))[0]
    primitive = Primitive(s)
    arraySinglePrimitive['ObjectId'] = ObjectId
    arraySinglePrimitive['Length'] = Length
    arraySinglePrimitive['PrimitiveTypeEnum'] = primitive[0]
    arraySinglePrimitive['Values'] = []
    if primitive[0] != 'Byte':
        for o in range(Length):
            value = primitive[1](s, primitive[2])
            arraySinglePrimitive['Values'].append(value)
    else:
        arraySinglePrimitive['Values'] = popp(s, Length)
    return arraySinglePrimitive

def ArraySingleString(s):
    arraySingleString = {}
    ObjectId = struct.unpack('<I', popp(s, 4))[0]
    Length = struct.unpack('<I', popp(s, 4))[0]
    arraySingleString['ObjectId'] = ObjectId
    arraySingleString['Length'] = Length
    arraySingleString['Values'] = []
    i = 0
    while i < Length:
        value = parse_object(s)
        arraySingleString['Values'].append(value)
        if value[0].startswith('ObjectNullMultiple'):
            i += value[1]['NullCount']
        else:
            i += 1
    return arraySingleString

def MethodCall(s):
    methodCall = {}
    MessageEnum = struct.unpack('<I', popp(s, 4))[0]
    MethodName = StringValueWithCode(s)
    TypeName = StringValueWithCode(s)
    methodCall['MessageEnum'] = MessageEnum
    methodCall['MethodName'] = MethodName
    methodCall['TypeName'] = TypeName

    if MessageEnum & MessageFlagsEnum['NoContext'] == 0:
        CallContext = StringValueWithCode(s)
        methodCall['CallContext'] = CallContext

    if MessageEnum & MessageFlagsEnum['NoArgs'] == 0:
        Args = ArrayOfValueWithCode(s)
        methodCall['Args'] = Args

    return methodCall

def ArrayOfValueWithCode(s):
    arrayOfValueWithCode = {}
    arrayOfValueWithCode['Length'] = struct.unpack('<I', popp(s, 4))[0]
    arrayOfValueWithCode['ListOfValueWithCode'] = []

    for v in range(arrayOfValueWithCode['Length']):
        value = {}
        PrimitiveEnum = popp(s, 1)[0]
        value['PrimitiveTypeEnum'] = PrimitiveTypeEnumeration[PrimitiveEnum][0]
        value['Value'] = PrimitiveTypeEnumeration[PrimitiveEnum][1](s, PrimitiveTypeEnumeration[PrimitiveEnum][2])
        arrayOfValueWithCode['ListOfValueWithCode'].append(value)
    return arrayOfValueWithCode


def StringValueWithCode(s):
    popp(s, 1)
    return LengthPrefixedString(s)

MessageFlagsEnum = {
    'NoArgs': 0x00000001,
    'ArgsInline': 0x00000002,
    'ArgsIsArray': 0x00000004,
    'ArgsInArray': 0x00000008,
    'NoContext': 0x00000010,
    'ContextInline': 0x00000020,
    'ContextInArray': 0x00000040,
    'MethodSignatureInArray': 0x00000080,
    'PropertiesInArray': 0x00000100,
    'NoReturnValue': 0x00000200,
    'ReturnValueVoid': 0x00000400,
    'ReturnValueInline': 0x00000800,
    'ReturnValueInArray': 0x00001000,
    'ExceptionInArray': 0x00002000,
    'GenericMethod': 0x00008000
}

RecordTypeEnum = {
0:['SerializedStreamHeader', SerializedStreamHeader],
1:['ClassWithId', ClassWithId],
2:['SystemClassWithMembers', SystemClassWithMembers],
3:['ClassWithMembers', ClassWithMembers],
4:['SystemClassWithMembersAndTypes', SystemClassWithMembersAndTypes],
5:['ClassWithMembersAndTypes', ClassWithMembersAndTypes],
6:['BinaryObjectString', BinaryObjectString],
7:['BinaryArray', BinaryArray],
8:['MemberPrimitiveTyped', MemberPrimitiveTyped],
9:['MemberReference', MemberReference],
10:['ObjectNull', ObjectNull],
11:['MessageEnd', none],
12:['BinaryLibrary', BinaryLibrary],
13:['ObjectNullMultiple256', ObjectNullMultiple256],
14:['ObjectNullMultiple', ObjectNullMultiple],
15:['ArraySinglePrimitive', ArraySinglePrimitive],
16:['ArraySingleObject', ArraySingleObject],
17:['ArraySingleString', ArraySingleString],
20:['ArrayOfType', ArraySingleString],
21:['MethodCall', MethodCall],
22:['MethodReturn']
}


def popp(s, n):
    a = bytearray(s[:n])
    del s[:n]
    return a

def parse_object(s):
    RecordType = popp(s, 1)[0]
    return (RecordTypeEnum[RecordType][0], RecordTypeEnum[RecordType][1](s))

def parse_objects(stream):
    global theClass
    theClass = {}
    stream = bytearray(stream)
    myObject = []
    while(len(stream)!=0):
        a = parse_object(stream)
        myObject.append(a)
        if a[0] == 'MessageEnd':
            break
    return myObject
