#!python3.6
# Undo_MoM2e.py - Undo for Fantasy Flight Games apps
# Copyright (C) 2017 Christopher Gurnee
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


import threading, io, hashlib, collections, json, shutil, contextlib, \
       ctypes, ctypes.wintypes, sys, os, time, traceback
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, BadZipFile
from multiprocessing.connection import Listener, Client
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import nrbf

__version__ = '1.0'
DEFAULT_MAX_UNDO_STATES = 20

MOM = 'MoM2e'
RTL = 'RtL'
DEFAULT_GAME = MOM


# Return the binary hash of the files inside a MoM SaveGame directory
EMPTY_BINHASH = hashlib.md5().digest()
def dir_binhash(directory):
    hash = hashlib.md5()
    for f in directory.iterdir():
        if not f.name.lower().startswith('log') and f.is_file():  # exclude Log files
            hash.update(f.read_bytes())
    return hash.digest()

# Helper functions for validating Windows API return values
def returned_invalid_handle(result, func, arguments):
    if result == 0xFFFF_FFFF or result == 0xFFFF_FFFF_FFFF_FFFF:
        raise ctypes.WinError()
    return result
def returned_false(result, func, arguments):
    if result == 0:
        raise ctypes.WinError()
    return result

# Watch the specified directories for any file changes, and call the callback once they've been completed.
# If a specified directory doesn't exist but its *immediate* parent does, the callback is called if the
# directory is later created, but only after one or more files are also created inside it. This works
# works correctly for multiple such nonexistent directories if and only if they share the same parent.
watcher_skip_next = False  # if True and the next change takes 0.5s or less to finish, it is skipped
def watch_directory(directories, callback):
    handles = FindCloseChangeNotification = None  # used in the finally suite below
    try:
        if not hasattr(directories, '__len__'):
            directories = directories,
        assert callable(callback)

        # Load the Windows API functions and constants we need
        t = ctypes.wintypes
        FindFirstChangeNotification = ctypes.windll.kernel32.FindFirstChangeNotificationW
        FindFirstChangeNotification.argtypes = t.LPCWSTR, t.BOOL, t.DWORD
        FindFirstChangeNotification.restype  = t.HANDLE
        FindFirstChangeNotification.errcheck = returned_invalid_handle
        FALSE = t.BOOL(0)
        FILE_NOTIFY_CHANGE_FILE_NAME_or_LAST_WRITE = t.DWORD(0x0000_0001 | 0x0000_0010)
        FILE_NOTIFY_CHANGE_DIR_NAME                = t.DWORD(0x0000_0002)
        #
        FindCloseChangeNotification = ctypes.windll.kernel32.FindCloseChangeNotification
        FindCloseChangeNotification.argtypes = t.HANDLE,
        FindCloseChangeNotification.restype  = t.BOOL
        FindCloseChangeNotification.errcheck = returned_false
        #
        WaitForMultipleObjects = ctypes.windll.kernel32.WaitForMultipleObjects
        WaitForMultipleObjects.argtypes = t.DWORD, t.LPHANDLE, t.BOOL, t.DWORD
        WaitForMultipleObjects.restype  = t.DWORD
        WaitForMultipleObjects.errcheck = returned_invalid_handle
        INFINITE = t.DWORD(0xFFFF_FFFF)

        global watcher_skip_next
        # The handles list has one entry per potentially-monitored directory: one per specified directory
        # plus one at the end for a single parent directory. Its entries are non-None iff the handle is open.
        handles           = [None] * (len(directories) + 1)  # all closed initially
        parent_handle_num = len(handles) - 1                 # the index of the last one
        # Utility function to set or clear a handle in the handles list; returns 1 if set, 0 otherwise
        def set_handle(handle_num, dir, set, filter):
            if set:
                if not handles[handle_num]:
                    handles[handle_num] = FindFirstChangeNotification(str(dir), FALSE, filter)
                return 1
            else:
                if handles[handle_num]:
                    FindCloseChangeNotification(handles[handle_num])
                    handles[handle_num] = None
                return 0
        # ARRAY_TYPES[n] evaluates to the ctype "array of HANDLEs of length n" (i.e. the C type "HANDLE[n]")
        ARRAY_TYPES = [t.HANDLE * i for i in range(len(handles) + 1)]  # (ARRAY_TYPES[0] is unused)

        while True:
            # Figure out which directories we can/should monitor and build the fixed-sized
            # array of ChangeNotification HANDLES needed by WaitForMultipleObjects
            watch_parent = False
            array_len    = 0  # the length of the C array we'll need to construct
            for dir_num, dir in enumerate(directories):
                is_dir = dir.is_dir()
                array_len += set_handle(dir_num, dir, is_dir, FILE_NOTIFY_CHANGE_FILE_NAME_or_LAST_WRITE)
                if not is_dir and not watch_parent and dir.parent.is_dir():
                    watch_parent = dir.parent
            array_len += set_handle(parent_handle_num, watch_parent, watch_parent, FILE_NOTIFY_CHANGE_DIR_NAME)
            #
            handle_array        = ARRAY_TYPES[array_len]()  # creates the fixed-size array
            array_to_list_index = [None] * array_len  # will map indexes from handle_array to those in the handles list
            array_index         = 0
            for handle_num, handle in enumerate(handles):
                if handle is not None:
                    handle_array       [array_index] = handle
                    array_to_list_index[array_index] = handle_num
                    array_index += 1
            assert array_index == array_len, 'added all non-None handles to fixed-size handle_array'

            # Wait for the next change
            changed_array_index = WaitForMultipleObjects(len(handle_array), handle_array, FALSE, INFINITE)
            changed_handle_num  = array_to_list_index[changed_array_index]

            # Only the HANDLE of the directory that was modified is closed (we poll it below until
            # it's done changing). The others remain open and monitored and are reused later above.
            FindCloseChangeNotification(handles[changed_handle_num])
            handles[changed_handle_num] = None

            time.sleep(0.5)
            if watcher_skip_next:
                watcher_skip_next = False
                continue

            # For any new directories, wait in total for 1 second before looking for new files
            if changed_handle_num == parent_handle_num:
                time.sleep(0.5)
                changed_directories = []
                for dir_num, dir in enumerate(directories):
                    if handles[dir_num] is None and dir.is_dir():  # if it wasn't being watched before
                        for f in dir.iterdir():
                            changed_directories.append(dir)  # found at least one file,
                            break                            # continue to the next directory
            else:
                changed_directories = [directories[changed_handle_num]]

            for changed_directory in changed_directories:  # usually just one
                # Wait until files remain unchanged for a half-second stretch (but at least one second in total)
                last_binhash = dir_binhash(changed_directory)
                while True:
                    time.sleep(0.5)
                    cur_binhash = dir_binhash(changed_directory)
                    if cur_binhash == last_binhash:
                        break
                    last_binhash = cur_binhash
                callback(changed_directory, cur_binhash)

    except BaseException:
        callback(error=sys.exc_info())
        raise
    finally:
        if handles and FindCloseChangeNotification:
            for handle in handles:
                if handle is not None:
                    try:
                        FindCloseChangeNotification(handle)
                    except WindowsError:
                        pass


# Loads the CreateFile Windows API function and some constants we need
def load_CreateFile():
    global CreateFile, CloseHandle, GENERIC_READ_and_WRITE, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL
    t = ctypes.wintypes
    CreateFile = ctypes.windll.kernel32.CreateFileW
    CreateFile.argtypes = t.LPCWSTR, t.DWORD, t.DWORD, t.LPVOID, t.DWORD, t.DWORD, t.HANDLE
    CreateFile.restype  = t.HANDLE
    CreateFile.errcheck = returned_invalid_handle
    GENERIC_READ_and_WRITE = t.DWORD(0x8000_0000 | 0x4000_0000)
    OPEN_EXISTING          = t.DWORD(3)
    FILE_ATTRIBUTE_NORMAL  = t.DWORD(0x80)
    #
    CloseHandle = ctypes.windll.kernel32.CloseHandle
    CloseHandle.argtypes = t.HANDLE,
    CloseHandle.restype  = t.BOOL
    CloseHandle.errcheck = returned_false
load_CreateFile()
#
# Return True iff the specified file can be opened *exclusively* for read/write
# (exists and isn't opened by another process)
def can_open_exclusively(filepath):
    try:
        handle = CreateFile(
            str(filepath),           # lpFileName
            GENERIC_READ_and_WRITE,  # dwDesiredAccess
            0,                       # dwShareMode (0 == sharing not permitted)
            None,                    # lpSecurityAttributes
            OPEN_EXISTING,           # dwCreationDisposition
            FILE_ATTRIBUTE_NORMAL,   # dwFlagsAndAttributes
            None)                    # hTemplateFile
    except WindowsError as e:
        if e.winerror == 32:  # "The process cannot access the file because it is being used by another process."
            return False
        raise  # unexpected error
    CloseHandle(handle)
    return True


# Ignores the specified exception(s) inside a with statement, e.g.:
#   with ignored(AttributeError): obj.non_existant_attrib()  # won't raise
@contextlib.contextmanager
def ignored(*exceptions):
    try: yield
    except exceptions: pass

# Read the contents of a MoM GameData.dat file to retrieve the scenario name,
# player count, and round number (ignoring errors resulting from format changes)
def parse_mom_gamedata(savefile):
    savedata = nrbf.read_stream(savefile)
    scenario = players = round = ''
    # Remove the last word of the VariantName (I suspect it's the map variant):
    with ignored(AttributeError): scenario = savedata.VariantName[:savedata.VariantName.rfind(' ')]
    # InvestigatorIds is a comma-separated string; count its values:
    with ignored(AttributeError): players = savedata.InvestigatorIds.count(',') + 1
    with ignored(AttributeError): round   = savedata.Round
    return scenario, players, round

# Read the contents of a MoM_SaveGame file to retrieve the tile count, monster count,
# and highest-threat monster (ignoring errors resulting from format changes)
def parse_mom_savegame(savefile):
    savedata = nrbf.read_stream(savefile)
    tiles = 0
    try:
        for tile in savedata.TileSaveData.values():
            tiles += 1 if tile.Visible else 0
    except AttributeError:
        tiles = ''
    monsters = 0
    threat_name       = ''
    threat_is_unique  = False
    threat_max_damage = 0
    threat_health     = 0
    try:
        for node in savedata.NodeSaveData.values():
            if type(node).__name__ == 'FFG_MoM_MoM_SavedNodeMonster':
                cur_max_damage = node.MaxDamage
                cur_health     = cur_max_damage - node.DamageCount
                # Generated monsters are always visible, and of course so are
                # 'Visible' ones, but either is only counted if it's still alive.
                if (node.WasGenerated or node.Visible) and cur_health > 0:
                    monsters += 1
                    # Replace the current highest threat with this monster if it's greater;
                    # unique monsters are always greater than normal ones, otherwise higher
                    # toughness monsters are greater, otherwise choose the highest health one
                    cur_is_unique = node.MonsterName.upper().startswith('UNIQUE')
                    cur_is_higher = False
                    if cur_is_unique and not threat_is_unique:
                        cur_is_higher = True
                    elif cur_is_unique == threat_is_unique:
                        if cur_max_damage > threat_max_damage:
                            cur_is_higher = True
                        elif cur_max_damage == threat_max_damage:
                            if cur_health > threat_health:
                                cur_is_higher = True
                    if cur_is_higher:
                        threat_name       = node.MonsterName
                        threat_is_unique  = cur_is_unique
                        threat_max_damage = cur_max_damage
                        threat_health     = cur_health
        if threat_name:
            threat_name = threat_name.upper().replace('_', ' ')
            if threat_name.startswith('UNIQUE '):
                threat_name = threat_name[7:]
            if threat_name.startswith('MONSTER '):
                threat_name = threat_name[8:]
            threat_name = threat_name.title()
    except AttributeError:
        monsters = threat_name = ''
    return tiles, monsters, threat_name

# Read the contents of an RtL SavedGameA file to retrieve the group name, scenario, difficulty,
# player count, location and combat round (ignoring errors resulting from format changes).
RTL_SCENARIOS_BY_ID    = {'CAM_1':'Goblins', 'CAM_2':'Kindred Fire', 'CAM_3':'The Delve', 'CAM_4':'Nerekhall', 'CAM_5':'Frostgate'}
RTL_DIFFICULTIES_BY_ID = {0: 'Normal', 1: 'Hard'}
RTL_CITIES_BY_ID       = {'CITY_0':'Tamalir', 'CITY_1':'Nerekhall', 'CITY_2':'Greyhaven'}
def parse_rtl_savedgame(savefile):
    savedata = nrbf.read_stream(savefile)
    group = scenario = difficulty = players = location = round = ''
    with ignored(AttributeError): group      = savedata.PartyName
    with ignored(AttributeError): scenario   = RTL_SCENARIOS_BY_ID   .get(savedata.CampaignId,         '')
    with ignored(AttributeError): difficulty = RTL_DIFFICULTIES_BY_ID.get(savedata.CampaignDifficulty, '')
    with ignored(AttributeError): players    = len(savedata.HeroIds)
    location_id = None
    with ignored(AttributeError):
        for string_var in savedata.GlobalVarData.StringVars:
            if string_var.Name == "Campaign/CurrentLocation":
                location_id = string_var.Value
                break
    in_quest = True
    if location_id:
        location = RTL_CITIES_BY_ID.get(location_id)
        if location:  # if the location is a city
            in_quest = False
        else:         # else the location is a quest
            # If the quest has been completed, it's no longer active
            with ignored(AttributeError):
                if location_id in savedata.CampaignData.CompletedQuestIds:
                    location = ''
                    in_quest = False
            if in_quest:
                location = location_id.replace('_', ' ')
                if location.upper().startswith('QUEST '):
                    location = location[6:]
                location = location.title()
    if in_quest:
        with ignored(AttributeError):
            for int_var in savedata.GlobalVarData.IntVars:
                if int_var.Name == 'Round':
                    round = int_var.Value
                    break
    return group, scenario, difficulty, players, location, round


def init_gamespecific_globals(game):
    assert game in (MOM, RTL)
    global FFG_GAME
    FFG_GAME = game

    # known_undostate_hexhashes is a list containing one OrderedDict for each game
    # slave slot. They contain a hexhash for each "known" Undo State - a known
    # Undo State is persisted in the MYDATA_DIR and displayed in the treeview.
    global SLOT_COUNT, known_undostate_hexhashes
    if FFG_GAME == MOM:
        SLOT_COUNT = 0  # 0 is a flag for the special case where there's exactly one save slot
        known_undostate_hexhashes = [collections.OrderedDict()]
    elif FFG_GAME == RTL:
        SLOT_COUNT = 5
        known_undostate_hexhashes = [collections.OrderedDict() for i in range(SLOT_COUNT)]
    else: assert False
    assert len(str(SLOT_COUNT)) == 1  # a current code limitation: SLOT_COUNT must be < 10

    # Directory & filename constants
    global MYDATA_DIR, SETTINGS_FILENAME, STEAM_ID, SAVEGAME_DIR, LOG_FILENAMES, STEAMAPPS_DIR
    APPDATA_DIR = Path(os.environ['APPDATA'])
    assert APPDATA_DIR.is_dir(), 'located %APPDATA% directory'
    MYDATA_DIR        = APPDATA_DIR / 'Undo for MoM2e'
    SETTINGS_FILENAME = MYDATA_DIR  / f'{FFG_GAME}-settings.json'
    if FFG_GAME == MOM:
        STEAM_ID = 478980
        SAVEGAME_DIR  = APPDATA_DIR.parent / r'LocalLow\Fantasy Flight Games\Mansions of Madness Second Edition\SavedGame'
        LOG_FILENAMES = [SAVEGAME_DIR / 'Log']
    elif FFG_GAME == RTL:
        STEAM_ID = 477200
        # Look for the RtL install directory in the registry
        import winreg
        try:
            with winreg.OpenKeyEx(winreg.HKEY_LOCAL_MACHINE,
                    rf'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App {STEAM_ID}',
                    access=winreg.KEY_QUERY_VALUE | winreg.KEY_WOW64_64KEY) as regkey:
                SAVEGAME_DIR  = Path(winreg.QueryValueEx(regkey, 'InstallLocation')[0])
                STEAMAPPS_DIR = SAVEGAME_DIR.parent
                SAVEGAME_DIR /= r'Road to Legend_Data\SavedGames'
        # If the above fails or if it's not yet installed, fallback to the default install location
        except OSError:
            traceback.print_exc()
            STEAMAPPS_DIR  = Path(os.getenv('ProgramFiles(x86)') or os.environ['ProgramFiles'])
            assert STEAMAPPS_DIR.is_dir(), 'located %ProgramFiles% directory'
            STEAMAPPS_DIR /= r'Steam\SteamApps\common'
            SAVEGAME_DIR   = STEAMAPPS_DIR / r'Descent Road to Legend\Road to Legend_Data\SavedGames'
        LOG_FILENAMES = [SAVEGAME_DIR / rf'{slot}\LogA.txt' for slot in range(SLOT_COUNT)]
    else: assert False

    # Game-specific GUI strings
    global GAME_NAME_TEXT, OPEN_BUTON_TEXT
    if FFG_GAME == MOM:
        GAME_NAME_TEXT  = 'Mansions of Madness'
        OPEN_BUTON_TEXT = 'Open Mansions\nof Madness'
    elif FFG_GAME == RTL:
        GAME_NAME_TEXT  = 'Road to Legend'
        OPEN_BUTON_TEXT = 'Open Road\nto Legend'
    else: assert False

# Settings (all one of them)
settings = {}
MAX_UNDO_STATES = 'max_undo_states'

# A "binhash" is a bytes object containing the output from dir_binhash().
# A "hexhash" is string containing the first HEXHASH_LEN hex digits of the binhash;
# it is used in known_undostate_hexhashes, Undo State filenames, and treeview Item IDs.
HEXHASH_LEN        = 10
binhash_to_hexhash = lambda binhash: binhash.hex()[:HEXHASH_LEN]
EMPTY_HEXHASH      = binhash_to_hexhash(EMPTY_BINHASH)

CURRENT_ARROW = '\u2190Current'  # '<--Current'


# Override the exception handler for Tk events to cause the app to die (they're normally suppressed)
class UndoRoot(tk.Tk):
    def report_callback_exception(self, *error):
        msg = ''.join(traceback.format_exception(*error))
        print(msg, file=sys.stderr)
        messagebox.showerror('Exception', msg)
        self.destroy()

# The main application window
class UndoApplication(ttk.Frame):

    def __init__(self, master = None):
        super().__init__(master)
        self.master.title(f'Undo v{__version__} for {GAME_NAME_TEXT}')
        self.master.iconbitmap('Undo_MoM2e.ico')

        # Frame for treeview-related widgets
        frame = ttk.Frame()
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.TRUE, padx=12, pady=12)

        ttk.Label(frame, text='Undo States', underline=0).pack(padx=3, pady=3, anchor=tk.W)

        if FFG_GAME == MOM:
            col_headings = 'Scenario', 'Players', 'Round', 'Tiles', 'Monsters', 'Main Threat', 'Timestamp'
        elif FFG_GAME == RTL:
            col_headings = 'Group', 'Scenario', 'Difficulty', 'Players', 'Quest / Location', 'Round', 'Timestamp'
        else: assert False
        self.states_treeview = ttk.Treeview(frame,
            columns    = [c.lower() for c in col_headings] + ['current'],
            height     = min(settings[MAX_UNDO_STATES] + SLOT_COUNT, 40),
            selectmode = 'browse',                                      # only one item at a time may be selected
            show       = 'tree headings' if SLOT_COUNT else 'headings'  # only show the Slots column if required
        )
        if SLOT_COUNT:
            self.states_treeview.heading('#0', text='Slot')
            self.states_treeview.column ('#0', width=30)
        for col in col_headings:
            self.states_treeview.heading(col.lower(), text=col)
            if col != 'Timestamp':
                self.states_treeview.column(col.lower(), anchor=tk.CENTER, width=60)
        self.states_treeview.column('#1', anchor=tk.E, width=160)  # Scenario for MoM, Group for RtL
        if FFG_GAME == RTL:
            self.states_treeview.column('scenario', width=75)
            self.states_treeview.column('quest / location', width=180)
        if FFG_GAME == MOM:
            self.states_treeview.column('main threat', width=120)
        self.states_treeview.column('timestamp', width=120)
        self.states_treeview.column('current',   width=60)
        for slot in range(SLOT_COUNT):
            self.states_treeview.insert('', 'end', f'slot{slot}', text=slot+1, values=('\u2508'*100,))  # dotted line
        self.states_treeview.tag_configure('current_tag', background='yellow')
        self.states_treeview.bind('<<TreeviewSelect>>', self.handle_state_selected)
        self.states_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.TRUE)
        self.states_treeview.focus_set()
        self.master.bind('<Alt_L><u>', lambda e: self.states_treeview.focus_set())
        self.master.bind('<Alt_R><u>', lambda e: self.states_treeview.focus_set())

        # Don't know why, but if an item inside the treeview isn't given focus,
        # one can't use tab alone (w/o a mouse) to give focus to the treeview
        if SLOT_COUNT:
            self.states_treeview.focus('slot0')

        scrollbar = ttk.Scrollbar(frame, command=self.states_treeview.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.states_treeview.config(yscrollcommand=scrollbar.set)

        # Frame for Button widgets
        frame = ttk.Frame()
        frame.pack(side=tk.LEFT, fill=tk.Y, pady=30)

        ttk.Style().configure('TButton', justify='center')
        button_pady = 6

        # This is the only button whose command isn't a member function
        # (just to group the savegame-related code together down below)
        self.restore_button = ttk.Button(frame, text='Restore selected\nUndo State',
            command=handle_restore_clicked, underline=0, state=tk.DISABLED)
        self.restore_button.pack(fill=tk.X, pady=button_pady)
        self.master.bind('<Alt_L><r>', lambda e: self.restore_button.invoke())
        self.master.bind('<Alt_R><r>', lambda e: self.restore_button.invoke())

        self.save_as_button = ttk.Button(frame, text='Save selected\nUndo State as...',
            command=self.handle_save_as_clicked, underline=0, state=tk.DISABLED)
        self.save_as_button.pack(fill=tk.X, pady=button_pady)
        self.master.bind('<Alt_L><s>', lambda e: self.save_as_button.invoke())
        self.master.bind('<Alt_R><s>', lambda e: self.save_as_button.invoke())

        restore_from_button = ttk.Button(frame, text='Restore saved\nUndo State from...',
            command=self.handle_restore_from_clicked, underline=25)
        restore_from_button.pack(fill=tk.X, pady=button_pady)
        self.master.bind('<Alt_L><f>', lambda e: restore_from_button.invoke())
        self.master.bind('<Alt_R><f>', lambda e: restore_from_button.invoke())

        settings_button = ttk.Button(frame, text='Settings...',
            command=self.handle_settings_clicked, underline=2)
        settings_button.pack(side=tk.BOTTOM, fill=tk.X, pady=button_pady)
        self.master.bind('<Alt_L><t>', lambda e: settings_button.invoke())
        self.master.bind('<Alt_R><t>', lambda e: settings_button.invoke())

        open_game_button = ttk.Button(frame, text=OPEN_BUTON_TEXT,
            command=self.handle_open_game_clicked, underline=0)
        open_game_button.pack(side=tk.BOTTOM, fill=tk.X, pady=button_pady)  # gets placed *above* the settings button,
        open_game_button.lower(settings_button)                             # so move its tab-stop before settings too
        self.master.bind('<Alt_L><o>', lambda e: open_game_button.invoke())
        self.master.bind('<Alt_R><o>', lambda e: open_game_button.invoke())

        ttk.Sizegrip().pack(side=tk.BOTTOM)

    def handle_state_selected(self, event):
        selected = event.widget.selection()
        if selected:
            assert len(selected) == 1
            new_state = tk.DISABLED if selected[0].startswith('slot') else tk.NORMAL
        else:
            new_state = tk.DISABLED
        self.restore_button.config(state=new_state)
        self.save_as_button.config(state=new_state)

    FILEDIALOG_ARGS = None
    @classmethod
    def init_filedialog(cls):
        if not cls.FILEDIALOG_ARGS:
            cls.FILEDIALOG_ARGS = dict(
                filetypes        = ((f'{GAME_NAME_TEXT} Undo files', '*.undo'), ('All files', '*')),
                defaultextension = '.undo')

    @classmethod
    def handle_save_as_clicked(cls):
        selected = app.states_treeview.selection()
        assert selected and len(selected) == 1
        hexhash_slot = selected[0]
        scenario  = app.states_treeview.set(hexhash_slot, 'scenario')
        players   = app.states_treeview.set(hexhash_slot, 'players')
        round     = app.states_treeview.set(hexhash_slot, 'round')
        timestamp = app.states_treeview.set(hexhash_slot, 'timestamp')
        filename  = ''
        scenario_or_group = scenario
        if FFG_GAME == RTL:
            scenario_or_group = app.states_treeview.set(hexhash_slot, 'group')
            scenario_or_group = ''.join(c if c.isalnum() or c in " -'!" else '_' for c in scenario_or_group)  # sanitize
        if scenario_or_group:
            filename += scenario_or_group
            if players:
                filename += f' ({players}p)'
            filename += ', '
        elif players:
            filename += f'{players} players, '
        if FFG_GAME == RTL:
            if scenario:
                filename += scenario
                difficulty = app.states_treeview.set(hexhash_slot, 'difficulty')
                if difficulty:
                    filename += f' ({difficulty})'
                filename += ', '
            location = app.states_treeview.set(hexhash_slot, 'quest / location')
            if location:
                filename += f'{location}, '
        if round:
            filename += f'round {round}, '
        filename += timestamp[:-3].replace(':', '.')
        cls.init_filedialog()
        filename = filedialog.asksaveasfilename(title='Save Undo State as', initialfile=filename, **cls.FILEDIALOG_ARGS)
        if filename:
            glob_pattern = f'{FFG_GAME} ????-??-?? ??.??.?? {hexhash_slot}.zip'
            src_filename = next(MYDATA_DIR.glob(glob_pattern))  # next() gets the first (should be the only) filename
            shutil.copyfile(src_filename, filename)

    @classmethod
    def handle_restore_from_clicked(cls):
        if is_game_running():
            return
        cls.init_filedialog()
        filename = filedialog.askopenfilename(title='Restore Undo State from', **cls.FILEDIALOG_ARGS)
        if not filename or is_game_running():
            return

        # Ensure it's a valid Undo file
        try:
            with ZipFile(filename) as unzipper:
                for zipped_filename in unzipper.namelist():
                    zipped_filename = zipped_filename.lower()
                    # If at least one of the savegame files is present, assume it's valid
                    if FFG_GAME == MOM and zipped_filename in ('gamedata.dat', 'mom_savegame'):
                        break
                    elif FFG_GAME == RTL and zipped_filename.startswith('savedgame'):  # e.g. SavedGameA
                        break
                else:
                    raise BadZipFile("can't find any expected game save filename")
        except BadZipFile as e:
            messagebox.showerror('Error', f'This file is not a {GAME_NAME_TEXT} Undo file.\n({e})')
            return

        if SLOT_COUNT:
            slot = simpledialog.askinteger('Slot?', '\nWhich Save Slot would you like this\n'
                                                     f'Undo State to be restored into (1-{SLOT_COUNT}) ?\n',
                                           minvalue=1, maxvalue=SLOT_COUNT)
            if not slot or is_game_running():
                return
            slot -= 1  # (it's zero-based)
        else:
            slot = 0
        restore_undo_state(filename, slot, update_rtl_save_index= FFG_GAME==RTL)
        handle_new_savegame(slot)

    @staticmethod
    def handle_open_game_clicked():
        root.config(cursor='wait')
        os.startfile(f'steam://run/{STEAM_ID}')  # see https://developer.valvesoftware.com/wiki/Steam_browser_protocol
        root.config(cursor='')

    @staticmethod
    def handle_settings_clicked():
        new_max_undo_states = simpledialog.askinteger('Settings',
            '\nMaximum Undo States ' + ('(per Save Slot) :\n' if SLOT_COUNT else ':\n'),
            initialvalue=settings[MAX_UNDO_STATES], minvalue=2, maxvalue=1_000)
        if new_max_undo_states and new_max_undo_states != settings[MAX_UNDO_STATES]:
            to_be_deleted = 0
            for hexhashes in known_undostate_hexhashes:
                to_be_deleted += max(len(hexhashes) - new_max_undo_states, 0)
            if to_be_deleted:
                answered_yes = messagebox.askyesno('Delete Undo States?',
                    f'Decreasing the maximum Undo States below their current number will delete {"some" if to_be_deleted>1 else "one"} of them. '
                    f'Are you sure you want to continue and delete the {f"{to_be_deleted} oldest States" if to_be_deleted>1 else "oldest State"}?',
                    icon=messagebox.WARNING, default=messagebox.NO)
                if not answered_yes:
                    return
            settings[MAX_UNDO_STATES] = new_max_undo_states
            with SETTINGS_FILENAME.open('w') as settings_file:
                json.dump(settings, settings_file)
            trim_undo_states()


# Load the settings saved right above
def load_settings():
    global settings
    if SETTINGS_FILENAME.is_file():
        with SETTINGS_FILENAME.open() as settings_file:
            settings = json.load(settings_file)
    elif FFG_GAME == MOM:
        # Migrate an old MoM settings file to its new name
        legacy_settings_filename = MYDATA_DIR / 'settings.json'
        if legacy_settings_filename.is_file():
            legacy_settings_filename.rename(SETTINGS_FILENAME)
            load_settings()


# Load the preserved savegames from our data directory (done once at program start), returning a
# list of Bools indicating for each save slot if the current SaveGame has been seen & preserved
def load_undo_states():
    assert all(len(h) == 0 for h in known_undostate_hexhashes), 'load_undo_states() should only be called once'
    if SLOT_COUNT:
        assert all(len(app.states_treeview.get_children(f'slot{s}')) == 0 for s in range(SLOT_COUNT)), \
               'the states_treeview is empty (except for slot numbers)'
    else:
        assert len(app.states_treeview.get_children()) == 0, 'the states_treeview is empty'

    if FFG_GAME == MOM:
        # Migrate any old MoM Undo States to the new naming convention
        glob_pattern = f'????-??-?? ??.??.?? {"?"*HEXHASH_LEN}.zip'
        for zip_filename in MYDATA_DIR.glob(glob_pattern):
            zip_filename.rename(MYDATA_DIR / f'{MOM} {zip_filename.stem}-0.zip')

    cur_savegames_hexhash = []  # the hexhash of each current game save slot
    cur_savegames_found   = []  # for each game save slot, True if we've already preserved/found it
    for slot in range(len(known_undostate_hexhashes)):
        savegame_dir = SAVEGAME_DIR / str(slot) if SLOT_COUNT else SAVEGAME_DIR
        hexhash = binhash_to_hexhash(dir_binhash(savegame_dir)) if savegame_dir.is_dir() else EMPTY_HEXHASH
        cur_savegames_hexhash.append(hexhash)
        cur_savegames_found  .append(hexhash == EMPTY_HEXHASH)  # empty counts as "already preserved"

    latest_hexhash_slot = None  # the hexhash-slot id of the most recent game save
    glob_pattern  = f'{FFG_GAME} ????-??-?? ??.??.?? {"?"*HEXHASH_LEN}-?.zip'
    zip_filenames = list(MYDATA_DIR.glob(glob_pattern))
    zip_filenames.sort()  # sorts from oldest to newest
    for zip_filename in zip_filenames:
        zip_filestem  = zip_filename.stem  # (the .stem is the file name w/o the extension)
        slot          = int(zip_filestem[-1:])
        hexhash       = zip_filestem[-HEXHASH_LEN-2:-2]
        timestamp_str = zip_filestem[-HEXHASH_LEN-22:-HEXHASH_LEN-3].replace('.', ':')
        if hexhash in known_undostate_hexhashes[slot]:  # shouldn't happen often if ever
            continue
        known_undostate_hexhashes[slot][hexhash] = True

        # Parse the savegame files inside the zip for display purposes
        treeview_values = [None] * 3 if FFG_GAME == MOM else []  # for MoM preallocate the first 3
        with ZipFile(zip_filename) as zipper:
            for zipped_filename in zipper.namelist():
                if FFG_GAME == MOM:
                    if zipped_filename.lower() == 'gamedata.dat':
                        with zipper.open(zipped_filename) as savefile:
                            treeview_values[:3] = parse_mom_gamedata(savefile)
                    elif zipped_filename.lower() == 'mom_savegame':
                        with zipper.open(zipped_filename) as savefile:
                            treeview_values.extend(parse_mom_savegame(savefile))
                elif FFG_GAME == RTL:
                    if zipped_filename.lower() == 'savedgamea':
                        with zipper.open(zipped_filename) as savefile:
                            treeview_values = parse_rtl_savedgame(savefile)

        # Insert the Undo State into the treeview UI at the top
        hexhash_slot = f'{hexhash}-{slot}'
        app.states_treeview.insert(f'slot{slot}' if SLOT_COUNT else '', 0, hexhash_slot,
            values=(*treeview_values, timestamp_str, ''))
        if hexhash == cur_savegames_hexhash[slot]:
            app.states_treeview.set(hexhash_slot, 'current', CURRENT_ARROW)
            app.states_treeview.item(hexhash_slot, tags=('current_tag',))
            cur_savegames_found[slot] = True
            latest_hexhash_slot = hexhash_slot

    # The current focus/view is only modified if all current savegames were found;
    # otherwise it's presumed that handle_new_savegame() will do this later instead
    if all(cur_savegames_found):
        # Don't know why, but if an item inside the treeview isn't given focus,
        # one can't use tab alone (w/o a mouse) to give focus to the treeview
        if not SLOT_COUNT:  # (otherwise it's already been done in UndoApplication.init)
            if len(known_undostate_hexhashes[0]):        # (0 is the only slot since SLOT_COUNT==0)
                app.states_treeview.focus(hexhash_slot)  # the most recently added (top) item
        #
        # Ensure the most recent save game is visible (possibly expanding a slot tree)
        if latest_hexhash_slot:
            app.states_treeview.see(latest_hexhash_slot)

    # The number of items in the treeview should match the number of known Undo States
    if SLOT_COUNT:
        assert all(len(app.states_treeview.get_children(f'slot{s}')) == len(known_undostate_hexhashes[s])
                   for s in range(SLOT_COUNT)), 'known_undostate_hexhashes and states_treeview match in lengths'
    else:
        assert len(app.states_treeview.get_children()) == len(known_undostate_hexhashes[0]), \
               'known_undostate_hexhashes[0] and states_treeview match in length'
    return cur_savegames_found


# Callback which runs in the context of the directory watcher
# thread and is called when a new savegame is detected
watcher_thread_error = None
def send_new_savegame_event(directory = None, binhash = None, error = None):
    if error:
        global watcher_thread_error
        watcher_thread_error = error
        root.event_generate('<<watcher_error>>')
    else:
        assert binhash
        if SLOT_COUNT:
            slot = int(directory.name)
            assert 0 <= slot < SLOT_COUNT, 'directory name is a valid slot number'
        else:
            slot = 0
        if binhash_to_hexhash(binhash) not in known_undostate_hexhashes[slot]:
            root.event_generate(f'<<new_savegame_slot{slot}>>')

# Runs in the context of the main thread to handle watcher thread exceptions
def handle_watcher_error(event):
    global watcher_thread_error
    msg = ''.join(traceback.format_exception(*watcher_thread_error)) if watcher_thread_error \
          else 'An unknown directory watcher error occurred.'
    watcher_thread_error = None
    print(msg, file=sys.stderr)
    messagebox.showerror('Exception', msg)
    root.destroy()

# Runs in the context of the main thread to handle new savegames
# and also on start to preserve an existing savegame
settings[MAX_UNDO_STATES] = DEFAULT_MAX_UNDO_STATES  # can be overwritten later in load_settings()
def handle_new_savegame(slot, use_filetime = False):
    if SLOT_COUNT:
        assert 0 <= slot < SLOT_COUNT
        savegame_dir = SAVEGAME_DIR / str(slot)
    else:
        assert slot == 0
        savegame_dir = SAVEGAME_DIR
    # Iterate through the files in the SaveGame directory to:
    #   - zip them (in memory for now) into an "Undo State"
    #   - calculate their hexhash to name the Undo State
    #   - parse the savegame files for display purposes
    inmem_zip = io.BytesIO()
    zipper    = ZipFile(inmem_zip, 'w', compression=ZIP_DEFLATED)
    hash      = hashlib.md5()
    treeview_values = [None] * 3 if FFG_GAME == MOM else []  # for MoM preallocate the first 3
    main_savename   =  None  # may need the name of the main save file for later
    for f in savegame_dir.iterdir():
        if f.is_file():
            zipper.write(f, f.name)
            lower_name = f.name.lower()
            if lower_name.startswith('log'):  # same as in dir_binhash(), log files are excluded from hashes
                continue
            hash.update(f.read_bytes())
            if FFG_GAME == MOM:
                if lower_name == 'gamedata.dat':
                    with f.open('rb') as savefile:
                        treeview_values[:3] = parse_mom_gamedata(savefile)
                elif lower_name == 'mom_savegame':
                    main_savename = f
                    with f.open('rb') as savefile:
                        treeview_values.extend(parse_mom_savegame(savefile))
            elif FFG_GAME == RTL:
                if lower_name == 'savedgamea':
                    main_savename = f
                    with f.open('rb') as savefile:
                        treeview_values = parse_rtl_savedgame(savefile)
    zipper.close()
    binhash = hash.digest()

    slot_item_id = f'slot{slot}' if SLOT_COUNT else ''
    for id in app.states_treeview.tag_has('current_tag'):  # for all w/the current_tag:
        if SLOT_COUNT and app.states_treeview.parent(id) != slot_item_id:
            continue                                       # skip items in other slots; otherwise:
        app.states_treeview.set(id, 'current', '')         # erase its 'current' column
        app.states_treeview.item(id, tags=())              # and remove the highlighting tag

    if binhash == EMPTY_BINHASH:  # if there's no game save to preserve
        return
    hexhash = binhash_to_hexhash(binhash)

    hexhash_slot = f'{hexhash}-{slot}'
    if hexhash in known_undostate_hexhashes[slot]:  # can happen following handle_restore_from_clicked()
        app.states_treeview.set(hexhash_slot, 'current', CURRENT_ARROW)  # set the 'current' column
        app.states_treeview.item(hexhash_slot, tags=('current_tag',))    # and add the highlighting tag
    else:
        save_time = None
        if use_filetime and main_savename:
            with ignored(OSError): save_time = main_savename.stat().st_mtime
        save_time    = time.localtime(save_time)  # converts from file time to struct_time, or if None gets cur time
        timstamp_str = time.strftime('%Y-%m-%d %H:%M:%S', save_time)  # e.g. '2017-11-08 19:01:27'
        zip_filename = MYDATA_DIR / f'{FFG_GAME} {timstamp_str.replace(":", ".")} {hexhash_slot}.zip'
        zip_filename.write_bytes(inmem_zip.getbuffer())
        app.states_treeview.insert(slot_item_id, 0, hexhash_slot, tags=('current_tag',),
            values=(*treeview_values, timstamp_str, CURRENT_ARROW))
        known_undostate_hexhashes[slot][hexhash] = True

        # Don't know why, but if an item inside the treeview isn't given focus,
        # one can't use tab alone (w/o a mouse) to give focus to the treeview
        if not SLOT_COUNT:  # (otherwise it's already been done in UndoApplication.init)
            if len(known_undostate_hexhashes[0]) == 1:  # only need to do this once
                app.states_treeview.focus(hexhash_slot)

    # Ensure the new save game is visible (possibly expanding a slot tree)
    app.states_treeview.see(hexhash_slot)

    trim_undo_states()

# Delete Undo States if we're over the max
def trim_undo_states():
    assert settings[MAX_UNDO_STATES] > 0
    for slot, hexhashes in enumerate(known_undostate_hexhashes):
        while len(hexhashes) > settings[MAX_UNDO_STATES]:
            hexhash      = hexhashes.popitem(last=False)[0]     # removes and returns the oldest
            hexhash_slot = f'{hexhash}-{slot}'
            glob_pattern = f'{FFG_GAME} ????-??-?? ??.??.?? {hexhash_slot}.zip'
            for zip_filename in MYDATA_DIR.glob(glob_pattern):  # should be exactly one
                zip_filename.unlink()
            app.states_treeview.delete(hexhash_slot)


# If the game is being played, alert the user and return True, else return False
def is_game_running():
    for log_filename in LOG_FILENAMES:
        if log_filename.is_file() and not can_open_exclusively(log_filename):
            break
    else:
        return False
    messagebox.showwarning('Error', 'Please save your game and quit to the main\n'
                                    'menu in order to restore an Undo State.')
    return True

# Restore an Undo State into the MoM SaveGame directory and update the UI
def handle_restore_clicked():
    selected = app.states_treeview.selection()
    assert selected and len(selected) == 1
    if is_game_running():
        return
    app.states_treeview.selection_set()  # clear the selection

    hexhash_slot = selected[0]
    slot         = int(hexhash_slot[-1])
    glob_pattern = f'{FFG_GAME} ????-??-?? ??.??.?? {hexhash_slot}.zip'
    zip_filename = next(MYDATA_DIR.glob(glob_pattern))  # next() gets the first (should be the only) filename
    assert zip_filename
    restore_undo_state(zip_filename, slot)

    slot_item_id = f'slot{slot}'
    for id in app.states_treeview.tag_has('current_tag'):  # for all w/the current_tag:
        if SLOT_COUNT and app.states_treeview.parent(id) != slot_item_id:
            continue                                       # skip items in other slots; otherwise:
        app.states_treeview.set(id, 'current', '')         # erase its 'current' column
        app.states_treeview.item(id, tags=())              # and remove the highlighting tag
    app.states_treeview.set(hexhash_slot, 'current', CURRENT_ARROW)  # set the 'current' column
    app.states_treeview.item(hexhash_slot, tags=('current_tag',))    # and add the highlighting tag

def restore_undo_state(zip_filename, slot, update_rtl_save_index = False):
    if SLOT_COUNT:
        assert 0 <= slot < SLOT_COUNT
        savegame_dir = SAVEGAME_DIR / str(slot)
    else:
        assert slot == 0
        savegame_dir = SAVEGAME_DIR
    assert not update_rtl_save_index or FFG_GAME == RTL  # this arg is only supported for RTL
    global watcher_skip_next
    extracted_filenames = []
    try:
        with ZipFile(zip_filename) as unzipper:
            watcher_skip_next = True  # tell the directory watcher thread to skip the following changes
            if SLOT_COUNT:
                savegame_dir.mkdir(exist_ok=True)  # the slot may not exist yet
            for zipped_filename in unzipper.namelist():
                if zipped_filename == Path(zipped_filename).name:  # ensures we're unzipping to only the SaveGame dir
                    unzipper.extract(zipped_filename, savegame_dir)
                    extracted_filenames.append(zipped_filename)

        # For RtL, the SaveIndex (what the UI calls the save slot) is stored
        # inside the SavedGameA file, so if it's different it must be updated
        if update_rtl_save_index:
            savegame_filename = savegame_dir / 'SavedGameA'
            with savegame_filename.open('r+b') as savefile:
                save_ser = nrbf.serialization(savefile, can_overwrite_member=True)
                savedata = save_ser.read_stream()
                if slot != savedata.SaveIndex:
                    save_ser.overwrite_member(savedata, 'SaveIndex', slot)

    except Exception:
        # Undo any file extractions if there were any errors
        for filename in extracted_filenames:
            filename = savegame_dir / filename
            try:
                filename.unlink()
            except Exception:
                traceback.print_exc()
        raise


# Whenever the pipe receives a connection, restore the main window (see right below)
def restore_window_listener(pipe):
    while True:
        pipe.accept().close()
        root.deiconify()
        root.attributes('-topmost', 1)  # root.lift() doesn't work on modern
        root.attributes('-topmost', 0)  # versions of Windows, but this does

root = app = None
def main(game):
    global root, app
    init_gamespecific_globals(game)

    exclusive_pipe = None
    try:
        # Create a named pipe. When running a second instance of Undo_MoM2e, the pipe
        # creation will fail. If it does, instead connect to the already-existing
        # pipe which causes the first instance of Undo_MoM2e to restore its window.
        PIPE_NAME = rf'\\.\pipe\Undo_{FFG_GAME}'
        try:
            exclusive_pipe = Listener(PIPE_NAME)
        except PermissionError:
            try:
                Client(PIPE_NAME).close()
            except Exception:
                traceback.print_exc()
            sys.exit(f'{GAME_NAME_TEXT} is already running.')

        load_settings()
        root = UndoRoot()
        app  = UndoApplication(root)
        root.config(cursor='wait')
        root.update()

        if FFG_GAME == RTL and not STEAMAPPS_DIR.is_dir():
            raise RuntimeError(f"Undo can't find either the {GAME_NAME_TEXT} nor the SteamApps folder")
        if not SAVEGAME_DIR.is_dir():
            answered_yes = messagebox.askyesno("Can't find SaveGame folder",
               f"Undo can't find the {GAME_NAME_TEXT} SaveGame folder. This is usually because "
               f'a {GAME_NAME_TEXT} game has never been started before on this computer. '
                'Would you like to start one now? (If you choose No, Undo will exit.)',
                icon=messagebox.QUESTION, default=messagebox.YES)
            if answered_yes:
                app.handle_open_game_clicked()
                WaitForDirDialog(root, SAVEGAME_DIR)  # wait for the directory to be created
            if not SAVEGAME_DIR.is_dir():  # if it's still not there, then the user must have chosen to exit
                root.destroy()
                sys.exit(f"Can't find the {GAME_NAME_TEXT} SaveGame folder.")

        MYDATA_DIR.mkdir(exist_ok=True)

        cur_savegames_known = load_undo_states()
        for slot, is_known in enumerate(cur_savegames_known):
            if not is_known:
                handle_new_savegame(slot, use_filetime=True)

        for slot in range(len(known_undostate_hexhashes)):
            root.bind(f'<<new_savegame_slot{slot}>>', lambda e, s=slot: handle_new_savegame(s))
        root.bind('<<watcher_error>>', handle_watcher_error)
        savegame_dirs  = [SAVEGAME_DIR / str(slot) for slot in range(SLOT_COUNT)] if SLOT_COUNT else SAVEGAME_DIR
        watcher_thread = threading.Thread(target=watch_directory, args=(savegame_dirs, send_new_savegame_event), daemon=True)
        watcher_thread.start()

        restore_window_thread = threading.Thread(target=restore_window_listener, args=(exclusive_pipe,), daemon=True)
        restore_window_thread.start()

        root.config(cursor='')
        root.mainloop()

    finally:
        if exclusive_pipe:
            exclusive_pipe.close()

# A dialog box which waits for a directory to be created, or for the user to click Exit
class WaitForDirDialog(simpledialog.Dialog):
    def __init__(self, parent, directory):
        self.directory = directory
        super().__init__(parent, 'Waiting ...')
    def body(self, master):
        self.resizable(tk.FALSE, tk.FALSE)
        ttk.Label(self, text=f'Waiting for a {GAME_NAME_TEXT} game to start ...').pack(pady=8)
        progress_bar = ttk.Progressbar(self, length=350, mode='indeterminate')
        progress_bar.pack(padx=16)
        progress_bar.start()
        self.after(1000, self.check_dir)
    def buttonbox(self):
        exit_button = ttk.Button(self, text='Exit', command=self.cancel, underline=1)
        exit_button.pack(pady=8)
        self.bind('<Alt_L><x>', lambda e: exit_button.invoke())
        self.bind('<Alt_R><x>', lambda e: exit_button.invoke())
    def check_dir(self):
        if self.directory.is_dir():
            self.cancel()
        self.after(1000, self.check_dir)

if __name__ == '__main__':
    if len(sys.argv) > 1:
        if len(sys.argv) == 2 and sys.argv[1].startswith('--game='):
            game_arg = sys.argv[1][7:].lower()
            if game_arg == 'mom':
                game = MOM
            elif game_arg == 'rtl':
                game = RTL
            else:
                sys.exit(f'Unsupported game, must be one of: mom, rtl')
        else:
            sys.exit(f'Usage: {sys.argv[0]} [--game=mom|rtl]')
    else:
        game = DEFAULT_GAME
    try:
        main(game)
    except Exception:
        msg = ''.join(traceback.format_exc())
        messagebox.showerror('Exception', msg)
        if root:
            root.destroy()
        sys.exit(msg)
