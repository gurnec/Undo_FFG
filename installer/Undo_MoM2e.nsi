SetCompressor /FINAL /SOLID lzma

!include MUI2.nsh

Name "Undo for FFG Games"
OutFile "Undo_v2.1_for_FFG_setup.exe"

RequestExecutionLevel admin
ManifestSupportedOS all

InstallDir "$PROGRAMFILES\Undo for FFG Games"
InstallDirRegKey HKLM "Software\Undo for MoM2e" ""

!define MUI_ICON ..\Undo_MoM2e.ico
!define MUI_ABORTWARNING


Var StartMenuFolder

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "..\LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY

!define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKLM"
!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\Undo for MoM2e"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "FFG Start Menu folder"
!insertmacro MUI_PAGE_STARTMENU Application $StartMenuFolder

!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH


!insertmacro MUI_LANGUAGE "English"

VIProductVersion "2.1.0.0"
VIAddVersionKey /LANG=${LANG_ENGLISH} "ProductName" "Undo for FFG Games"
VIAddVersionKey /LANG=${LANG_ENGLISH} "ProductVersion" "2.1"
VIAddVersionKey /LANG=${LANG_ENGLISH} "Comments" "Installer distributed from https://github.com/gurnec/Undo_FFG/releases"
VIAddVersionKey /LANG=${LANG_ENGLISH} "LegalCopyright" "Copyright © 2017 Christopher Gurnee. All rights reserved."
VIAddVersionKey /LANG=${LANG_ENGLISH} "FileDescription" "Undo for FFG Games Installer"
VIAddVersionKey /LANG=${LANG_ENGLISH} "FileVersion" "2.1"


; The install script
;
Section

    SetOutPath "$INSTDIR"
    SetShellVarContext all

    ; Remove files from older versions which aren't overwritten by this version
    Delete /REBOOTOK "$INSTDIR\python36.dll"
    Delete /REBOOTOK "$INSTDIR\unicodedata.pyd"
    SetRebootFlag false

    ; Install the Visual Studio redistributable
    GetTempFileName $0
    File /oname=$0 VC_redist.x86.exe
    ExecWait '"$0" /quiet /norestart' $1
    ${If} $1 == 1638
        ; Installation stopped because a newer version is already installed
        ClearErrors
    ${ElseIf} $1 == 3010
        ; Installation succeeded, but a reboot is required
        SetRebootFlag true
        ClearErrors
    ${ElseIf} $1 != 0
        ; Installation failed; run it again, but interactively to possibly show a better error message
        ExecWait '"$0"'
        IfErrors abort_on_error
    ${EndIf}

    ; Install the files
    File /r ..\dist\Undo_MoM2e\*.*
    File ..\LICENSE.txt
    IfErrors abort_on_error

    ; Delete any old Start Menu shortcuts (the Start Menu folder has been renamed)
    ReadRegStr $1 HKLM "Software\Undo for MoM2e" "Start Menu folder"
    ${if} $1 != ""
        Delete "$SMPROGRAMS\$1\Undo for Mansions of Madness.lnk"
        Delete "$SMPROGRAMS\$1\Uninstall Undo for MoM.lnk"
        RMDir  "$SMPROGRAMS\$1"
        DeleteRegValue HKLM "Software\Undo for MoM2e" "Start Menu folder"
    ${EndIf}

    ; Install the Start Menu shortcuts
    !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
        CreateDirectory "$SMPROGRAMS\$StartMenuFolder"
        CreateShortcut  "$SMPROGRAMS\$StartMenuFolder\Undo for Mansions of Madness.lnk"     "$INSTDIR\Undo_MoM2e.exe" "--game=mom"
        CreateShortcut  "$SMPROGRAMS\$StartMenuFolder\Undo for Road to Legend.lnk"          "$INSTDIR\Undo_MoM2e.exe" "--game=rtl"
        CreateShortcut  "$SMPROGRAMS\$StartMenuFolder\Undo for Legends of the Alliance.lnk" "$INSTDIR\Undo_MoM2e.exe" "--game=lota"
        CreateShortcut  "$SMPROGRAMS\$StartMenuFolder\Uninstall Undo for FFG.lnk"           "$INSTDIR\Uninstall.exe"
    !insertmacro MUI_STARTMENU_WRITE_END

    ; Install the uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    WriteRegStr   HKLM "Software\Undo for MoM2e" "" $INSTDIR
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "DisplayIcon" "$INSTDIR\Undo_MoM2e.ico"
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "DisplayName" "Undo for FFG Games"
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "DisplayVersion" "2.1"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "EstimatedSize" 16722
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "HelpLink" "https://github.com/gurnec/Undo_FFG/issues"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "NoRepair" 1
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "URLInfoAbout" "https://github.com/gurnec/Undo_FFG"
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "URLUpdateInfo" "https://github.com/gurnec/Undo_FFG/releases/latest"

    Delete $0
    Return

    abort_on_error:
        IfSilent +2
        MessageBox MB_ICONSTOP|MB_OK "An unexpected error occurred during installation"
        Delete $0
        Quit

SectionEnd


Section "Uninstall"

    SetShellVarContext all

    Delete   /REBOOTOK "$INSTDIR\Undo_MoM2e.exe"
    Delete   /REBOOTOK "$INSTDIR\Undo_MoM2e.exe.manifest"
    Delete   /REBOOTOK "$INSTDIR\Uninstall.exe"
    Delete   /REBOOTOK "$INSTDIR\Undo_MoM2e.ico"
    Delete   /REBOOTOK "$INSTDIR\LICENSE.txt"
    Delete   /REBOOTOK "$INSTDIR\python38.dll"
    Delete   /REBOOTOK "$INSTDIR\tcl86t.dll"
    Delete   /REBOOTOK "$INSTDIR\tk86t.dll"
    Delete   /REBOOTOK "$INSTDIR\libffi-7.dll"
    Delete   /REBOOTOK "$INSTDIR\VCRUNTIME140.dll"
    Delete   /REBOOTOK "$INSTDIR\base_library.zip"
    Delete   /REBOOTOK "$INSTDIR\*.pyd"
    RMDir /r /REBOOTOK "$INSTDIR\tcl"
    RMDir /r /REBOOTOK "$INSTDIR\tk"
    RMDir    /REBOOTOK "$INSTDIR"

    !insertmacro MUI_STARTMENU_GETFOLDER Application $StartMenuFolder
    Delete /REBOOTOK "$SMPROGRAMS\$StartMenuFolder\Undo for Mansions of Madness.lnk"
    Delete /REBOOTOK "$SMPROGRAMS\$StartMenuFolder\Undo for Road to Legend.lnk"
    Delete /REBOOTOK "$SMPROGRAMS\$StartMenuFolder\Undo for Legends of the Alliance.lnk"
    Delete /REBOOTOK "$SMPROGRAMS\$StartMenuFolder\Uninstall Undo for FFG.lnk"
    RMDir  /REBOOTOK "$SMPROGRAMS\$StartMenuFolder"

    DeleteRegKey HKLM "Software\Undo for MoM2e"  ; settings are in %APPDATA%, they aren't deleted
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e"

SectionEnd
