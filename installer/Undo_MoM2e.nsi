SetCompressor /FINAL /SOLID lzma

!include MUI2.nsh

Name "Undo for Mansions of Madness"
OutFile "Undo_v1.0_for_MoM2e_setup.exe"

RequestExecutionLevel admin
ManifestSupportedOS all

InstallDir "$PROGRAMFILES\Undo for MoM2e"
InstallDirRegKey HKLM "Software\Undo for MoM2e" ""

!define MUI_ICON ..\Undo_MoM2e.ico
!define MUI_ABORTWARNING


Var StartMenuFolder

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "..\LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY

!define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKLM"
!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\Undo for MoM2e"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "Start Menu folder"
!insertmacro MUI_PAGE_STARTMENU Application $StartMenuFolder

!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH


!insertmacro MUI_LANGUAGE "English"

VIProductVersion "1.0.0.0"
VIAddVersionKey /LANG=${LANG_ENGLISH} "ProductName" "Undo for Mansions of Madness"
VIAddVersionKey /LANG=${LANG_ENGLISH} "ProductVersion" "1.0"
VIAddVersionKey /LANG=${LANG_ENGLISH} "Comments" "Installer distributed from https://github.com/gurnec/Undo_MoM2e/releases"
VIAddVersionKey /LANG=${LANG_ENGLISH} "LegalCopyright" "Copyright © 2017 Christopher Gurnee. All rights reserved."
VIAddVersionKey /LANG=${LANG_ENGLISH} "FileDescription" "Installer from https://github.com/gurnec/Undo_MoM2e/releases"
VIAddVersionKey /LANG=${LANG_ENGLISH} "FileVersion" "1.0"


; The install script
;
Section

    SetOutPath "$INSTDIR"
    SetShellVarContext all

    ; Install the Visual Studio redistributable
    GetTempFileName $0
    File /oname=$0 vc_redist.x86.exe
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

    ; Install the Start Menu shortcuts
    !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
        CreateDirectory "$SMPROGRAMS\$StartMenuFolder"
        CreateShortcut  "$SMPROGRAMS\$StartMenuFolder\Undo for Mansions of Madness.lnk" "$INSTDIR\Undo_MoM2e.exe"
        CreateShortcut  "$SMPROGRAMS\$StartMenuFolder\Uninstall Undo for MoM.lnk" "$INSTDIR\Uninstall.exe"
    !insertmacro MUI_STARTMENU_WRITE_END

    ; Install the uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    WriteRegStr   HKLM "Software\Undo for MoM2e" "" $INSTDIR
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "DisplayIcon" "$INSTDIR\Undo_MoM2e.ico"
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "DisplayName" "Undo for Mansions of Madness"
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "DisplayVersion" "1.0"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "EstimatedSize" 16690
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "HelpLink" "https://github.com/gurnec/Undo_MoM2e/issues"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "NoRepair" 1
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "URLInfoAbout" "https://github.com/gurnec/Undo_MoM2e"
    WriteRegStr   HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e" "URLUpdateInfo" "https://github.com/gurnec/Undo_MoM2e/releases/latest"

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
    Delete   /REBOOTOK "$INSTDIR\python36.dll"
    Delete   /REBOOTOK "$INSTDIR\tcl86t.dll"
    Delete   /REBOOTOK "$INSTDIR\tk86t.dll"
    Delete   /REBOOTOK "$INSTDIR\VCRUNTIME140.dll"
    Delete   /REBOOTOK "$INSTDIR\base_library.zip"
    Delete   /REBOOTOK "$INSTDIR\*.pyd"
    RMDir /r /REBOOTOK "$INSTDIR\tcl"
    RMDir /r /REBOOTOK "$INSTDIR\tk"
    RMDir    /REBOOTOK "$INSTDIR"

    !insertmacro MUI_STARTMENU_GETFOLDER Application $StartMenuFolder
    Delete "$SMPROGRAMS\$StartMenuFolder\Undo for Mansions of Madness.lnk"
    Delete "$SMPROGRAMS\$StartMenuFolder\Uninstall Undo for MoM.lnk"
    RMDir  "$SMPROGRAMS\$StartMenuFolder"

    DeleteRegKey HKLM "Software\Undo for MoM2e"  ; settings are in %APPDATA%, they aren't deleted
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Undo for MoM2e"

SectionEnd
