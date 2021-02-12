; FileBrowser Installer NSIS Script
; get NSIS at http://nsis.sourceforge.net/Download

!define ProgramName "WalkmanUtils"
; Icon ""

Name "${ProgramName}"
Caption "${ProgramName} Installer"
XPStyle on
ShowInstDetails show
AutoCloseWindow true

LicenseBkColor /windows
LicenseData "LICENSE.md"
LicenseForceSelection checkbox "I have read and understand this notice"
LicenseText "Please read the notice below before installing ${ProgramName}. If you understand the notice, click the checkbox below and click Next."

InstallDir $PROGRAMFILES\WalkmanOSS\${ProgramName}
InstallDirRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "InstallLocation"
OutFile "bin\Release\${ProgramName}-Installer.exe"

; Pages

Page license
Page components
Page directory
Page instfiles
Page custom postInstallShow postInstallFinish ": Install Complete"
UninstPage uninstConfirm
UninstPage instfiles

; Sections

Section "Executables & Uninstaller"
  SectionIn RO
  SetOutPath $INSTDIR
  File "bin\Release\CompressFile.exe"
  File "bin\Release\CompressFile.exe.config"
  File "bin\Release\GetCompressedSize.exe"
  File "bin\Release\GetCompressedSize.exe.config"
  File "bin\Release\GetFolderIcon.exe"
  File "bin\Release\GetFolderIcon.exe.config"
  File "bin\Release\GetOpenWith.exe"
  File "bin\Release\GetOpenWith.exe.config"
  File "bin\Release\HandleManager.exe"
  File "bin\Release\HandleManager.exe.config"
  ; File "bin\Release\Hash.exe"
  ; File "bin\Release\Hash.exe.config"
  File "bin\Release\IsAdmin.exe"
  File "bin\Release\IsAdmin.exe.config"
  File "bin\Release\MsgBox.exe"
  File "bin\Release\MsgBox.exe.config"
  File "bin\Release\MsgBox.com"
  File "bin\Release\MsgBox.com.config"
  File "bin\Release\OpenWith.exe"
  File "bin\Release\OpenWith.exe.config"
  File "bin\Release\RunAsAdmin.exe"
  File "bin\Release\RunAsAdmin.exe.config"
  File "bin\Release\SetAttribute.exe"
  File "bin\Release\SetAttribute.exe.config"
  File "bin\Release\TakeOwn.exe"
  File "bin\Release\TakeOwn.exe.config"
  File "bin\Release\UncompressFile.exe"
  File "bin\Release\UncompressFile.exe.config"
  File "bin\Release\WinProperties.exe"
  File "bin\Release\WinProperties.exe.config"
  WriteUninstaller "${ProgramName}-Uninst.exe"
SectionEnd

Section "Add to Windows Programs & Features"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "DisplayName" "${ProgramName}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "Publisher" "WalkmanOSS"
  
  ; WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "DisplayIcon" "$INSTDIR\${ProgramName}.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "InstallLocation" "$INSTDIR\"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "UninstallString" "$INSTDIR\${ProgramName}-Uninst.exe"
  
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "NoRepair" 1
  
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "HelpLink" "https://github.com/Walkman100/${ProgramName}/issues/new"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "URLInfoAbout" "https://github.com/Walkman100/${ProgramName}" ; Support Link
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" "URLUpdateInfo" "https://github.com/Walkman100/${ProgramName}/releases" ; Update Info Link
SectionEnd

Section "Start Menu Shortcut"
  CreateDirectory "$SMPROGRAMS\WalkmanOSS"
  ; CreateShortCut "$SMPROGRAMS\WalkmanOSS\${ProgramName}.lnk" "$INSTDIR\${ProgramName}.exe" "" "$INSTDIR\${ProgramName}.exe" "" "" "" "${ProgramName}"
  CreateShortCut "$SMPROGRAMS\WalkmanOSS\Uninstall ${ProgramName}.lnk" "$INSTDIR\${ProgramName}-Uninst.exe" "" "" "" "" "" "Uninstall ${ProgramName}"
  ;Syntax for CreateShortCut: link.lnk target.file [parameters [icon.file [icon_index_number [start_options [keyboard_shortcut [description]]]]]]
SectionEnd

; Section "Desktop Shortcut"
  ; CreateShortCut "$DESKTOP\${ProgramName}.lnk" "$INSTDIR\${ProgramName}.exe" "" "$INSTDIR\${ProgramName}.exe" "" "" "" "${ProgramName}"
; SectionEnd

; Section "Quick Launch Shortcut"
  ; CreateShortCut "$QUICKLAUNCH\${ProgramName}.lnk" "$INSTDIR\${ProgramName}.exe" "" "$INSTDIR\${ProgramName}.exe" "" "" "" "${ProgramName}"
; SectionEnd

; Functions

Function .onInit
  SetShellVarContext all
  SetAutoClose true
FunctionEnd

; Custom Install Complete page

!include nsDialogs.nsh
!include LogicLib.nsh ; For ${IF} logic
Var Dialog
Var Label
Var CheckboxReadme
Var CheckboxReadme_State
; Var CheckboxRunProgram
; Var CheckboxRunProgram_State

Function postInstallShow
  nsDialogs::Create 1018
  Pop $Dialog
  ${If} $Dialog == error
    Abort
  ${EndIf}
  
  ${NSD_CreateLabel} 0 0 100% 12u "Setup will launch these tasks when you click close:"
  Pop $Label
  
  ${NSD_CreateCheckbox} 10u 30u 100% 10u "&Open Readme"
  Pop $CheckboxReadme
  ${If} $CheckboxReadme_State == ${BST_CHECKED}
    ${NSD_Check} $CheckboxReadme
  ${EndIf}
  
  ; ${NSD_CreateCheckbox} 10u 50u 100% 10u "&Launch ${ProgramName}"
  ; Pop $CheckboxRunProgram
  ; ${If} $CheckboxRunProgram_State == ${BST_CHECKED}
    ; ${NSD_Check} $CheckboxRunProgram
  ; ${EndIf}
  
  # alternative for the above ${If}:
  #${NSD_SetState} $Checkbox_State
  nsDialogs::Show
FunctionEnd

Function postInstallFinish
  ${NSD_GetState} $CheckboxReadme $CheckboxReadme_State
  ; ${NSD_GetState} $CheckboxRunProgram $CheckboxRunProgram_State
  
  ${If} $CheckboxReadme_State == ${BST_CHECKED}
    ExecShell "open" "https://github.com/Walkman100/${ProgramName}/blob/master/README.md"
  ${EndIf}
  ; ${If} $CheckboxRunProgram_State == ${BST_CHECKED}
    ; ExecShell "open" "$INSTDIR\${ProgramName}.exe"
  ; ${EndIf}
FunctionEnd

; Uninstaller

Section "Uninstall"
  Delete "$INSTDIR\${ProgramName}-Uninst.exe" ; Remove Application Files
  Delete "$INSTDIR\CompressFile.exe"
  Delete "$INSTDIR\CompressFile.exe.config"
  Delete "$INSTDIR\GetCompressedSize.exe"
  Delete "$INSTDIR\GetCompressedSize.exe.config"
  Delete "$INSTDIR\GetFolderIcon.exe"
  Delete "$INSTDIR\GetFolderIcon.exe.config"
  Delete "$INSTDIR\GetOpenWith.exe"
  Delete "$INSTDIR\GetOpenWith.exe.config"
  Delete "$INSTDIR\HandleManager.exe"
  Delete "$INSTDIR\HandleManager.exe.config"
  ; Delete "$INSTDIR\Hash.exe"
  ; Delete "$INSTDIR\Hash.exe.config"
  Delete "$INSTDIR\IsAdmin.exe"
  Delete "$INSTDIR\IsAdmin.exe.config"
  Delete "$INSTDIR\MsgBox.exe"
  Delete "$INSTDIR\MsgBox.exe.config"
  Delete "$INSTDIR\MsgBox.com"
  Delete "$INSTDIR\MsgBox.com.config"
  Delete "$INSTDIR\OpenWith.exe"
  Delete "$INSTDIR\OpenWith.exe.config"
  Delete "$INSTDIR\RunAsAdmin.exe"
  Delete "$INSTDIR\RunAsAdmin.exe.config"
  Delete "$INSTDIR\SetAttribute.exe"
  Delete "$INSTDIR\SetAttribute.exe.config"
  Delete "$INSTDIR\TakeOwn.exe"
  Delete "$INSTDIR\TakeOwn.exe.config"
  Delete "$INSTDIR\UncompressFile.exe"
  Delete "$INSTDIR\UncompressFile.exe.config"
  Delete "$INSTDIR\WinProperties.exe"
  Delete "$INSTDIR\WinProperties.exe.config"
  RMDir "$INSTDIR"
  
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgramName}" ; Remove Windows Programs & Features integration (uninstall info)
  
  ; Delete "$SMPROGRAMS\WalkmanOSS\${ProgramName}.lnk" ; Remove Start Menu Shortcuts & Folder
  Delete "$SMPROGRAMS\WalkmanOSS\Uninstall ${ProgramName}.lnk"
  RMDir "$SMPROGRAMS\WalkmanOSS"
  
  ; Delete "$DESKTOP\${ProgramName}.lnk"     ; Remove Desktop      Shortcut
  ; Delete "$QUICKLAUNCH\${ProgramName}.lnk" ; Remove Quick Launch Shortcut
SectionEnd

; Uninstaller Functions

Function un.onInit
  SetShellVarContext all
  SetAutoClose true
FunctionEnd

Function un.onUninstFailed
  MessageBox MB_OK "Uninstall Cancelled"
FunctionEnd

Function un.onUninstSuccess
  MessageBox MB_OK "Uninstall Completed"
FunctionEnd
