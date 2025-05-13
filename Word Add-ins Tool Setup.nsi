!define LIBRARY_X64
!include "MUI.nsh"
!include "Library.nsh"

; Basic settings
Unicode true
ManifestDPIAware true
ManifestSupportedOS Win10
BrandingText ""

; HM NIS Edit Wizard helper defines
!define PRODUCT_NAME "Word Add-ins Tool"
!define PRODUCT_VERSION "1.0.0.0"
!define PRODUCT_PUBLISHER "ikiwi"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "Resources\favicon.ico"
!define MUI_UNICON "Resources\Uninstall.ico"

; Language Selection Dialog Settings
!define MUI_LANGDLL_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_LANGDLL_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_LANGDLL_REGISTRY_VALUENAME "NSIS:Language"

; Welcome page
!insertmacro MUI_PAGE_WELCOME
; License page
!insertmacro MUI_PAGE_LICENSE "LICENSE"
; Directory page
!insertmacro MUI_PAGE_DIRECTORY
; Instfiles page
!insertmacro MUI_PAGE_INSTFILES
; Finish page
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "SimpChinese"

; MUI end ------

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "Word Add-ins Tool Setup ${PRODUCT_VERSION}.exe"

; Set version info for the installer EXE
VIProductVersion "${PRODUCT_VERSION}"
VIAddVersionKey "FileDescription" "Word插件工具的安装程序，用于使用定制Word的模板进行便捷操作"
VIAddVersionKey "ProductName" "Office Word Add-ins Tool 安装程序"
VIAddVersionKey "ProductVersion" "${PRODUCT_VERSION}"
VIAddVersionKey "FileVersion" "1.0"
VIAddVersionKey "LegalCopyright" "Copyright 2025 ikiwi, all rights reserved"
VIAddVersionKey "CompanyName" "${PRODUCT_PUBLISHER}"

InstallDir "$PROGRAMFILES64\Word Add-ins Tool"
ShowInstDetails show
ShowUnInstDetails show

Function .onInit
  System::Call "user32::SetProcessDPIAware()"
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

Section "MainSection" SEC01

  SetOutPath "$INSTDIR"
  File "bin\Release\x64\Word Add-ins Tool.dll"
  File "bin\Release\x64\Word Add-ins Tool.dll.manifest"
  File "bin\Release\x64\Word Add-ins Tool.vsto"
  File "bin\Release\x64\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
  File "bin\Release\x64\Microsoft.Toolkit.Uwp.Notifications.dll"
  File "bin\Release\x64\Microsoft.VisualStudio.Tools.Applications.Runtime.dll"
  File "bin\Release\x64\Newtonsoft.Json.dll"
  File "bin\Release\x64\System.ValueTuple.dll"
  File "bin\Release\x64\UtfUnknown.dll"
  File "bin\Release\x64\README.md"
  File "bin\Release\x64\LICENSE"
  File "bin\Release\x64\markdown-test.md"
  File "bin\Release\x64\code-test.md"
  
  SetOutPath "$INSTDIR\Resources"
  File "bin\Release\x64\Resources\addin-template.docx"
  File "bin\Release\x64\Resources\favicon.ico"
  File "bin\Release\x64\Resources\Uninstall.ico"
  File "bin\Release\x64\Resources\Normal.dotm"
  File "bin\Release\x64\Resources\pandoc-reference.docx"
  File "bin\Release\x64\Resources\pandoc.exe"
  File "bin\Release\x64\Resources\replace_template.bat"
  SetOutPath "$INSTDIR\en-US"
  File "bin\Release\x64\en-US\Word Add-ins Tool.resources.dll"
SectionEnd

Section -AdditionalIcons
  CreateDirectory "$SMPROGRAMS\Word Add-ins Tool"
  CreateShortCut "$SMPROGRAMS\Word Add-ins Tool\卸载Word Add-ins Tooll.lnk" "$INSTDIR\uninst.exe" "" "$INSTDIR\Resources\Uninstall.ico"
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\Resources\favicon.ico"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"

  WriteRegStr HKLM "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool" "Description" "Word插件工具，用于使用定制Word的模板进行便捷操作"
  WriteRegStr HKLM "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool" "FriendlyName" "Word Add-ins Tool"
  WriteRegDWORD HKLM "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool" "LoadBehavior" 0x3
  WriteRegStr HKLM "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool" "Manifest" "file:///$INSTDIR/Word Add-ins Tool.vsto|vstolocal"

  WriteRegStr HKCU "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool" "Description" "Word插件工具，用于使用定制Word的模板进行便捷操作"
  WriteRegStr HKCU "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool" "FriendlyName" "Word Add-ins Tool"
  WriteRegDWORD HKCU "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool" "LoadBehavior" 0x3
  WriteRegStr HKCU "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool" "Manifest" "file:///$INSTDIR/Word Add-ins Tool.vsto|vstolocal"
SectionEnd

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) 已成功从您的计算机移除。"
FunctionEnd

Function un.onInit
System::Call "user32::SetProcessDPIAware()"
!insertmacro MUI_UNGETLANGUAGE
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "您确实要完全移除 $(^Name) 及其所有的组件吗？" IDYES +2
  Abort
FunctionEnd

Section Uninstall
  ; Remove Word add-in registry entries
  DeleteRegKey HKLM "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool"
  DeleteRegKey HKCU "Software\Microsoft\Office\Word\Addins\Word Add-ins Tool"
  
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\README.md"
  Delete "$INSTDIR\LICENSE"
  Delete "$INSTDIR\markdown-test.md"
  Delete "$INSTDIR\code-test.md"
  Delete "$INSTDIR\Resources\replace_template.bat"
  Delete "$INSTDIR\Resources\pandoc.exe"
  Delete "$INSTDIR\Resources\pandoc-reference.docx"
  Delete "$INSTDIR\Resources\Normal.dotm"
  Delete "$INSTDIR\Resources\favicon.ico"
  Delete "$INSTDIR\Resources\Uninstall.ico"
  Delete "$INSTDIR\Resources\addin-template.docx"
  Delete "$INSTDIR\Newtonsoft.Json.dll"
  Delete "$INSTDIR\UtfUnknown.dll"
  Delete "$INSTDIR\System.ValueTuple.dll"
  Delete "$INSTDIR\Microsoft.VisualStudio.Tools.Applications.Runtime.dll"
  Delete "$INSTDIR\Microsoft.Toolkit.Uwp.Notifications.dll"
  Delete "$INSTDIR\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
  Delete "$INSTDIR\Word Add-ins Tool.vsto"
  Delete "$INSTDIR\Word Add-ins Tool.dll.manifest"
  Delete "$INSTDIR\Word Add-ins Tool.dll"
  Delete "$INSTDIR\en-US\Word Add-ins Tool.resources.dll"

  Delete "$SMPROGRAMS\Word Add-ins Tool\卸载Word Add-ins Tooll.lnk"

  RMDir "$SMPROGRAMS\Word Add-ins Tool"
  RMDir "$INSTDIR\Resources"
  RMDir "$INSTDIR\en-US"
  RMDir "$INSTDIR"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  SetAutoClose true
SectionEnd
