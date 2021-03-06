; Script generated by the HM NIS Edit Script Wizard.

; HM NIS Edit Wizard helper defines
!define PRODUCT_NAME "BRI Reporting Tool"
!define PRODUCT_VERSION "1.0"
!define PRODUCT_PUBLISHER "Dhani Novan"
!define PRODUCT_WEB_SITE "http://www.bri.co.id"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\lua.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; MUI 1.67 compatible ------
!include "MUI.nsh"

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Welcome page
!insertmacro MUI_PAGE_WELCOME
; License page
!insertmacro MUI_PAGE_LICENSE "lisence.txt"
; Instfiles page
!insertmacro MUI_PAGE_INSTFILES
; Finish page
!define MUI_FINISHPAGE_RUN "$INSTDIR\lua.exe"
!define MUI_FINISHPAGE_RUN_PARAMETERS "$INSTDIR\script\AppDI321PNDIFF.lua"
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "English"

; MUI end ------

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "Setup_BRI_Reporting_Tool.exe"
InstallDir "C:\Lua"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

Section "MainSection" SEC01
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
  File "C:\Program Files (x86)\Lua\5.1\lua.exe"
  CreateDirectory "$SMPROGRAMS\BRI Reporting Tool"
  SetOutPath "$INSTDIR\script"
  File "AppDI321PNDIFF.lua"
  CreateShortCut "$DESKTOP\DI321.lnk" "$INSTDIR\lua.exe" "$INSTDIR\script\AppDI321PNDIFF.lua"
  CreateShortCut "$SMPROGRAMS\BRI Reporting Tool\DI321.lnk" "$INSTDIR\lua.exe" "$INSTDIR\script\AppDI321PNDIFF.lua"
  File "AppDI319PNDIFF.lua"
  CreateShortCut "$DESKTOP\DI319.lnk" "$INSTDIR\lua.exe" "$INSTDIR\script\AppDI319PNDIFF.lua"
  CreateShortCut "$SMPROGRAMS\BRI Reporting Tool\DI319.lnk" "$INSTDIR\lua.exe" "$INSTDIR\script\AppDI319PNDIFF.lua"
  File "AppCI324PNDIFF.lua"
  CreateShortCut "$DESKTOP\CI324.lnk" "$INSTDIR\lua.exe" "$INSTDIR\script\AppCI324PNDIFF.lua"
  CreateShortCut "$SMPROGRAMS\BRI Reporting Tool\CI324.lnk" "$INSTDIR\lua.exe" "$INSTDIR\script\AppCI324PNDIFF.lua"
  CreateShortCut "$SMPROGRAMS\BRI Reporting Tool\Uninstall.lnk" "$INSTDIR\uninst.exe"
  File "C:\Lua\script\gzip.lua"  
  WriteRegStr HKCR "*\shell\Compress with gzip\command" "" '"C:\Lua\lua.exe" "C:\Lua\script\gzip.lua" "%1"'
SectionEnd

Section "SampleReport" SEC02
  SetOutPath "$INSTDIR\data"
  File "data\20201231 CI324Modif.csv.gz"
  File "data\20201231 DI321 PN PENGELOLAH.csv.gz"
  File "data\20210130 DI321 PN PENGELOLAH.csv.gz"
  File "data\20210131 CI324Modif.csv.gz"
  File "data\20210131 DI319 MULTI PN.csv.gz"
  File "data\20210225 DI319 MULTI PN.csv.gz"
SectionEnd

;Section -AdditionalIcons
;  SetOutPath "$INSTDIR"
;  CreateShortCut "$SMPROGRAMS\BRI Reporting Tool\Uninstall.lnk" "$INSTDIR\uninst.exe"
;SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\lua.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\lua.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd


Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) was successfully removed from your computer."
FunctionEnd

Function un.onInit
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "Are you sure you want to completely remove $(^Name) and all of its components?" IDYES +2
  Abort
FunctionEnd

Section Uninstall
  Delete "$INSTDIR\uninst.exe"
  
  Delete "$INSTDIR\data\20201231 CI324Modif.csv.gz"
  Delete "$INSTDIR\data\20201231 DI321 PN PENGELOLAH.csv.gz"
  Delete "$INSTDIR\data\20210130 DI321 PN PENGELOLAH.csv.gz"
  Delete "$INSTDIR\data\20210131 CI324Modif.csv.gz"
  Delete "$INSTDIR\data\20210131 DI319 MULTI PN.csv.gz"
  Delete "$INSTDIR\data\20210225 DI319 MULTI PN.csv.gz"
  
  Delete "$INSTDIR\data\*.csv"
  Delete "$INSTDIR\data\*.htm"
  Delete "$INSTDIR\data\*.gz"
  Delete "$INSTDIR\script\*.lua"
  Delete "$INSTDIR\script\*.csv"
  Delete "$INSTDIR\script\*.htm"
  Delete "$INSTDIR\lua.exe"
  Delete "$INSTDIR\*.csv"
  Delete "$INSTDIR\*.htm"

  Delete "$SMPROGRAMS\BRI Reporting Tool\Uninstall.lnk"
  Delete "$DESKTOP\DI321.lnk"
  Delete "$SMPROGRAMS\BRI Reporting Tool\DI321.lnk"
  Delete "$DESKTOP\DI319.lnk"
  Delete "$SMPROGRAMS\BRI Reporting Tool\DI319.lnk"
  Delete "$DESKTOP\CI324.lnk"
  Delete "$SMPROGRAMS\BRI Reporting Tool\CI324.lnk"

  RMDir "$SMPROGRAMS\BRI Reporting Tool"
  RMDir "$INSTDIR\script"
  RMDir "$INSTDIR\data"
  RMDir "$INSTDIR"
  RMDir ""

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  DeleteRegKey HKCR "*\shell\Compress with gzip"
  SetAutoClose true
SectionEnd