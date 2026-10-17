; installer.nsh — electron-builder NSIS 추가 설정
; 설치: 탐색기 오른쪽 클릭 메뉴에 "Office Kit로 열기"를 단다 (사용자별, HKCU).
;       SystemFileAssociations 아래에만 쓰므로 기본 연결 프로그램은 바뀌지 않는다.
; 제거: 메뉴를 지우고, 앱이 만든 명령줄 파일(office-kit.cmd)도 지운다.

!define OK_VERB_KEY "shell\OfficeKit"

!macro OfficeKitAddVerb EXT
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\.${EXT}\${OK_VERB_KEY}" "" "Office Kit로 열기"
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\.${EXT}\${OK_VERB_KEY}" "Icon" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr HKCU "Software\Classes\SystemFileAssociations\.${EXT}\${OK_VERB_KEY}\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
!macroend

!macro OfficeKitRemoveVerb EXT
  DeleteRegKey HKCU "Software\Classes\SystemFileAssociations\.${EXT}\${OK_VERB_KEY}"
!macroend

; 지원 형식 (src/core/index.mjs 의 OPENABLE_EXTS 와 같게)
!macro OfficeKitForEachExt MACRO
  !insertmacro ${MACRO} hwp
  !insertmacro ${MACRO} hwpx
  !insertmacro ${MACRO} hwt
  !insertmacro ${MACRO} docx
  !insertmacro ${MACRO} doc
  !insertmacro ${MACRO} rtf
  !insertmacro ${MACRO} odt
  !insertmacro ${MACRO} pptx
  !insertmacro ${MACRO} ppt
  !insertmacro ${MACRO} odp
  !insertmacro ${MACRO} xlsx
  !insertmacro ${MACRO} xls
  !insertmacro ${MACRO} xlsm
  !insertmacro ${MACRO} csv
  !insertmacro ${MACRO} ods
!macroend

!macro customInstall
  !insertmacro OfficeKitForEachExt OfficeKitAddVerb
  ; 탐색기에 연결 정보가 바뀌었다고 알린다 (SHCNE_ASSOCCHANGED)
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
  !insertmacro OfficeKitForEachExt OfficeKitRemoveVerb
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'

  ; 업데이트가 아니라 진짜 제거일 때만: 앱이 만든 office-kit.cmd 지우기 (둘째 줄 표시로 확인)
  ${ifNot} ${isUpdated}
    StrCpy $R8 "$LOCALAPPDATA\Microsoft\WindowsApps\office-kit.cmd"
    ${if} ${FileExists} "$R8"
      ClearErrors
      FileOpen $R9 "$R8" r
      ${ifNot} ${Errors}
        FileRead $R9 $R7
        FileRead $R9 $R7
        FileClose $R9
        StrCpy $R7 $R7 19
        ${if} $R7 == "rem office-kit-shim"
          Delete "$R8"
        ${endIf}
      ${endIf}
    ${endIf}
  ${endIf}
!macroend
