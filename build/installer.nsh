; installer.nsh — electron-builder NSIS 추가 설정
; 설치: 탐색기 오른쪽 클릭 메뉴에 "Office Kit로 열기"를 단다 (사용자별, HKCU).
;       SystemFileAssociations 아래에만 쓰므로 기본 연결 프로그램은 바뀌지 않는다.
; 제거: 메뉴(와 함께 생긴 빈 키)를 지우고, 앱이 만든 명령줄 파일(office-kit.cmd)과 설치 파일 사본도 지운다.

!ifndef OK_SFA ; 시험할 때만 다른 키로 바꿔 끼운다
  !define OK_SFA "Software\Classes\SystemFileAssociations"
!endif
!define OK_VERB_KEY "shell\OfficeKit"

!macro OfficeKitAddVerb EXT
  WriteRegStr HKCU "${OK_SFA}\.${EXT}\${OK_VERB_KEY}" "" "Office Kit로 열기"
  WriteRegStr HKCU "${OK_SFA}\.${EXT}\${OK_VERB_KEY}" "Icon" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr HKCU "${OK_SFA}\.${EXT}\${OK_VERB_KEY}\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
!macroend

; 하위 키도 값도 없는 키만 지운다 (설치할 때 함께 생긴 .확장자\shell, .확장자).
; NSIS 3.04의 DeleteRegKey /ifempty는 값은 보지 않으므로, 다른 프로그램이 넣은 값이 있으면 남긴다
!macro OfficeKitDeleteKeyIfEmpty KEY
  EnumRegKey $R7 HKCU "${KEY}" 0
  EnumRegValue $R8 HKCU "${KEY}" 0
  ClearErrors
  ReadRegStr $R9 HKCU "${KEY}" "" ; 기본값이 있으면 오류 없이 읽힌다
  ${if} $R7 == ""
  ${andIf} $R8 == ""
  ${andIf} ${Errors}
    DeleteRegKey /ifempty HKCU "${KEY}"
  ${endIf}
!macroend

!macro OfficeKitRemoveVerb EXT
  DeleteRegKey HKCU "${OK_SFA}\.${EXT}\${OK_VERB_KEY}"
  !insertmacro OfficeKitDeleteKeyIfEmpty "${OK_SFA}\.${EXT}\shell"
  !insertmacro OfficeKitDeleteKeyIfEmpty "${OK_SFA}\.${EXT}"
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

  ; 업데이트가 아니라 진짜 제거일 때만
  ${ifNot} ${isUpdated}
    ; 앱이 만든 office-kit.cmd 지우기 (둘째 줄 표시로 확인)
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
    ; electron-builder가 설치할 때마다 남기는 설치 파일 사본 (자동 업데이트용 — Office Kit는 쓰지 않음).
    ; 이름은 package.json name + "-updater"
    RMDir /r "$LOCALAPPDATA\office-kit-updater"
  ${endIf}
!macroend
