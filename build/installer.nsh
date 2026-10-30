; installer.nsh — electron-builder NSIS 추가 설정
; 설치: 탐색기 오른쪽 클릭 메뉴에 "Office Kit로 열기"를 단다 (사용자별, HKCU).
;       SystemFileAssociations 아래에만 쓰므로 기본 연결 프로그램은 바뀌지 않는다.
;       Office Kit를 '기본 프로그램' 후보로 등록한다: ProgID(OfficeKit.<확장자>) + OpenWithProgids + Capabilities
;       + RegisteredApplications. 기본 프로그램은 사용자가 Windows 설정·'연결 프로그램'에서 고른다
;       (src/setup/fileassoc.mjs registrationEntries 와 같은 키 — 앱이 키를 고쳐 쓸 때도 같은 것을 쓴다).
; 제거: 메뉴·등록(과 함께 생긴 빈 키)을 지우고, 앱이 만든 명령줄 파일(office-kit.cmd)과 설치 파일 사본도 지운다.

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

; ── 기본 프로그램 후보 (ProgID · Capabilities) ─────────────

!ifndef OK_ROOT ; 시험할 때만 다른 키로 바꿔 끼운다
  !define OK_ROOT "Software"
!endif
!define OK_CLASSES "${OK_ROOT}\Classes"
!define OK_CAPS "${OK_ROOT}\OfficeKit\Capabilities"
; Applications\<exe> 키는 매크로 안에서 쓴다 (APP_EXECUTABLE_FILENAME은 이 파일을 읽은 뒤에 정해질 수 있어 !define에 넣지 않는다)

!macro OfficeKitRegisterType EXT DESC
  WriteRegStr HKCU "${OK_CLASSES}\OfficeKit.${EXT}" "" "${DESC}"
  WriteRegStr HKCU "${OK_CLASSES}\OfficeKit.${EXT}" "FriendlyTypeName" "${DESC}"
  WriteRegStr HKCU "${OK_CLASSES}\OfficeKit.${EXT}\DefaultIcon" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr HKCU "${OK_CLASSES}\OfficeKit.${EXT}\Application" "ApplicationName" "Office Kit"
  WriteRegStr HKCU "${OK_CLASSES}\OfficeKit.${EXT}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  WriteRegStr HKCU "${OK_CLASSES}\.${EXT}\OpenWithProgids" "OfficeKit.${EXT}" ""
  WriteRegStr HKCU "${OK_CAPS}\FileAssociations" ".${EXT}" "OfficeKit.${EXT}"
  WriteRegStr HKCU "${OK_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".${EXT}" ""
!macroend

!macro OfficeKitUnregisterType EXT DESC
  DeleteRegKey HKCU "${OK_CLASSES}\OfficeKit.${EXT}"
  DeleteRegValue HKCU "${OK_CLASSES}\.${EXT}\OpenWithProgids" "OfficeKit.${EXT}"
  !insertmacro OfficeKitDeleteKeyIfEmpty "${OK_CLASSES}\.${EXT}\OpenWithProgids"
  !insertmacro OfficeKitDeleteKeyIfEmpty "${OK_CLASSES}\.${EXT}"
!macroend

; 형식과 탐색기 '유형' 이름 (fileassoc.mjs ASSOC_TYPES 와 같게)
!macro OfficeKitForEachType MACRO
  !insertmacro ${MACRO} hwp "한글 문서"
  !insertmacro ${MACRO} hwpx "한글 문서 (HWPX)"
  !insertmacro ${MACRO} hwt "한글 서식 파일"
  !insertmacro ${MACRO} docx "Word 문서"
  !insertmacro ${MACRO} doc "Word 97-2003 문서"
  !insertmacro ${MACRO} rtf "서식 있는 텍스트 문서"
  !insertmacro ${MACRO} odt "ODF 텍스트 문서"
  !insertmacro ${MACRO} pptx "PowerPoint 프레젠테이션"
  !insertmacro ${MACRO} ppt "PowerPoint 97-2003 프레젠테이션"
  !insertmacro ${MACRO} odp "ODF 프레젠테이션"
  !insertmacro ${MACRO} xlsx "Excel 통합 문서"
  !insertmacro ${MACRO} xls "Excel 97-2003 통합 문서"
  !insertmacro ${MACRO} xlsm "Excel 매크로 사용 통합 문서"
  !insertmacro ${MACRO} csv "CSV 파일"
  !insertmacro ${MACRO} ods "ODF 스프레드시트"
  !insertmacro ${MACRO} pdf "PDF 문서"
!macroend

!macro OfficeKitRegisterApp
  !insertmacro OfficeKitForEachType OfficeKitRegisterType
  WriteRegStr HKCU "${OK_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}" "FriendlyAppName" "Office Kit"
  WriteRegStr HKCU "${OK_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}\DefaultIcon" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr HKCU "${OK_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  WriteRegStr HKCU "${OK_CAPS}" "ApplicationName" "Office Kit"
  WriteRegStr HKCU "${OK_CAPS}" "ApplicationDescription" "한글·Word·PowerPoint·Excel·PDF 문서를 보고, 직접 고치고, AI 대화로 수정하는 프로그램"
  WriteRegStr HKCU "${OK_CAPS}" "ApplicationIcon" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr HKCU "${OK_ROOT}\RegisteredApplications" "Office Kit" "${OK_CAPS}"
!macroend

!macro OfficeKitUnregisterApp
  !insertmacro OfficeKitForEachType OfficeKitUnregisterType
  DeleteRegKey HKCU "${OK_CLASSES}\Applications\${APP_EXECUTABLE_FILENAME}"
  DeleteRegValue HKCU "${OK_ROOT}\RegisteredApplications" "Office Kit"
  DeleteRegKey HKCU "${OK_ROOT}\OfficeKit"
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
  !insertmacro ${MACRO} pdf
!macroend

!macro customInstall
  !insertmacro OfficeKitForEachExt OfficeKitAddVerb
  !insertmacro OfficeKitRegisterApp
  ; 탐색기에 연결 정보가 바뀌었다고 알린다 (SHCNE_ASSOCCHANGED)
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
  !insertmacro OfficeKitForEachExt OfficeKitRemoveVerb
  ; 기본 프로그램 등록은 진짜 제거일 때만 지운다 (업데이트 중에 지우면 Windows가 고른 기본 프로그램이 잠깐 끊긴다)
  ${ifNot} ${isUpdated}
    !insertmacro OfficeKitUnregisterApp
  ${endIf}
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
