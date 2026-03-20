; ============================================
; COPYSTYLE.AU3 - Module Copy Style
; Su dung _SmartCopyStyle (Ultimate Fix)
; FIX: Save truoc copy, xu ly ten file co dau cham, loc style
; REFACTORED: Cac ham hotkey da duoc tach ra thanh module rieng
; Xem: Modules\StyleHotkey.au3 va GUI\HotkeyDialogs.au3
; ============================================

#include-once

; NOTE: Constants da duoc dinh nghia trong Config.au3:
; - $wdOrganizerObjectStyles = 0
; - $wdCollapseEnd = 0
; - $wdFormatOriginalFormatting = 16
; - $g_aImportantStyles[]

; =========================================================
; HAM CORE: _SmartCopyStyle (Ultimate Fix)
; Tac dung: Copy style an toan, tu dong Save va ep hien thi
; Tra ve: 1 (Thanh cong), 0 (That bai)
; =========================================================
Func _SmartCopyStyle($oWordApp, $oDocSource, $oDocTarget, $sStyleName)
    ; 1. Kiem tra doi tuong hop le
    If Not IsObj($oDocSource) Or Not IsObj($oDocTarget) Then Return 0
    If Not IsObj($oWordApp) Then Return 0
    
    ; 2. Lam sach ten Style (Xoa khoang trang thua - Rat quan trong)
    $sStyleName = StringStripWS($sStyleName, 3)
    If $sStyleName = "" Then Return 0

    Local $sSrcPath = $oDocSource.FullName
    Local $sDstPath = $oDocTarget.FullName

    ; ---------------------------------------------------------
    ; BUOC QUAN TRONG NHAT: BAT BUOC LUU FILE TRUOC KHI COPY
    ; De dong bo du lieu tu RAM xuong HDD cho lenh OrganizerCopy doc
    ; ---------------------------------------------------------
    $oDocSource.Save()
    $oDocTarget.Save()
    ; ---------------------------------------------------------

    ; 3. Thuc hien Copy bang OrganizerCopy (Mode 3 = Styles)
    $oWordApp.OrganizerCopy($sSrcPath, $sDstPath, $sStyleName, $wdOrganizerObjectStyles)
    
    ; Kiem tra ngay lap tuc xem co loi khong
    If @error Then
        ConsoleWrite("! Loi: Khong tim thay style '" & $sStyleName & "' trong file nguon." & @CRLF)
        Return 0
    EndIf

    ; 4. Ep Style hien len thanh Gallery (Quick Style)
    Local $oStyle = $oDocTarget.Styles($sStyleName)
    If IsObj($oStyle) Then
        $oStyle.QuickStyle = True       ; Hien len Gallery
        $oStyle.Priority = 1            ; Dua len vi tri dau tien
        $oStyle.UnhideWhenUsed = False  ; Bo an
        $oStyle.Hidden = False          ; Bo an
        
        ; Meo: Refresh giao dien Word bang cach kich hoat lai cua so
        $oDocTarget.Activate()
        Return 1 ; Thanh cong
    Else
        ConsoleWrite("! Loi: Copy xong nhung khong truy cap duoc Style object." & @CRLF)
        Return 0
    EndIf
EndFunc

; === HELPER: Lay Document tu ComboBox ===
; FIX: Xu ly ten file co dau cham (vd: "Bao.Cao.docx")
Func _GetSelectedDoc($hCombo)
    Local $sSelected = GUICtrlRead($hCombo)
    If $sSelected = "" Then Return 0
    
    ; FIX: Chi lay phan so truoc dau cham DAU TIEN
    Local $iDotPos = StringInStr($sSelected, ".")
    If $iDotPos = 0 Then Return 0
    
    Local $iIndex = Int(StringLeft($sSelected, $iDotPos - 1))
    
    ; Validate index
    If Not IsObj($g_oWord) Then Return 0
    If $iIndex < 1 Or $iIndex > $g_oWord.Documents.Count Then
        MsgBox($MB_ICONWARNING, "Loi", "File khong ton tai hoac da bi dong!")
        Return 0
    EndIf
    
    Return $g_oWord.Documents.Item($iIndex)
EndFunc

; === FUNCTION: Refresh List ===
Func _RefreshStyleDocsList()
    $g_oWord = ObjGet("", "Word.Application")
    If Not IsObj($g_oWord) Then
        MsgBox($MB_ICONWARNING, "Loi", "Vui long mo Word truoc!")
        Return
    EndIf

    Local $oDocs = $g_oWord.Documents
    If $oDocs.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Loi", "Chua co file nao mo!")
        Return
    EndIf
    
    Local $sDocList = ""
    For $i = 1 To $oDocs.Count
        $sDocList &= $i & ". " & $oDocs.Item($i).Name & "|"
    Next
    
    ; Xoa ky tu | cuoi cung
    $sDocList = StringTrimRight($sDocList, 1)

    GUICtrlSetData($g_cboSourceDoc, "")
    GUICtrlSetData($g_cboTargetDoc, "")
    GUICtrlSetData($g_cboSourceDoc, $sDocList)
    GUICtrlSetData($g_cboTargetDoc, $sDocList)
    
    ; Auto select
    If $oDocs.Count >= 1 Then GUICtrlSendMsg($g_cboSourceDoc, 0x14E, 0, 0)
    If $oDocs.Count >= 2 Then GUICtrlSendMsg($g_cboTargetDoc, 0x14E, 1, 0)
    
    _UpdateProgress("Da lam moi danh sach file!")
EndFunc

; === FUNCTION: Copy All Styles ===
Func _CopyAllStyles()
    Local $oSource = _GetSelectedDoc($g_cboSourceDoc)
    Local $oTarget = _GetSelectedDoc($g_cboTargetDoc)
    
    If Not IsObj($oSource) Or Not IsObj($oTarget) Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon file nguon va dich!")
        Return
    EndIf
    
    If $oSource.FullName = $oTarget.FullName Then
        MsgBox($MB_ICONWARNING, "Loi", "File nguon va dich phai khac nhau!")
        Return
    EndIf
    
    _UpdateProgress("Dang copy styles...")
    
    Local $iCopied = 0
    Local $iFailed = 0
    
    ; FIX: Chi copy style dang dung (InUse) hoac user-defined (khong BuiltIn)
    Local $oStyles = $oSource.Styles
    Local $iTotal = $oStyles.Count
    
    For $i = 1 To $iTotal
        Local $oStyle = $oStyles.Item($i)
        If Not IsObj($oStyle) Then ContinueLoop
        
        ; FIX: Loc style - chi lay InUse hoac khong phai BuiltIn
        If $oStyle.InUse = False And $oStyle.BuiltIn = True Then ContinueLoop
        
        Local $sName = $oStyle.NameLocal
        If StringLeft($sName, 1) = "_" Then ContinueLoop ; Skip internal
        
        ; Cap nhat progress
        If Mod($i, 20) = 0 Then
            _UpdateProgress("Dang xu ly... " & $i & "/" & $iTotal)
        EndIf
        
        ; SU DUNG _SmartCopyStyle (Ultimate Fix)
        Local $iResult = _SmartCopyStyle($g_oWord, $oSource, $oTarget, $sName)
        
        If $iResult = 1 Then
            $iCopied += 1
        Else
            $iFailed += 1
        EndIf
    Next
    
    Local $sMsg = "Thanh cong: " & $iCopied & " styles"
    If $iFailed > 0 Then $sMsg &= @CRLF & "Loi: " & $iFailed
    
    _UpdateProgress("Hoan tat!")
    MsgBox($MB_ICONINFORMATION, "Ket qua", $sMsg)
EndFunc

; === FUNCTION: Show Style Selector (with Hotkey support) ===
; NOTE: Cac ham hotkey (_EditHotkeyInline, _OpenWordModifyStyleDialog, etc.)
; da duoc tach ra thanh module rieng (StyleHotkey.au3, HotkeyDialogs.au3)
Func _ShowStyleSelector()
    Local $oSource = _GetSelectedDoc($g_cboSourceDoc)
    If Not IsObj($oSource) Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon file nguon truoc!")
        Return
    EndIf
    
    ; FIX: Save file nguon truoc khi doc danh sach
    $oSource.Save()
    
    ; Doc hotkey da luu tu file INI
    Local $sIniFile = _GetHotkeyIniPath()
    
    ; Tao popup window - tang kich thuoc de chua them cot hotkey
    Local $hPopup = GUICreate("Chon Style de Copy", 780, 620, -1, -1, _
        BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xF5F5F5)
    
    GUICtrlCreateLabel("File nguon: " & $oSource.Name, 10, 10, 760, 20)
    GUICtrlSetFont(-1, 9, 600)
    
    ; ListView voi them cot Hotkey
    $g_listStyles = GUICtrlCreateListView("Ten Style|Loai|Font|Hotkey", 10, 35, 760, 380, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS))
    _GUICtrlListView_SetExtendedListViewStyle($g_listStyles, _
        BitOR($LVS_EX_CHECKBOXES, $LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
    _GUICtrlListView_SetColumnWidth($g_listStyles, 0, 300)
    _GUICtrlListView_SetColumnWidth($g_listStyles, 1, 80)
    _GUICtrlListView_SetColumnWidth($g_listStyles, 2, 250)
    _GUICtrlListView_SetColumnWidth($g_listStyles, 3, 100)
    
    ; Mang luu thong tin style [Name, Type, FontInfo, Hotkey]
    Local $aAllStyles[1][4]
    Local $iValidCount = 0
    
    ; FIX: Load Styles CO LOC - chi hien style dang dung hoac user-defined
    Local $oStyles = $oSource.Styles
    
    For $i = 1 To $oStyles.Count
        Local $oStyle = $oStyles.Item($i)
        If Not IsObj($oStyle) Then ContinueLoop
        
        ; FIX: Chi hien style InUse hoac khong phai BuiltIn (tranh hang nghin style rac)
        If $oStyle.InUse = True Or $oStyle.BuiltIn = False Then
            Local $sName = $oStyle.NameLocal
            If StringLeft($sName, 1) = "_" Then ContinueLoop
            
            Local $sType = ($oStyle.Type = 1) ? "Doan van" : "Ky tu"
            
            ; Lay thong tin font
            Local $sFont = ""
            If IsObj($oStyle.Font) Then
                $sFont = $oStyle.Font.Name & ", " & $oStyle.Font.Size & "pt"
                If $oStyle.Font.Bold Then $sFont &= ", Bold"
                If $oStyle.Font.Italic Then $sFont &= ", Italic"
            EndIf
            
            ; Doc hotkey da luu (neu co) - Su dung ham tu StyleHotkey.au3
            Local $sHotkey = _LoadHotkeyFromIni($sName, $sIniFile)
            
            ; Luu vao mang
            ReDim $aAllStyles[$iValidCount + 1][4]
            $aAllStyles[$iValidCount][0] = $sName
            $aAllStyles[$iValidCount][1] = $sType
            $aAllStyles[$iValidCount][2] = $sFont
            $aAllStyles[$iValidCount][3] = $sHotkey
            
            ; Them vao ListView
            GUICtrlCreateListViewItem($sName & "|" & $sType & "|" & $sFont & "|" & $sHotkey, $g_listStyles)
            $iValidCount += 1
        EndIf
    Next
    
    GUICtrlCreateLabel("Tim thay " & $iValidCount & " styles", 10, 420, 150, 20)
    
    ; === PHAN GAN PHIM TAT ===
    GUICtrlCreateGroup(" Gan phim tat ", 15, 440, 750, 55)
    GUICtrlCreateLabel("Double-click vao cot Hotkey de nhap phim tat, hoac:", 25, 463, 280, 20)
    $g_btnOpenModifyStyle = GUICtrlCreateButton("Mo Modify Style", 310, 458, 120, 28)
    GUICtrlSetBkColor(-1, 0x3498DB)
    GUICtrlSetTip(-1, "Mo dialog Modify Style cua Word de gan phim tat")
    $g_btnRefreshHotkeys = GUICtrlCreateButton("Cap nhat", 440, 458, 80, 28)
    GUICtrlSetBkColor(-1, 0xF39C12)
    Local $btnEditHotkey = GUICtrlCreateButton("Nhap phim tat", 530, 458, 100, 28)
    GUICtrlSetBkColor(-1, 0x27AE60)
    GUICtrlSetTip(-1, "Chon style va nhan de nhap phim tat")
    Local $btnClearHotkey = GUICtrlCreateButton("Xoa", 640, 458, 60, 28)
    GUICtrlSetBkColor(-1, 0xE74C3C)
    GUICtrlSetTip(-1, "Xoa phim tat cua style da chon")
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; === NUT CHUC NANG ===
    Local $btnSelectAll = GUICtrlCreateButton("Chon tat ca", 10, 505, 90, 35)
    Local $btnDeselectAll = GUICtrlCreateButton("Bo chon", 110, 505, 90, 35)
    
    ; Checkbox tuy chon
    Local $chkPageSetup = GUICtrlCreateCheckbox(" Page Setup", 220, 510, 100, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    Local $chkHeaderFooter = GUICtrlCreateCheckbox(" Header/Footer", 330, 510, 120, 22)
    
    Local $btnSaveHotkeys = GUICtrlCreateButton("Luu Hotkeys", 530, 500, 100, 40)
    GUICtrlSetBkColor(-1, 0x9B59B6)
    GUICtrlSetFont(-1, 9, 600)
    
    Local $btnCopy = GUICtrlCreateButton("COPY", 530, 555, 120, 45, $BS_DEFPUSHBUTTON)
    GUICtrlSetFont(-1, 11, 700)
    GUICtrlSetBkColor(-1, 0x27AE60)
    Local $btnCancel = GUICtrlCreateButton("Dong", 660, 555, 100, 45)
    
    GUISetState(@SW_SHOW, $hPopup)
    
    ; Event loop
    While 1
        Local $iMsg = GUIGetMsg()
        
        ; Xu ly double-click vao ListView de edit hotkey
        If $iMsg = $g_listStyles Then
            Local $aInfo = GUIGetCursorInfo($hPopup)
            If $aInfo[4] <> 0 Then ; Co click
                Local $aHit = _GUICtrlListView_SubItemHitTest($g_listStyles)
                If $aHit[0] <> -1 And $aHit[1] = 3 Then ; Click vao cot Hotkey (column 3)
                    ; Goi ham tu HotkeyDialogs.au3
                    _EditHotkeyInline($g_listStyles, $aHit[0], $aAllStyles, $sIniFile, $hPopup)
                EndIf
            EndIf
        EndIf
        
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                ExitLoop
                
            Case $btnSelectAll
                _GUICtrlListView_SetItemChecked($g_listStyles, -1, True)
                
            Case $btnDeselectAll
                _GUICtrlListView_SetItemChecked($g_listStyles, -1, False)
                
            Case $btnEditHotkey
                ; Nhap phim tat cho style da chon
                Local $iSelIdx = _GUICtrlListView_GetSelectedIndices($g_listStyles, False)
                If $iSelIdx = "" Or $iSelIdx = -1 Then
                    MsgBox($MB_ICONWARNING, "Chua chon", "Vui long chon 1 style trong danh sach!", 0, $hPopup)
                    ContinueLoop
                EndIf
                ; Goi ham tu HotkeyDialogs.au3
                _EditHotkeyInline($g_listStyles, $iSelIdx, $aAllStyles, $sIniFile, $hPopup)
                
            Case $btnClearHotkey
                ; Xoa phim tat cua style da chon
                Local $iDelIdx = _GUICtrlListView_GetSelectedIndices($g_listStyles, False)
                If $iDelIdx = "" Or $iDelIdx = -1 Then
                    MsgBox($MB_ICONWARNING, "Chua chon", "Vui long chon 1 style de xoa hotkey!", 0, $hPopup)
                    ContinueLoop
                EndIf
                _GUICtrlListView_SetItemText($g_listStyles, $iDelIdx, "", 3)
                $aAllStyles[$iDelIdx][3] = ""
                IniDelete($sIniFile, "Hotkeys", $aAllStyles[$iDelIdx][0])
                _UpdateProgress("Da xoa phim tat cua style: " & $aAllStyles[$iDelIdx][0])
                
            Case $g_btnOpenModifyStyle
                ; Lay style duoc chon
                Local $iSelected = _GUICtrlListView_GetSelectedIndices($g_listStyles, False)
                If $iSelected = "" Or $iSelected = -1 Then
                    MsgBox($MB_ICONWARNING, "Chua chon", "Vui long chon 1 style trong danh sach!", 0, $hPopup)
                    ContinueLoop
                EndIf
                
                Local $sSelectedStyle = _GUICtrlListView_GetItemText($g_listStyles, $iSelected, 0)
                If $sSelectedStyle = "" Then ContinueLoop
                
                ; Mo Word's Modify Style dialog - Goi ham tu StyleHotkey.au3
                If _OpenWordModifyStyleDialog($oSource, $sSelectedStyle) Then
                    ; Thong bao huong dan
                    MsgBox($MB_ICONINFORMATION, "Huong dan", _
                        "Dialog 'Modify Style' da mo trong Word!" & @CRLF & @CRLF & _
                        "HUONG DAN GAN PHIM TAT:" & @CRLF & _
                        "1. Trong dialog 'Modify Style', click 'Format' (goc duoi ben trai)" & @CRLF & _
                        "2. Chon 'Shortcut key...' tu menu Format" & @CRLF & _
                        "3. Nhan to hop phim ban muon (VD: Ctrl+1, Alt+H...)" & @CRLF & _
                        "4. Click 'Assign' de gan phim tat" & @CRLF & _
                        "5. Click 'Close' roi 'OK' de hoan thanh" & @CRLF & @CRLF & _
                        "Sau do nhan 'Cap nhat' de cap nhat danh sach hotkey!", 0, $hPopup)
                Else
                    MsgBox($MB_ICONERROR, "Loi", "Khong the mo dialog Modify Style!" & @CRLF & _
                        "Kiem tra ket noi Word va style '" & $sSelectedStyle & "'", 0, $hPopup)
                EndIf
                
            Case $g_btnRefreshHotkeys
                ; Cap nhat lai danh sach hotkey tu Word - Goi ham tu StyleHotkey.au3
                _RefreshHotkeyListFromWord($oSource, $g_listStyles, $aAllStyles, $iValidCount, $sIniFile)
                MsgBox($MB_ICONINFORMATION, "Da cap nhat", "Da cap nhat danh sach hotkey tu Word!", 0, $hPopup)
                
            Case $btnSaveHotkeys
                ; Luu tat ca hotkeys vao file INI
                Local $iSaved = 0
                For $j = 0 To $iValidCount - 1
                    Local $sStyleName = $aAllStyles[$j][0]
                    Local $sHk = $aAllStyles[$j][3]
                    If $sHk <> "" Then
                        _SaveHotkeyToIni($sStyleName, $sHk, $sIniFile)
                        $iSaved += 1
                    Else
                        IniDelete($sIniFile, "Hotkeys", $sStyleName)
                    EndIf
                Next
                MsgBox($MB_ICONINFORMATION, "Da luu", "Da luu " & $iSaved & " phim tat vao file:" & @CRLF & $sIniFile, 0, $hPopup)
                
            Case $btnCopy
                ; Dem so style da chon
                Local $aSelectedStyles[1]
                Local $aSelectedHotkeys[1]
                Local $iSelected = 0
                For $j = 0 To $iValidCount - 1
                    If _GUICtrlListView_GetItemChecked($g_listStyles, $j) Then
                        ReDim $aSelectedStyles[$iSelected + 1]
                        ReDim $aSelectedHotkeys[$iSelected + 1]
                        $aSelectedStyles[$iSelected] = $aAllStyles[$j][0]
                        $aSelectedHotkeys[$iSelected] = $aAllStyles[$j][3]
                        $iSelected += 1
                    EndIf
                Next

                If $iSelected = 0 Then
                    MsgBox($MB_ICONWARNING, "Chua chon", "Vui long chon it nhat 1 style de copy!", 0, $hPopup)
                    ContinueLoop
                EndIf

                ; Xac nhan
                If MsgBox($MB_YESNO, "Xac nhan", "Copy " & $iSelected & " styles da chon?", 0, $hPopup) <> $IDYES Then
                    ContinueLoop
                EndIf

                ; Luu hotkeys truoc khi copy
                For $j = 0 To $iValidCount - 1
                    Local $sStyleNm = $aAllStyles[$j][0]
                    Local $sHkVal = $aAllStyles[$j][3]
                    If $sHkVal <> "" Then
                        _SaveHotkeyToIni($sStyleNm, $sHkVal, $sIniFile)
                    EndIf
                Next

                ; Luu trang thai checkbox truoc khi xoa GUI
                Local $bCopyPageSetup = (GUICtrlRead($chkPageSetup) = $GUI_CHECKED)
                Local $bCopyHeaderFooter = (GUICtrlRead($chkHeaderFooter) = $GUI_CHECKED)
                
                ; Thuc hien copy
                GUISetState(@SW_HIDE, $hPopup)
                Local $oTarget = _GetSelectedDoc($g_cboTargetDoc)
                If Not IsObj($oTarget) Then
                    MsgBox($MB_ICONWARNING, "Loi", "Chon file dich!")
                    GUISetState(@SW_SHOW, $hPopup)
                    ContinueLoop
                EndIf
                
                ; Copy styles
                Local $iCopied = _CopySelectedStylesByName($oSource, $oTarget, $aSelectedStyles)

                ; Tu dong gan hotkeys cho cac style da copy - Luu vao Normal.dotm
                _UpdateProgress("Dang gan phim tat vao Normal.dotm...")
                ConsoleWrite(@CRLF & "=== BAT DAU GAN HOTKEYS VAO NORMAL.DOTM ===" & @CRLF)
                Local $iHotkeyApplied = 0
                For $j = 0 To $iSelected - 1
                    ConsoleWrite("- Kiem tra style " & ($j+1) & "/" & $iSelected & ": " & $aSelectedStyles[$j] & @CRLF)
                    If $aSelectedHotkeys[$j] <> "" Then
                        ConsoleWrite("  Hotkey: " & $aSelectedHotkeys[$j] & @CRLF)
                        ; Su dung ham moi: _SaveHotkeysToNormalDotm
                        If _SaveHotkeysToNormalDotm($aSelectedStyles[$j], $aSelectedHotkeys[$j]) Then
                            $iHotkeyApplied += 1
                            ConsoleWrite("  => Thanh cong!" & @CRLF)
                        Else
                            ConsoleWrite("  => That bai!" & @CRLF)
                        EndIf
                    Else
                        ConsoleWrite("  Khong co hotkey" & @CRLF)
                    EndIf
                Next
                ConsoleWrite("=== KET THUC GAN HOTKEYS: " & $iHotkeyApplied & "/" & $iSelected & " ===" & @CRLF & @CRLF)

                ; Copy them Page Setup neu chon
                If $bCopyPageSetup Then
                    _CopyPageSetup($oSource, $oTarget)
                EndIf
                If $bCopyHeaderFooter Then
                    _CopyHeaderFooter($oSource, $oTarget)
                EndIf
                
                ; LUU DOCUMENT SAU KHI HOAN TAT (Quan trong!)
                _UpdateProgress("Dang luu document...")
                $oTarget.Save()
                ConsoleWrite("=== DA LUU DOCUMENT ===" & @CRLF)

                GUIDelete($hPopup)

                ; Hien thi ket qua
                Local $sResult = "DA COPY THANH CONG!" & @CRLF & @CRLF
                $sResult &= "Styles da copy: " & $iCopied & "/" & $iSelected & @CRLF
                If $bCopyPageSetup Then $sResult &= "Page Setup: Da copy" & @CRLF
                If $bCopyHeaderFooter Then $sResult &= "Header/Footer: Da copy" & @CRLF
                ; Hien thi hotkeys da gan
                If $iHotkeyApplied > 0 Then
                    $sResult &= @CRLF & "PHIM TAT DA GAN VAO NORMAL.DOTM: " & $iHotkeyApplied & " hotkeys" & @CRLF
                    For $j = 0 To $iSelected - 1
                        If $aSelectedHotkeys[$j] <> "" Then
                            $sResult &= "  - " & $aSelectedStyles[$j] & " = " & $aSelectedHotkeys[$j] & @CRLF
                        EndIf
                    Next
                    $sResult &= @CRLF & "=> Phim tat se hoat dong voi MOI document co style nay!"
                    $sResult &= @CRLF & "=> Khong can copy hotkey cho tung file nua!"
                EndIf
                _LogPreview($sResult)
                MsgBox($MB_ICONINFORMATION, "Hoan thanh", "Da copy " & $iCopied & " styles va gan " & $iHotkeyApplied & " phim tat vao Normal.dotm!" & @CRLF & @CRLF & "Phim tat se hoat dong voi moi document co style nay!")
                Return
        EndSwitch
    WEnd
    
    GUIDelete($hPopup)
EndFunc

; === FUNCTION: Copy Selected Styles By Name ===
Func _CopySelectedStylesByName($oSource, $oTarget, $aStyleNames)
    If Not IsObj($oSource) Or Not IsObj($oTarget) Then Return 0
    If UBound($aStyleNames) = 0 Then Return 0

    Local $iCopied = 0
    Local $iTotal = UBound($aStyleNames)

    For $i = 0 To $iTotal - 1
        Local $sStyleName = $aStyleNames[$i]
        _UpdateProgress("Dang copy " & ($i + 1) & "/" & $iTotal & ": " & $sStyleName)

        ; SU DUNG _SmartCopyStyle (Ultimate Fix)
        If _SmartCopyStyle($g_oWord, $oSource, $oTarget, $sStyleName) = 1 Then
            $iCopied += 1
        EndIf
    Next

    _UpdateProgress("Da copy " & $iCopied & " styles!")
    Return $iCopied
EndFunc

; === FUNCTION: Copy Page Setup ===
Func _CopyPageSetup($oSource, $oTarget)
    If Not IsObj($oSource) Or Not IsObj($oTarget) Then Return
    $oTarget.PageSetup.LeftMargin = $oSource.PageSetup.LeftMargin
    $oTarget.PageSetup.RightMargin = $oSource.PageSetup.RightMargin
    $oTarget.PageSetup.TopMargin = $oSource.PageSetup.TopMargin
    $oTarget.PageSetup.BottomMargin = $oSource.PageSetup.BottomMargin
    $oTarget.PageSetup.PageWidth = $oSource.PageSetup.PageWidth
    $oTarget.PageSetup.PageHeight = $oSource.PageSetup.PageHeight
EndFunc

; === FUNCTION: Copy Header/Footer ===
Func _CopyHeaderFooter($oSource, $oTarget)
    If Not IsObj($oSource) Or Not IsObj($oTarget) Then Return

    For $i = 1 To $oSource.Sections.Count
        Local $oSrcSection = $oSource.Sections.Item($i)
        If $i <= $oTarget.Sections.Count Then
            Local $oTgtSection = $oTarget.Sections.Item($i)

            ; Copy Primary Header
            If IsObj($oSrcSection.Headers(1)) And IsObj($oTgtSection.Headers(1)) Then
                $oSrcSection.Headers(1).Range.Copy()
                $oTgtSection.Headers(1).Range.Delete()
                $oTgtSection.Headers(1).Range.PasteAndFormat($wdFormatOriginalFormatting)
            EndIf

            ; Copy Primary Footer
            If IsObj($oSrcSection.Footers(1)) And IsObj($oTgtSection.Footers(1)) Then
                $oSrcSection.Footers(1).Range.Copy()
                $oTgtSection.Footers(1).Range.Delete()
                $oTgtSection.Footers(1).Range.PasteAndFormat($wdFormatOriginalFormatting)
            EndIf

            ; Copy First Page Header/Footer (neu co)
            If $oSrcSection.PageSetup.DifferentFirstPageHeaderFooter Then
                $oTgtSection.PageSetup.DifferentFirstPageHeaderFooter = True
                If IsObj($oSrcSection.Headers(2)) And IsObj($oTgtSection.Headers(2)) Then
                    $oSrcSection.Headers(2).Range.Copy()
                    $oTgtSection.Headers(2).Range.Delete()
                    $oTgtSection.Headers(2).Range.PasteAndFormat($wdFormatOriginalFormatting)
                EndIf
                If IsObj($oSrcSection.Footers(2)) And IsObj($oTgtSection.Footers(2)) Then
                    $oSrcSection.Footers(2).Range.Copy()
                    $oTgtSection.Footers(2).Range.Delete()
                    $oTgtSection.Footers(2).Range.PasteAndFormat($wdFormatOriginalFormatting)
                EndIf
            EndIf
        EndIf
    Next
EndFunc

; === FUNCTION: Copy Selected Options ===
Func _CopySelectedStyles()
    Local $oSource = _GetSelectedDoc($g_cboSourceDoc)
    Local $oTarget = _GetSelectedDoc($g_cboTargetDoc)
    
    If Not IsObj($oSource) Or Not IsObj($oTarget) Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon file nguon va dich!")
        Return
    EndIf
    
    If $oSource.FullName = $oTarget.FullName Then
        MsgBox($MB_ICONWARNING, "Loi", "File nguon va dich phai khac nhau!")
        Return
    EndIf
    
    _UpdateProgress("Dang copy theo tuy chon...")
    Local $sResult = "=== KET QUA COPY THEO TUY CHON ===" & @CRLF & @CRLF
    
    ; Copy Page Setup
    If GUICtrlRead($g_chkCopyPageSetup) = $GUI_CHECKED Then
        _UpdateProgress("Dang copy Page Setup...")
        $oTarget.PageSetup.LeftMargin = $oSource.PageSetup.LeftMargin
        $oTarget.PageSetup.RightMargin = $oSource.PageSetup.RightMargin
        $oTarget.PageSetup.TopMargin = $oSource.PageSetup.TopMargin
        $oTarget.PageSetup.BottomMargin = $oSource.PageSetup.BottomMargin
        $oTarget.PageSetup.PageWidth = $oSource.PageSetup.PageWidth
        $oTarget.PageSetup.PageHeight = $oSource.PageSetup.PageHeight
        $sResult &= "[x] Page Setup: OK" & @CRLF
    Else
        $sResult &= "[ ] Page Setup: Bo qua" & @CRLF
    EndIf
    
    ; Copy Header/Footer (su dung PasteAndFormat de giu nguyen dinh dang)
    If GUICtrlRead($g_chkCopyHeaderFooter) = $GUI_CHECKED Then
        _UpdateProgress("Dang copy Header/Footer...")
        _CopyHeaderFooter($oSource, $oTarget)
        $sResult &= "[x] Header/Footer: OK (PasteAndFormat)" & @CRLF
    Else
        $sResult &= "[ ] Header/Footer: Bo qua" & @CRLF
    EndIf
    
    ; Copy Styles
    If GUICtrlRead($g_chkCopyStyles) = $GUI_CHECKED Then
        _UpdateProgress("Dang copy Styles...")
        Local $iCopied = _CopyDocStyles($oSource, $oTarget)
        $sResult &= "[x] Styles: " & $iCopied & " styles" & @CRLF
    Else
        $sResult &= "[ ] Styles: Bo qua" & @CRLF
    EndIf
    
    _LogPreview($sResult)
    _UpdateProgress("Da copy theo tuy chon!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da copy theo tuy chon!")
EndFunc

; === FUNCTION: Copy Doc Styles (OrganizerCopy) ===
Func _CopyDocStyles($oSource, $oTarget)
    If Not IsObj($oSource) Or Not IsObj($oTarget) Then Return 0

    Local $sSrcPath = $oSource.FullName
    Local $sTgtPath = $oTarget.FullName

    ; Kiem tra file da duoc luu chua
    If $sTgtPath = "" Or StringInStr($sTgtPath, "\") = 0 Then
        MsgBox($MB_ICONWARNING, "Loi", "File dich chua duoc Luu (Save)!" & @CRLF & _
            "Hay luu file ra o cung truoc khi copy style.")
        Return 0
    EndIf

    Local $oSrcStyles = $oSource.Styles
    If Not IsObj($oSrcStyles) Then Return 0

    Local $iCopied = 0, $iFailed = 0
    Local $iTotal = $oSrcStyles.Count
    Local Const $wdStyleTypeParagraph = 1
    Local Const $wdStyleTypeCharacter = 2

    ; Thu thap danh sach style can copy
    Local $aStylesToCopy[1]
    Local $iStyleCount = 0

    For $i = 1 To $iTotal
        Local $oSrcStyle = $oSrcStyles.Item($i)
        If Not IsObj($oSrcStyle) Then ContinueLoop

        ; Chi copy paragraph va character styles
        If $oSrcStyle.Type <> $wdStyleTypeParagraph And $oSrcStyle.Type <> $wdStyleTypeCharacter Then ContinueLoop

        ; Chi copy style dang su dung hoac khong phai built-in
        If Not $oSrcStyle.InUse And $oSrcStyle.BuiltIn Then ContinueLoop

        ReDim $aStylesToCopy[$iStyleCount + 1]
        $aStylesToCopy[$iStyleCount] = $oSrcStyle.NameLocal
        $iStyleCount += 1
    Next

    ; Copy tung style bang SmartCopyStyle
    For $i = 0 To $iStyleCount - 1
        Local $sStyleName = $aStylesToCopy[$i]
        If Mod($i, 10) = 0 Then
            _UpdateProgress("Dang copy style " & ($i + 1) & "/" & $iStyleCount)
        EndIf

        If _SmartCopyStyle($g_oWord, $oSource, $oTarget, $sStyleName) = 1 Then
            $iCopied += 1
        Else
            $iFailed += 1
        EndIf
    Next

    ; Set QuickStyle cho cac style quan trong
    _UpdateProgress("Dang cap nhat Quick Styles Gallery...")
    For $i = 0 To UBound($g_aImportantStyles) - 1
        _SetStyleQuickGallery($oTarget, $g_aImportantStyles[$i], True, 10 + $i)
    Next

    _UpdateProgress("Da copy " & $iCopied & "/" & $iStyleCount & " styles!")
    Return $iCopied
EndFunc

; === FUNCTION: Set Style Quick Gallery ===
Func _SetStyleQuickGallery($oDoc, $sStyleName, $bQuickStyle = True, $iPriority = 1)
    If Not IsObj($oDoc) Then Return False

    Local $oStyle = 0
    For $i = 1 To $oDoc.Styles.Count
        If $oDoc.Styles.Item($i).NameLocal = $sStyleName Then
            $oStyle = $oDoc.Styles.Item($i)
            ExitLoop
        EndIf
    Next

    If Not IsObj($oStyle) Then Return False

    $oStyle.QuickStyle = $bQuickStyle
    $oStyle.Priority = $iPriority

    Return True
EndFunc

; === FUNCTION: Preview Source Styles ===
Func _PreviewSourceStyles()
    Local $oSource = _GetSelectedDoc($g_cboSourceDoc)
    If Not IsObj($oSource) Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon file nguon!")
        Return
    EndIf
    
    Local $sMsg = "STYLES TRONG FILE NGUON:" & @CRLF & @CRLF
    Local $oStyles = $oSource.Styles
    Local $iCount = 0
    
    For $i = 1 To $oStyles.Count
        Local $oStyle = $oStyles.Item($i)
        If Not IsObj($oStyle) Then ContinueLoop
        
        ; Chi hien style InUse hoac user-defined
        If $oStyle.InUse = True Or $oStyle.BuiltIn = False Then
            Local $sName = $oStyle.NameLocal
            If StringLeft($sName, 1) <> "_" Then
                $sMsg &= "- " & $sName & @CRLF
                $iCount += 1
                If $iCount >= 40 Then
                    $sMsg &= "... va " & ($oStyles.Count - 40) & " styles khac"
                    ExitLoop
                EndIf
            EndIf
        EndIf
    Next
    
    _LogPreview($sMsg)
EndFunc

; === HELPER: Copy mot Style don le ===
Func _CopySingleStyle($oSource, $oTarget, $sStyleName)
    If Not IsObj($oSource) Or Not IsObj($oTarget) Then Return False
    If $oSource.FullName = $oTarget.FullName Then Return False
    
    ; SU DUNG _SmartCopyStyle (Ultimate Fix)
    Return (_SmartCopyStyle($g_oWord, $oSource, $oTarget, $sStyleName) = 1)
EndFunc

; === HELPER: Copy Heading styles (1-9) ===
Func _CopyHeadingStyles($oSource, $oTarget)
    Local $iCopied = 0
    For $i = 1 To 9
        If _CopySingleStyle($oSource, $oTarget, "Heading " & $i) Then
            $iCopied += 1
        EndIf
    Next
    Return $iCopied
EndFunc

; === HELPER: Copy TOC styles (1-9) ===
Func _CopyTOCStyles($oSource, $oTarget)
    Local $iCopied = 0
    For $i = 1 To 9
        If _CopySingleStyle($oSource, $oTarget, "TOC " & $i) Then
            $iCopied += 1
        EndIf
    Next
    Return $iCopied
EndFunc

; ============================================
; NOTE: Cac ham hotkey da duoc tach ra thanh module rieng
; Xem: Modules\StyleHotkey.au3 va GUI\HotkeyDialogs.au3
; 
; Cac ham da duoc tach:
; - _ApplyHotkeysToCurrentDoc()
; - _ApplyAllSavedHotkeys()
; - _ApplyStyleHotkeyViaWord()
; - _ParseHotkeyToWordKeys()
; - _ValidateHotkeyFormat()
; - _ConvertWordKeyToString()
; - _RefreshHotkeyListFromWord()
; - _OpenWordModifyStyleDialog()
; - _SaveHotkeyToIni()
; - _LoadHotkeyFromIni()
; - _LoadAllHotkeysFromIni()
; - _BackupHotkeysNow()
; - _RestoreHotkeysFromBackup()
; - _ExportHotkeysToText()
; - _ShowBackupHotkeysDialog()
; - _CreateBackupGroup()
; - _GetHotkeyFileInfo()
; - _LoadBackupList()
; - _EditHotkeyInline()
; - _CaptureHotkeyPress()
; - _IsKeyPressed()
; ============================================

; ============================================
; HOTKEY INTEGRATION WRAPPERS
; ============================================

; === FUNCTION: Refresh Hotkeys from Word (Wrapper for EventLoop) ===
Func _RefreshHotkeysFromWord()
    If Not _CheckConnection() Then Return False
    
    ; Get current source document
    Local $oSource = _GetSelectedDoc($g_cboSourceDoc)
    If Not IsObj($oSource) Then
        MsgBox($MB_ICONWARNING, "Loi", "Vui long chon file nguon!")
        Return False
    EndIf
    
    ; Note: This function needs to be called with proper context
    ; For now, show a message that user should use the button in Style Selector dialog
    MsgBox($MB_ICONINFORMATION, "Thong bao", _
        "Chuc nang nay hoat dong trong dialog 'Chon Style'." & @CRLF & _
        "Vui long:" & @CRLF & _
        "1. Click 'Chon Style de Copy'" & @CRLF & _
        "2. Trong dialog, click 'Cap nhat' de refresh hotkeys tu Word")
    
    Return True
EndFunc

; === FUNCTION: Open Modify Style Dialog (Wrapper for EventLoop) ===
Func _OpenModifyStyleDialog()
    If Not _CheckConnection() Then Return False
    
    ; Get current source document
    Local $oSource = _GetSelectedDoc($g_cboSourceDoc)
    If Not IsObj($oSource) Then
        MsgBox($MB_ICONWARNING, "Loi", "Vui long chon file nguon!")
        Return False
    EndIf
    
    ; Prompt user to select a style
    Local $sStyleName = InputBox("Mo Modify Style", _
        "Nhap ten style can chinh sua:" & @CRLF & _
        "(VD: Heading 1, Normal, Caption...)", _
        "Heading 1")
    
    If @error Or $sStyleName = "" Then Return False
    
    ; Open Word's Modify Style dialog
    Local $bResult = _OpenWordModifyStyleDialog($oSource, $sStyleName)
    
    If $bResult Then
        _UpdateProgress("Da mo dialog Modify Style cho: " & $sStyleName)
    Else
        MsgBox($MB_ICONWARNING, "Loi", "Khong the mo dialog Modify Style!" & @CRLF & _
            "Kiem tra ten style co dung khong.")
    EndIf
    
    Return $bResult
EndFunc
