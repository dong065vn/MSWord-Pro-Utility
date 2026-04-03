; ============================================
; PDFFIX.AU3 - Module Sua loi PDF
; ============================================

#include-once

; Sua vung chon
Func _FixSelectedText()
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then
        MsgBox($MB_ICONWARNING, "Loi", "Khong lay duoc Selection!")
        Return
    EndIf
    
    ; Type 1 = wdSelectionIP (insertion point, no selection)
    If $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon van ban can sua!")
        Return
    EndIf
    
    _UpdateProgress("Dang sua vung chon...")
    Local $sText = $oSel.Text
    $sText = _ApplyTextFixes($sText)
    $oSel.Text = $sText

    If GUICtrlRead($g_chkFixLineSpacing) = $GUI_CHECKED Then
        _FixLineSpacingRange($oSel.Range, 1.5)
    EndIf
    If GUICtrlRead($g_chkResetSpacing) = $GUI_CHECKED Then
        _FixLineSpacingRange($oSel.Range, 1.0)
    EndIf

    _LogPreview("DA SUA VUNG CHON:" & @CRLF & StringLeft($sText, 500))
    _UpdateProgress("Hoan thanh!")
EndFunc

; Sua toan bo
Func _FixAllDocument()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", "Sua toan bo? (Nen backup truoc)") <> $IDYES Then Return
    
    _UpdateProgress("Dang sua toan bo...")
    Local $oFind = $g_oDoc.Content.Find
    If IsObj($oFind) Then _ApplyFindReplaceFixes($oFind)
    If GUICtrlRead($g_chkRemoveFakeNumbering) = $GUI_CHECKED Then _RemoveFakeNumberingInRange($g_oDoc.Content)

    If GUICtrlRead($g_chkFixLineSpacing) = $GUI_CHECKED Then
        _FixLineSpacingRange($g_oDoc.Content, 1.5)
    EndIf
    If GUICtrlRead($g_chkResetSpacing) = $GUI_CHECKED Then
        _FixLineSpacingRange($g_oDoc.Content, 1.0)
    EndIf

    _LogPreview("DA SUA TOAN BO TAI LIEU!")
    _UpdateProgress("Hoan thanh!")
EndFunc

; Quick Fix
Func _QuickFixAll()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Quick Fix", "Sua nhanh tat ca? (Nen backup truoc)") <> $IDYES Then Return

    ; Check tat ca options
    GUICtrlSetState($g_chkLineBreaks, $GUI_CHECKED)
    GUICtrlSetState($g_chkExtraSpaces, $GUI_CHECKED)
    GUICtrlSetState($g_chkHyphenation, $GUI_CHECKED)
    GUICtrlSetState($g_chkSpecialChars, $GUI_CHECKED)
    GUICtrlSetState($g_chkParagraphs, $GUI_CHECKED)
    GUICtrlSetState($g_chkTabs, $GUI_CHECKED)
    GUICtrlSetState($g_chkFixQuotes, $GUI_CHECKED)
    GUICtrlSetState($g_chkFixLineSpacing, $GUI_CHECKED)
    GUICtrlSetState($g_chkRemoveEmptyLines, $GUI_CHECKED)
    GUICtrlSetState($g_chkFixSpacingBefore, $GUI_CHECKED)
    GUICtrlSetState($g_chkRemoveFakeNumbering, $GUI_UNCHECKED)

    _FixAllDocument()
EndFunc

; Apply text fixes (cho string)
Func _ApplyTextFixes($sText)
    If GUICtrlRead($g_chkLineBreaks) = $GUI_CHECKED Then
        $sText = StringReplace($sText, Chr(11), " ")
        $sText = StringReplace($sText, @LF, " ")
        $sText = StringReplace($sText, @CR & @CR, "{{PARA}}")
        $sText = StringReplace($sText, @CR, " ")
        $sText = StringReplace($sText, "{{PARA}}", @CRLF & @CRLF)
    EndIf
    
    If GUICtrlRead($g_chkExtraSpaces) = $GUI_CHECKED Then
        While StringInStr($sText, "  ")
            $sText = StringReplace($sText, "  ", " ")
        WEnd
        $sText = StringStripWS($sText, 3)
    EndIf
    
    If GUICtrlRead($g_chkHyphenation) = $GUI_CHECKED Then
        $sText = StringRegExpReplace($sText, "-\s*[\r\n]+\s*", "")
    EndIf
    
    If GUICtrlRead($g_chkSpecialChars) = $GUI_CHECKED Or GUICtrlRead($g_chkFixQuotes) = $GUI_CHECKED Then
        $sText = StringReplace($sText, ChrW(8220), '"')
        $sText = StringReplace($sText, ChrW(8221), '"')
        $sText = StringReplace($sText, ChrW(8216), "'")
        $sText = StringReplace($sText, ChrW(8217), "'")
        $sText = StringReplace($sText, ChrW(160), " ")
        $sText = StringReplace($sText, ChrW(8211), "-")
        $sText = StringReplace($sText, ChrW(8212), "-")
    EndIf

    If GUICtrlRead($g_chkRemoveFakeNumbering) = $GUI_CHECKED Then
        Local $aLines = StringSplit(StringReplace($sText, @CRLF, @LF), @LF, 1)
        If Not @error Then
            For $i = 1 To $aLines[0]
                $aLines[$i] = _StripUnnecessaryLeadingNumbering($aLines[$i])
            Next
            $sText = _ArrayToString($aLines, @CRLF, 1)
        EndIf
    EndIf
    
    Return $sText
EndFunc


; Apply Find/Replace fixes (cho document)
Func _ApplyFindReplaceFixes($oFind)
    If Not IsObj($oFind) Then Return

    If GUICtrlRead($g_chkLineBreaks) = $GUI_CHECKED Then
        _DoReplace($oFind, "^l", " ")
    EndIf

    If GUICtrlRead($g_chkExtraSpaces) = $GUI_CHECKED Then
        For $i = 1 To 5
            _DoReplace($oFind, "  ", " ")
        Next
        _DoReplace($oFind, "^p ", "^p")
        _DoReplace($oFind, " ^p", "^p")
    EndIf

    If GUICtrlRead($g_chkHyphenation) = $GUI_CHECKED Then
        _DoReplace($oFind, "- ^p", "")
        _DoReplace($oFind, "-^p", "")
    EndIf

    If GUICtrlRead($g_chkSpecialChars) = $GUI_CHECKED Or GUICtrlRead($g_chkFixQuotes) = $GUI_CHECKED Then
        _DoReplace($oFind, ChrW(8220), '"')
        _DoReplace($oFind, ChrW(8221), '"')
        _DoReplace($oFind, ChrW(8216), "'")
        _DoReplace($oFind, ChrW(8217), "'")
        _DoReplace($oFind, ChrW(160), " ")
    EndIf

    If GUICtrlRead($g_chkParagraphs) = $GUI_CHECKED Or GUICtrlRead($g_chkRemoveEmptyLines) = $GUI_CHECKED Then
        For $i = 1 To 10
            _DoReplace($oFind, "^p^p^p", "^p^p")
        Next
    EndIf

    If GUICtrlRead($g_chkTabs) = $GUI_CHECKED Then
        For $i = 1 To 3
            _DoReplace($oFind, "^t^t", "^t")
        Next
    EndIf
EndFunc

Func _CollapseExtraParagraphBreaks($oRange, $bCompact = True)
    If Not IsObj($oRange) Then Return

    Local $oFind = $oRange.Find
    If Not IsObj($oFind) Then Return

    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    $oFind.Forward = True
    $oFind.Wrap = 0
    $oFind.Format = False
    $oFind.MatchCase = False
    $oFind.MatchWholeWord = False
    $oFind.MatchWildcards = False
    $oFind.MatchSoundsLike = False
    $oFind.MatchAllWordForms = False

    Local $sFind = "^p^p"
    Local $sReplace = "^p"
    If Not $bCompact Then
        $sFind = "^p^p^p"
        $sReplace = "^p^p"
    EndIf

    For $i = 1 To 12
        $oFind.Text = $sFind
        $oFind.Replacement.Text = $sReplace
        $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
    Next
EndFunc

Func _StripUnnecessaryLeadingNumbering($sText)
    Local $sOriginal = $sText
    Local $sCore = StringRegExpReplace($sText, "[\r\x07]+$", "")
    If StringStripWS($sCore, 3) = "" Then Return $sOriginal

    If StringRegExp($sCore, "^\s*\d+(\.\d+)+([\.\)])?\s+") Then Return $sOriginal

    Local $sUpdated = StringRegExpReplace($sCore, "^\s*(\(?\d{1,3}[\.\)]|\(?[A-Za-z][\.\)])\s+(?=\S)", "", 1)
    If @extended = 0 Then Return $sOriginal

    Return $sUpdated & StringTrimLeft($sOriginal, StringLen($sCore))
EndFunc

Func _RemoveFakeNumberingInRange($oRange)
    If Not IsObj($oRange) Then Return 0

    Local $oParas = $oRange.Paragraphs
    If Not IsObj($oParas) Then Return 0

    Local $iRemoved = 0
    For $i = $oParas.Count To 1 Step -1
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        If $oPara.Range.ListFormat.ListType <> 0 Then ContinueLoop

        Local $sOriginal = $oPara.Range.Text
        Local $sStripped = _StripUnnecessaryLeadingNumbering($sOriginal)
        If $sStripped = $sOriginal Then ContinueLoop

        Local $oTextRange = $oPara.Range.Duplicate
        If Not IsObj($oTextRange) Then ContinueLoop

        Local $sTail = ""
        If StringLen($sOriginal) > 0 Then
            Local $sLastChar = StringRight($sOriginal, 1)
            If $sLastChar = @CR Or $sLastChar = Chr(7) Then
                $sTail = $sLastChar
                $oTextRange.End = $oTextRange.End - 1
            EndIf
        EndIf

        $oTextRange.Text = StringRegExpReplace($sStripped, "[\r\x07]+$", "") & $sTail
        $iRemoved += 1
    Next

    Return $iRemoved
EndFunc


; =========================================================
; HAM: _CleanUpDocument
; Tac dung:
; 1. Reset dinh dang (xoa in dam/nghieng/mau tay) nhung GIU NGUYEN STYLE
; 2. Xoa cac dong trong thua (2 Enter -> 1 Enter)
; Ky thuat: Font.Reset() + ParagraphFormat.Reset() + Wildcard Find/Replace
; =========================================================
Func _CleanUpDocument()
    If Not _CheckConnection() Then Return
    
    Local $iChoice = MsgBox($MB_YESNOCANCEL + $MB_ICONQUESTION, "Clean Up Document", _
        "Chuc nang nay se:" & @CRLF & @CRLF & _
        "1. Reset dinh dang (xoa mau, in dam tay...) nhung GIU NGUYEN STYLE" & @CRLF & _
        "   (Heading van la Heading, List van la List)" & @CRLF & @CRLF & _
        "2. Xoa dong trong thua (2+ Enter -> 1 Enter)" & @CRLF & @CRLF & _
        "Chon YES de xu ly TOAN BO tai lieu" & @CRLF & _
        "Chon NO de chi xu ly VUNG CHON")
    
    If $iChoice = $IDCANCEL Then Return
    
    Local $oRange
    Local $sScope = ""
    
    If $iChoice = $IDYES Then
        ; Xu ly toan bo document
        $oRange = $g_oDoc.Content
        $sScope = "TOAN BO TAI LIEU"
    Else
        ; Xu ly vung chon
        Local $oSel = $g_oWord.Selection
        If Not IsObj($oSel) Or $oSel.Type = 1 Then
            MsgBox($MB_ICONWARNING, "Loi", "Vui long chon vung van ban can xu ly!")
            Return
        EndIf
        $oRange = $oSel.Range
        $sScope = "VUNG CHON"
    EndIf
    
    _UpdateProgress("Dang Clean Up " & $sScope & "...")
    
    Local $sLog = "=== CLEAN UP DOCUMENT ===" & @CRLF
    $sLog &= "Pham vi: " & $sScope & @CRLF & @CRLF
    
    ; --- BUOC 1: RESET DINH DANG (GIU NGUYEN STYLE) ---
    ; Tuong duong boi den toan bo -> Bam Ctrl+Space (Reset Font) va Ctrl+Q (Reset Paragraph)
    
    _UpdateProgress("Buoc 1: Reset dinh dang Font...")
    
    ; 1.1. Dat lai dinh dang Font (ve mac dinh cua Style hien tai)
    ; Lenh nay xoa mau sac, in dam, font chu bi chinh tay...
    $oRange.Font.Reset()
    $sLog &= "[OK] Da Reset Font (xoa mau, in dam/nghieng tay)" & @CRLF
    
    _UpdateProgress("Buoc 2: Reset dinh dang Paragraph...")
    
    ; 1.2. Dat lai dinh dang Doan van (ve mac dinh cua Style hien tai)
    ; Lenh nay xoa thut dau dong thu cong, gian dong thu cong...
    $oRange.ParagraphFormat.Reset()
    $sLog &= "[OK] Da Reset Paragraph (xoa thut le, gian dong tay)" & @CRLF
    
    ; --- BUOC 2: XU LY DONG TRONG (2 ENTER VE 1) ---
    _UpdateProgress("Buoc 3: Xoa dong trong thua...")
    
    Local $oFind = $oRange.Find
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    
    _CollapseExtraParagraphBreaks($oRange, True)
    
    $sLog &= "[OK] Da xoa dong trong thua (2+ Enter -> 1 Enter)" & @CRLF
    
    ; --- BUOC 3: XOA KHOANG TRANG THUA ---
    _UpdateProgress("Buoc 4: Xoa khoang trang thua...")
    
    ; Reset Find object
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    $oFind.MatchWildcards = False
    
    ; Xoa khoang trang dau dong
    $oFind.Text = "^p "
    $oFind.Replacement.Text = "^p"
    $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
    
    ; Xoa khoang trang cuoi dong
    $oFind.Text = " ^p"
    $oFind.Replacement.Text = "^p"
    $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
    
    ; Xoa nhieu khoang trang lien tiep
    For $i = 1 To 5
        $oFind.Text = "  "
        $oFind.Replacement.Text = " "
        $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
    Next
    
    $sLog &= "[OK] Da xoa khoang trang thua" & @CRLF

    If GUICtrlRead($g_chkRemoveFakeNumbering) = $GUI_CHECKED Then
        _UpdateProgress("Buoc 5: Bo so dau dong thua...")
        Local $iRemoved = _RemoveFakeNumberingInRange($oRange)
        $sLog &= "[OK] Da bo so dau dong thua: " & $iRemoved & @CRLF
    EndIf
    
    ; --- HOAN TAT ---
    $sLog &= @CRLF & "=== HOAN TAT CLEAN UP ===" & @CRLF
    $sLog &= "Luu y: Style (Heading, List, Normal...) duoc giu nguyen!" & @CRLF
    
    _LogPreview($sLog)
    _UpdateProgress("Clean Up hoan tat!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da Clean Up " & $sScope & " thanh cong!" & @CRLF & @CRLF & _
        "- Reset dinh dang (giu Style)" & @CRLF & _
        "- Xoa dong trong thua" & @CRLF & _
        "- Xoa khoang trang thua")
EndFunc

; =========================================================
; HAM: _CleanUpDocumentAdvanced
; Phien ban nang cao voi tuy chon
; =========================================================
Func _CleanUpDocumentAdvanced()
    If Not _CheckConnection() Then Return
    
    ; Tao dialog tuy chon
    Local $hPopup = GUICreate("Clean Up - Tuy chon", 400, 300, -1, -1, _
        BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xF5F5F5)
    
    GUICtrlCreateLabel("CLEAN UP DOCUMENT", 15, 10, 370, 25)
    GUICtrlSetFont(-1, 12, 700)
    GUICtrlSetColor(-1, 0x2C3E50)
    
    ; Tuy chon Reset Format
    GUICtrlCreateGroup(" Reset dinh dang (giu Style) ", 15, 40, 370, 80)
    Local $chkResetFont = GUICtrlCreateCheckbox(" Reset Font (xoa mau, in dam/nghieng tay)", 25, 60, 350, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    Local $chkResetPara = GUICtrlCreateCheckbox(" Reset Paragraph (xoa thut le, gian dong tay)", 25, 85, 350, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; Tuy chon Xu ly dong trong
    GUICtrlCreateGroup(" Xu ly dong trong ", 15, 125, 370, 80)
    Local $chkRemoveEmpty = GUICtrlCreateCheckbox(" Xoa dong trong thua", 25, 145, 200, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    Local $radioCompact = GUICtrlCreateRadio(" Compact (2+ Enter -> 1)", 35, 168, 160, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    Local $radioKeepOne = GUICtrlCreateRadio(" Giu 1 dong trong (3+ -> 2)", 200, 168, 170, 22)
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; Tuy chon bo sung
    GUICtrlCreateGroup(" Bo sung ", 15, 210, 370, 40)
    Local $chkRemoveSpaces = GUICtrlCreateCheckbox(" Xoa khoang trang thua", 25, 228, 200, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; Buttons
    Local $btnApplyAll = GUICtrlCreateButton("Ap dung TOAN BO", 15, 260, 120, 30)
    GUICtrlSetBkColor(-1, 0x27AE60)
    GUICtrlSetFont(-1, 9, 600)
    Local $btnApplySel = GUICtrlCreateButton("Ap dung VUNG CHON", 145, 260, 130, 30)
    GUICtrlSetBkColor(-1, 0x3498DB)
    GUICtrlSetFont(-1, 9, 600)
    Local $btnCancel = GUICtrlCreateButton("Huy", 320, 260, 65, 30)
    
    GUISetState(@SW_SHOW, $hPopup)
    
    While 1
        Local $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                GUIDelete($hPopup)
                Return
                
            Case $btnApplyAll, $btnApplySel
                ; Lay tuy chon
                Local $bResetFont = (GUICtrlRead($chkResetFont) = $GUI_CHECKED)
                Local $bResetPara = (GUICtrlRead($chkResetPara) = $GUI_CHECKED)
                Local $bRemoveEmpty = (GUICtrlRead($chkRemoveEmpty) = $GUI_CHECKED)
                Local $bCompact = (GUICtrlRead($radioCompact) = $GUI_CHECKED)
                Local $bRemoveSpaces = (GUICtrlRead($chkRemoveSpaces) = $GUI_CHECKED)
                Local $bApplyAll = ($iMsg = $btnApplyAll)
                
                GUIDelete($hPopup)
                
                ; Thuc hien Clean Up
                _DoCleanUpWithOptions($bApplyAll, $bResetFont, $bResetPara, $bRemoveEmpty, $bCompact, $bRemoveSpaces)
                Return
        EndSwitch
    WEnd
EndFunc

; Ham thuc hien Clean Up voi tuy chon
Func _DoCleanUpWithOptions($bApplyAll, $bResetFont, $bResetPara, $bRemoveEmpty, $bCompact, $bRemoveSpaces)
    Local $oRange
    Local $sScope = ""
    
    If $bApplyAll Then
        $oRange = $g_oDoc.Content
        $sScope = "TOAN BO"
    Else
        Local $oSel = $g_oWord.Selection
        If Not IsObj($oSel) Or $oSel.Type = 1 Then
            MsgBox($MB_ICONWARNING, "Loi", "Vui long chon vung van ban!")
            Return
        EndIf
        $oRange = $oSel.Range
        $sScope = "VUNG CHON"
    EndIf
    
    _UpdateProgress("Dang Clean Up...")
    Local $sLog = "=== CLEAN UP (" & $sScope & ") ===" & @CRLF
    
    ; Reset Font
    If $bResetFont Then
        $oRange.Font.Reset()
        $sLog &= "[OK] Reset Font" & @CRLF
    EndIf
    
    ; Reset Paragraph
    If $bResetPara Then
        $oRange.ParagraphFormat.Reset()
        $sLog &= "[OK] Reset Paragraph" & @CRLF
    EndIf
    
    ; Xu ly dong trong
    If $bRemoveEmpty Then
        _CollapseExtraParagraphBreaks($oRange, $bCompact)
        $sLog &= "[OK] Xu ly dong trong" & @CRLF
    EndIf
    
    ; Xoa khoang trang thua
    If $bRemoveSpaces Then
        Local $oFind2 = $oRange.Find
        $oFind2.ClearFormatting()
        $oFind2.Replacement.ClearFormatting()
        $oFind2.MatchWildcards = False
        
        $oFind2.Text = "^p "
        $oFind2.Replacement.Text = "^p"
        $oFind2.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
        
        $oFind2.Text = " ^p"
        $oFind2.Replacement.Text = "^p"
        $oFind2.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
        
        For $i = 1 To 5
            $oFind2.Text = "  "
            $oFind2.Replacement.Text = " "
            $oFind2.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
        Next
        $sLog &= "[OK] Xoa khoang trang thua" & @CRLF
    EndIf

    If GUICtrlRead($g_chkRemoveFakeNumbering) = $GUI_CHECKED Then
        Local $iRemoved = _RemoveFakeNumberingInRange($oRange)
        $sLog &= "[OK] Bo so dau dong thua: " & $iRemoved & @CRLF
    EndIf
    
    $sLog &= @CRLF & "=== HOAN TAT ===" & @CRLF
    _LogPreview($sLog)
    _UpdateProgress("Clean Up hoan tat!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da Clean Up thanh cong!")
EndFunc


; ==============================================================================
; MODULE FIX LAYOUT
; 1. Fix loi Justify bi gian chu (Thay Shift+Enter bang Enter)
; 2. Fix loi Bang dinh chu/cach xa (Chuyen ve Inline va set khoang cach chuan)
; ==============================================================================
Func _FixLayoutProblems()
    If Not _CheckConnection() Then Return
    
    Local $iChoice = MsgBox($MB_YESNOCANCEL + $MB_ICONQUESTION, "Fix Layout Problems", _
        "Chuc nang nay se sua:" & @CRLF & @CRLF & _
        "1. LOI JUSTIFY BI GIAN CHU" & @CRLF & _
        "   (Do Shift+Enter -> Thay bang Enter)" & @CRLF & @CRLF & _
        "2. LOI BANG (TABLE) DINH CHU / CACH XA" & @CRLF & _
        "   (Chuyen ve Inline, set khoang cach chuan 12pt)" & @CRLF & @CRLF & _
        "Chon YES de xu ly TAT CA" & @CRLF & _
        "Chon NO de chi xu ly LOI JUSTIFY")
    
    If $iChoice = $IDCANCEL Then Return
    
    _UpdateProgress("Dang Fix Layout...")
    Local $sLog = "=== FIX LAYOUT PROBLEMS ===" & @CRLF & @CRLF
    
    ; --- PHAN 1: FIX LOI KHOANG TRANG DO JUSTIFY ---
    ; Nguyen ly: Tim ky tu ngat dong thu cong (^l) thay bang ngat doan (^p)
    _UpdateProgress("Dang fix loi Justify (Shift+Enter)...")
    
    Local $oFind = $g_oDoc.Content.Find
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    
    $oFind.Text = "^l" ; ^l la Shift+Enter (Manual Line Break)
    $oFind.Replacement.Text = "^p" ; ^p la Enter (Paragraph Mark)
    $oFind.Forward = True
    $oFind.Wrap = 1 ; wdFindContinue
    $oFind.Format = False
    $oFind.MatchCase = False
    $oFind.MatchWholeWord = False
    $oFind.MatchWildcards = False
    
    ; wdReplaceAll = 2
    $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
    
    $sLog &= "[OK] Da xu ly loi Justify (Shift+Enter -> Enter)" & @CRLF
    $sLog &= "     Nguyen nhan: Dong ket thuc bang Shift+Enter bi Word keo gian ra 2 le" & @CRLF & @CRLF
    
    ; --- PHAN 2: FIX LOI KHOANG CACH BANG (TABLE SPACING) ---
    If $iChoice = $IDYES Then
        _UpdateProgress("Dang chuan hoa cac Bang (Tables)...")
        
        Local $oTables = $g_oDoc.Tables
        Local $iCount = $oTables.Count
        
        If $iCount > 0 Then
            $sLog &= "Tim thay " & $iCount & " bang trong tai lieu." & @CRLF & @CRLF
            
            For $i = 1 To $iCount
                Local $oTbl = $oTables.Item($i)
                
                _UpdateProgress("Dang xu ly bang " & $i & "/" & $iCount & "...")
                
                ; 2.1. Tat che do "Troi noi" (Quan trong nhat)
                ; Chuyen Text Wrapping ve 'None' (Inline) de bang nam yen, khong de chu
                $oTbl.Rows.WrapAroundText = False
                
                ; 2.2. Can giua bang (Thuong bang nen can giua trang)
                $oTbl.Rows.Alignment = 1 ; 0=Left, 1=Center, 2=Right
                
                ; 2.3. Thiet lap khoang cach an toan voi van ban
                ; Thay vi chinh Table Properties, ta chinh Paragraph bao chua bang
                Local $oRange = $oTbl.Range
                $oRange.ParagraphFormat.SpaceBefore = 12 ; Cach doan tren 12pt
                $oRange.ParagraphFormat.SpaceAfter = 12  ; Cach doan duoi 12pt
                $oRange.ParagraphFormat.LineSpacingRule = 0 ; wdLineSpaceSingle
                
                ; 2.4. AutoFit de bang khong bi tran le
                $oTbl.AutoFitBehavior(2) ; wdAutoFitWindow (Tu co gian theo kho giay)
                
                $sLog &= "[OK] Bang " & $i & ": Wrap=None, Align=Center, Spacing=12pt" & @CRLF
            Next
            
            $sLog &= @CRLF & "Da chuan hoa " & $iCount & " bang thanh cong!" & @CRLF
        Else
            $sLog &= "[INFO] Khong tim thay bang nao trong tai lieu." & @CRLF
        EndIf
    Else
        $sLog &= "[SKIP] Bo qua xu ly Bang (chi fix Justify)" & @CRLF
    EndIf
    
    $sLog &= @CRLF & "=== HOAN TAT FIX LAYOUT ===" & @CRLF
    
    _LogPreview($sLog)
    _UpdateProgress("Fix Layout hoan tat!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da Fix Layout thanh cong!" & @CRLF & @CRLF & _
        "- Sua loi Justify bi gian chu" & @CRLF & _
        (($iChoice = $IDYES) ? "- Chuan hoa khoang cach cac Bang" : ""))
EndFunc

; ==============================================================================
; HAM: _FixTableSpacing
; Tac dung: Chi xu ly khoang cach Bang (khong fix Justify)
; ==============================================================================
Func _FixTableSpacing()
    If Not _CheckConnection() Then Return
    
    Local $oTables = $g_oDoc.Tables
    Local $iCount = $oTables.Count
    
    If $iCount = 0 Then
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong tim thay bang nao trong tai lieu!")
        Return
    EndIf
    
    ; Hien thi dialog tuy chon
    Local $hPopup = GUICreate("Fix Table Spacing", 400, 280, -1, -1, _
        BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xF5F5F5)
    
    GUICtrlCreateLabel("CHUAN HOA KHOANG CACH BANG", 15, 10, 370, 25)
    GUICtrlSetFont(-1, 11, 700)
    GUICtrlSetColor(-1, 0x2C3E50)
    
    GUICtrlCreateLabel("Tim thay " & $iCount & " bang trong tai lieu.", 15, 40, 370, 20)
    
    ; Tuy chon
    GUICtrlCreateGroup(" Tuy chon xu ly ", 15, 65, 370, 120)
    
    Local $chkWrapNone = GUICtrlCreateCheckbox(" Tat che do troi noi (WrapAroundText = False)", 25, 85, 350, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    GUICtrlSetTip(-1, "Quan trong! Giup bang khong de chu hoac cach xa bat thuong")
    
    Local $chkAlignCenter = GUICtrlCreateCheckbox(" Can giua bang", 25, 110, 200, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    
    Local $chkAutoFit = GUICtrlCreateCheckbox(" AutoFit theo kho giay", 230, 110, 150, 22)
    GUICtrlSetState(-1, $GUI_CHECKED)
    
    GUICtrlCreateLabel("Khoang cach tren/duoi (pt):", 25, 138, 150, 20)
    Local $inputSpacing = GUICtrlCreateInput("12", 180, 135, 50, 22)
    
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; Buttons
    Local $btnApply = GUICtrlCreateButton("Ap dung", 100, 200, 100, 35)
    GUICtrlSetBkColor(-1, 0x27AE60)
    GUICtrlSetFont(-1, 10, 600)
    Local $btnCancel = GUICtrlCreateButton("Huy", 210, 200, 80, 35)
    
    GUISetState(@SW_SHOW, $hPopup)
    
    While 1
        Local $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                GUIDelete($hPopup)
                Return
                
            Case $btnApply
                Local $bWrapNone = (GUICtrlRead($chkWrapNone) = $GUI_CHECKED)
                Local $bAlignCenter = (GUICtrlRead($chkAlignCenter) = $GUI_CHECKED)
                Local $bAutoFit = (GUICtrlRead($chkAutoFit) = $GUI_CHECKED)
                Local $iSpacing = Int(GUICtrlRead($inputSpacing))
                If $iSpacing < 0 Or $iSpacing > 72 Then $iSpacing = 12
                
                GUIDelete($hPopup)
                
                ; Thuc hien xu ly
                _DoFixTableSpacing($bWrapNone, $bAlignCenter, $bAutoFit, $iSpacing)
                Return
        EndSwitch
    WEnd
EndFunc

; Ham thuc hien fix table spacing
Func _DoFixTableSpacing($bWrapNone, $bAlignCenter, $bAutoFit, $iSpacing)
    _UpdateProgress("Dang xu ly cac Bang...")
    
    Local $oTables = $g_oDoc.Tables
    Local $iCount = $oTables.Count
    Local $sLog = "=== FIX TABLE SPACING ===" & @CRLF & @CRLF
    
    For $i = 1 To $iCount
        Local $oTbl = $oTables.Item($i)
        _UpdateProgress("Dang xu ly bang " & $i & "/" & $iCount & "...")
        
        If $bWrapNone Then
            $oTbl.Rows.WrapAroundText = False
        EndIf
        
        If $bAlignCenter Then
            $oTbl.Rows.Alignment = 1
        EndIf
        
        If $bAutoFit Then
            $oTbl.AutoFitBehavior(2)
        EndIf
        
        ; Set spacing
        Local $oRange = $oTbl.Range
        $oRange.ParagraphFormat.SpaceBefore = $iSpacing
        $oRange.ParagraphFormat.SpaceAfter = $iSpacing
        
        $sLog &= "[OK] Bang " & $i & @CRLF
    Next
    
    $sLog &= @CRLF & "Da xu ly " & $iCount & " bang!" & @CRLF
    $sLog &= "Spacing: " & $iSpacing & "pt" & @CRLF
    
    _LogPreview($sLog)
    _UpdateProgress("Hoan tat!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da chuan hoa " & $iCount & " bang!")
EndFunc

; ==============================================================================
; HAM: _FixJustifyOnly
; Tac dung: Chi fix loi Justify (Shift+Enter -> Enter)
; ==============================================================================
Func _FixJustifyOnly()
    If Not _CheckConnection() Then Return
    
    _UpdateProgress("Dang fix loi Justify...")
    
    Local $oFind = $g_oDoc.Content.Find
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    
    $oFind.Text = "^l"
    $oFind.Replacement.Text = "^p"
    $oFind.Forward = True
    $oFind.Wrap = 1
    $oFind.MatchWildcards = False
    
    $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
    
    _UpdateProgress("Hoan tat!")
    _LogPreview("Da thay the tat ca Shift+Enter (^l) bang Enter (^p)" & @CRLF & _
        "Loi gian chu khi Justify da duoc sua!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da fix loi Justify (gian chu)!")
EndFunc


; ==============================================================================
; HAM: _ShowPDFFixHelp
; Tac dung: Hien thi huong dan su dung cac chuc nang trong tab Sua loi PDF
; ==============================================================================
Func _ShowPDFFixHelp()
    Local $hHelp = GUICreate("Huong dan Sua loi PDF", 700, 580, -1, -1, _
        BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xFFFFFF)
    
    ; Tieu de
    GUICtrlCreateLabel("HUONG DAN SUA LOI PDF TO WORD", 20, 15, 660, 30)
    GUICtrlSetFont(-1, 14, 700)
    GUICtrlSetColor(-1, 0x2C3E50)
    
    ; Noi dung huong dan
    Local $sHelp = ""
    $sHelp &= "=== CAC NUT CHUC NANG ===" & @CRLF & @CRLF
    
    $sHelp &= "1. SUA VUNG CHON" & @CRLF
    $sHelp &= "   - Boi den (select) phan van ban can sua trong Word" & @CRLF
    $sHelp &= "   - Nhan nut nay de chi sua phan da chon" & @CRLF & @CRLF
    
    $sHelp &= "2. SUA TOAN BO" & @CRLF
    $sHelp &= "   - Sua tat ca van ban trong file Word" & @CRLF
    $sHelp &= "   - Nen BACKUP truoc khi dung!" & @CRLF & @CRLF
    
    $sHelp &= "3. QUICK FIX (Xanh la)" & @CRLF
    $sHelp &= "   - Tick tat ca cac tuy chon va sua nhanh toan bo" & @CRLF
    $sHelp &= "   - Danh cho truong hop can sua nhanh, khong can tuy chinh" & @CRLF & @CRLF
    
    $sHelp &= "4. CLEAN UP (Tim)" & @CRLF
    $sHelp &= "   - Reset dinh dang (xoa mau, in dam tay...) NHUNG GIU NGUYEN STYLE" & @CRLF
    $sHelp &= "   - Heading van la Heading, List van la List" & @CRLF
    $sHelp &= "   - Xoa dong trong thua (2+ Enter -> 1 Enter)" & @CRLF
    $sHelp &= "   - Ky thuat: Font.Reset() + ParagraphFormat.Reset()" & @CRLF & @CRLF
    
    $sHelp &= "5. FIX LAYOUT (Xanh duong)" & @CRLF
    $sHelp &= "   a) Fix loi JUSTIFY bi gian chu:" & @CRLF
    $sHelp &= "      - Nguyen nhan: Dong ket thuc bang Shift+Enter (^l)" & @CRLF
    $sHelp &= "      - Word co keo gian dong do ra 2 le -> Khoang trang lon" & @CRLF
    $sHelp &= "      - Fix: Thay Shift+Enter bang Enter (^p)" & @CRLF & @CRLF
    $sHelp &= "   b) Fix loi BANG (Table) dinh chu hoac cach xa:" & @CRLF
    $sHelp &= "      - Nguyen nhan: Bang dang o che do 'Troi noi' (WrapAroundText)" & @CRLF
    $sHelp &= "      - Fix: Chuyen ve Inline, set khoang cach chuan 12pt" & @CRLF & @CRLF
    
    $sHelp &= "6. UNDO" & @CRLF
    $sHelp &= "   - Hoan tac thao tac vua thuc hien (Ctrl+Z)" & @CRLF & @CRLF
    
    $sHelp &= "=== CAC TUY CHON TICK ===" & @CRLF & @CRLF
    
    $sHelp &= "LOI PHO BIEN TU PDF:" & @CRLF
    $sHelp &= "- Xoa xuong dong thua (^l): Xoa Manual Line Break" & @CRLF
    $sHelp &= "- Xoa khoang trang thua: Gop nhieu dau cach thanh 1" & @CRLF
    $sHelp &= "- Noi tu bi ngat (hyphen): Noi lai tu bi ngat dong (vi-" & @CRLF
    $sHelp &= "- Sua ky tu dac biet: Thay dau ngoac kep cong, gach ngang dai..." & @CRLF
    $sHelp &= "- Chuan hoa doan van: Xoa dong trong thua giua cac doan" & @CRLF
    $sHelp &= "- Xoa tab thua: Gop nhieu Tab thanh 1" & @CRLF
    $sHelp &= "- Bo so dau dong thua: Xoa marker kieu 1. / a. o dau dong neu khong can" & @CRLF & @CRLF
    
    $sHelp &= "XU LY CACH DONG:" & @CRLF
    $sHelp &= "- Fix cach dong (1.5): Dat gian dong 1.5 (chuan luan van)" & @CRLF
    $sHelp &= "- Reset cach dong (1.0): Dat gian dong don" & @CRLF
    $sHelp &= "- Xoa dong trong thua: 2+ dong trong -> 1 dong trong" & @CRLF
    $sHelp &= "- Xoa spacing thua: Xoa Space Before/After thua" & @CRLF & @CRLF
    
    $sHelp &= "=== LUU Y QUAN TRONG ===" & @CRLF & @CRLF
    $sHelp &= "1. LUON BACKUP file truoc khi sua!" & @CRLF
    $sHelp &= "2. Nen thu tren VUNG CHON truoc, ok roi moi sua TOAN BO" & @CRLF
    $sHelp &= "3. Neu sai, nhan UNDO hoac Ctrl+Z trong Word" & @CRLF
    $sHelp &= "4. CLEAN UP giu nguyen Style, chi xoa dinh dang tay" & @CRLF
    $sHelp &= "5. FIX LAYOUT danh cho loi Justify va Bang" & @CRLF
    
    ; Edit box hien thi huong dan
    Local $editHelp = GUICtrlCreateEdit($sHelp, 20, 50, 660, 480, _
        BitOR($ES_MULTILINE, $ES_READONLY, $WS_VSCROLL))
    GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
    GUICtrlSetBkColor(-1, 0xFFFFF0)
    
    ; Nut dong
    Local $btnClose = GUICtrlCreateButton("Dong", 300, 540, 100, 30)
    GUICtrlSetFont(-1, 10, 600)
    
    GUISetState(@SW_SHOW, $hHelp)
    
    While 1
        Local $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnClose
                GUIDelete($hHelp)
                Return
        EndSwitch
    WEnd
EndFunc
