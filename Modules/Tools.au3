; ============================================
; TOOLS.AU3 - Module Cong cu
; ============================================

#include-once

; Find & Replace
Func _DoFindReplace()
    If Not _CheckConnection() Then Return
    Local $sFind = GUICtrlRead($g_inputFind)
    Local $sReplace = GUICtrlRead($g_inputReplace)
    If $sFind = "" Then
        MsgBox($MB_ICONWARNING, "Loi", "Nhap tu can tim!")
        Return
    EndIf

    _Process_Start("Find Replace", 1, "toan bo tai lieu")
    Local $oFind = $g_oDoc.Content.Find
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    
    Local $bCase = (GUICtrlRead($g_chkMatchCase) = $GUI_CHECKED)
    Local $bWord = (GUICtrlRead($g_chkWholeWord) = $GUI_CHECKED)
    
    If $g_bDomDryRun Then
        _Process_AddSkipped()
    Else
        Local $vResult = $oFind.Execute($sFind, $bCase, $bWord, False, False, False, True, 1, False, $sReplace, $WD_REPLACE_ALL)
        If $vResult Then _Process_AddChanged()
    EndIf
    _Process_Step("Da thay the xong", 1, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    _Process_Done("Find Replace: changed=" & $g_iProcessChanged & ", skipped=" & $g_iProcessSkipped & ", errors=" & $g_iProcessErrors)
EndFunc

; Find Next
Func _DoFindNext()
    If Not _CheckConnection() Then Return
    Local $sFind = GUICtrlRead($g_inputFind)
    If $sFind = "" Then Return

    Local $oFind = $g_oWord.Selection.Find
    $oFind.ClearFormatting()
    $oFind.Execute($sFind)
EndFunc

Func _ConfigureParenthesesFind($oFind)
    If Not IsObj($oFind) Then Return
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    $oFind.Text = "\([!\)]@\)"
    $oFind.Replacement.Text = ""
    $oFind.MatchWildcards = True
    $oFind.Forward = True
    $oFind.Wrap = 0
    $oFind.Format = False
    $oFind.MatchCase = False
    $oFind.MatchWholeWord = False
EndFunc

Func _CollectParenthesizedMatches($sText)
    Local $aMatches[0][3]
    Local $iLen = StringLen($sText)
    Local $iPos = 1

    While $iPos <= $iLen
        Local $iStart = StringInStr($sText, "(", 0, 1, $iPos)
        If $iStart = 0 Then ExitLoop

        Local $iEnd = StringInStr($sText, ")", 0, 1, $iStart + 1)
        If $iEnd = 0 Then ExitLoop

        Local $sInner = StringMid($sText, $iStart + 1, $iEnd - $iStart - 1)
        If StringStripWS($sInner, 3) <> "" And Not StringInStr($sInner, "(") And Not StringInStr($sInner, ")") And _
                _ShouldRemoveEnglishParenthesizedText($sInner) Then
            Local $iDeleteStart = $iStart
            Local $sPrev = ""
            Local $sNext = ""
            If $iStart > 1 Then $sPrev = StringMid($sText, $iStart - 1, 1)
            If $iEnd < $iLen Then $sNext = StringMid($sText, $iEnd + 1, 1)

            If $sPrev = " " Then
                If $sNext = "" Or $sNext = @CR Or $sNext = @LF Or $sNext = " " Or StringRegExp($sNext, "[\.,;:\!\?\)\]]") Then
                    $iDeleteStart -= 1
                EndIf
            EndIf

            Local $iCount = UBound($aMatches, 1)
            ReDim $aMatches[$iCount + 1][3]
            $aMatches[$iCount][0] = $iDeleteStart
            $aMatches[$iCount][1] = $iEnd
            $aMatches[$iCount][2] = "(" & $sInner & ")"
        EndIf

        $iPos = $iEnd + 1
    WEnd

    Return $aMatches
EndFunc

Func _ShouldRemoveEnglishParenthesizedText($sInner)
    Local $sTrimmed = StringStripWS($sInner, 3)
    If $sTrimmed = "" Then Return False

    ; Giu lai viet tat/to hop viet hoa: KPI, KOL, B2B, ROI...
    If StringRegExp($sTrimmed, "^[A-Z0-9][A-Z0-9\s\/&\-\.\+]{1,20}$") Then Return False

    ; Giu lai nam/trich dan ngan gon: (2011), (2011a), (2020, p. 12)
    If StringRegExp($sTrimmed, "^\d{4}[a-z]?(?:\s*[,;:]\s*(?:p|pp)\.?\s*\d+(?:\s*-\s*\d+)?)?$") Then Return False

    ; Giu lai citation tac gia-nam: (Smith, 2011), (PR Smith, 2011), (Smith et al., 2011)
    If StringRegExp($sTrimmed, "^(?:[A-Z][A-Za-z'`\-]*|et al\.?)(?:\s+(?:[A-Z][A-Za-z'`\-]*|et al\.?))*[,]?\s+\d{4}[a-z]?(?:\s*[,;:]\s*(?:p|pp)\.?\s*\d+(?:\s*-\s*\d+)?)?$") Then Return False

    ; Chi xu ly cum co chu cai Latin va khong co dau tieng Viet.
    If Not StringRegExp($sTrimmed, "[A-Za-z]") Then Return False
    If StringRegExp($sTrimmed, "[^A-Za-z0-9\s,;:\.\-\/&'\+%]") Then Return False

    Local $sLettersOnly = StringRegExpReplace($sTrimmed, "[^A-Za-z]", "")
    If StringLen($sLettersOnly) < 4 Then Return False

    ; Neu toan bo la 1 tu Viet hoa/ngan gon thi uu tien giu lai.
    If StringRegExp($sTrimmed, "^[A-Z][A-Za-z0-9\-]{0,5}$") Then Return False

    Return True
EndFunc

Func _ParenthesesPreviewContext($sText, $iStart1Based, $iEnd1Based)
    Local $iContext = 28
    Local $iStart = $iStart1Based - $iContext
    Local $iLen = ($iEnd1Based - $iStart1Based + 1) + ($iContext * 2)

    If $iStart < 1 Then
        $iLen -= (1 - $iStart)
        $iStart = 1
    EndIf
    If $iLen < 1 Then $iLen = 1

    Local $sSnippet = StringMid($sText, $iStart, $iLen)
    $sSnippet = StringReplace($sSnippet, @CR, " ")
    $sSnippet = StringReplace($sSnippet, @LF, " ")
    $sSnippet = StringRegExpReplace($sSnippet, "\s+", " ")
    $sSnippet = StringStripWS($sSnippet, 3)
    If $iStart > 1 Then $sSnippet = "..." & $sSnippet
    If ($iStart + $iLen - 1) < StringLen($sText) Then $sSnippet &= "..."
    Return $sSnippet
EndFunc

Func _DeleteParenthesizedMatchesInWordRange($oScopeRange, $aMatches)
    If Not IsObj($oScopeRange) Or Not IsArray($aMatches) Then Return 0

    Local $iRemoved = 0
    Local $iBaseStart = $oScopeRange.Start
    For $i = UBound($aMatches, 1) - 1 To 0 Step -1
        Local $iStart = $iBaseStart + $aMatches[$i][0] - 1
        Local $iEnd = $iBaseStart + $aMatches[$i][1]
        Local $oDeleteRange = $g_oDoc.Range($iStart, $iEnd)
        If IsObj($oDeleteRange) And _Perf_RetrySetRangeText($oDeleteRange, "") Then
            $iRemoved += 1
        EndIf
    Next

    If $iRemoved > 0 Then Return $iRemoved

    ; Fallback for unusual Word ranges whose offsets do not map cleanly.
    For $i = 0 To UBound($aMatches, 1) - 1
        If _DeleteFirstParenthesizedTextInScope($oScopeRange, $aMatches[$i][2]) Then $iRemoved += 1
    Next
    Return $iRemoved
EndFunc

Func _DeleteFirstParenthesizedTextInScope($oScopeRange, $sTargetText)
    If Not IsObj($oScopeRange) Or $sTargetText = "" Then Return False

    Local $oSearch = $oScopeRange.Duplicate
    Local $oFind = $oSearch.Find
    If Not IsObj($oFind) Then Return False

    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()

    Local $bFound = $oFind.Execute($sTargetText, False, False, False, False, False, True, 1, False, "", 0)
    If Not $bFound Then Return False

    Local $iStart = $oSearch.Start
    Local $iEnd = $oSearch.End
    Local $sPrev = ""
    Local $sNext = ""

    If $iStart > 0 Then $sPrev = $g_oDoc.Range($iStart - 1, $iStart).Text
    If $iEnd < $g_oDoc.Content.End Then $sNext = $g_oDoc.Range($iEnd, $iEnd + 1).Text

    If $sPrev = " " Then
        If $sNext = "" Or $sNext = @CR Or $sNext = @LF Or $sNext = " " Or StringRegExp($sNext, "[\.,;:\!\?\)\]]") Then
            $iStart -= 1
        EndIf
    EndIf

    Local $oDeleteRange = $g_oDoc.Range($iStart, $iEnd)
    If Not IsObj($oDeleteRange) Then Return False
    $oDeleteRange.Text = ""
    Return True
EndFunc

Func _RemoveParenthesizedPhrasesInSelection()
    If Not _CheckConnection() Then Return
    If Not IsObj($g_oWord) Or Not IsObj($g_oWord.Selection) Then Return

    Local $oSelectionRange = $g_oWord.Selection.Range
    If Not IsObj($oSelectionRange) Then Return

    If $oSelectionRange.Start = $oSelectionRange.End Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Hay boi den vung can xoa truoc.")
        Return
    EndIf

    _RemoveParenthesizedPhrasesInRange($oSelectionRange, "vung chon")
EndFunc

Func _RemoveParenthesizedPhrasesDocument()
    If Not _CheckConnection() Then Return
    _RemoveParenthesizedPhrasesInRange($g_oDoc.Content, "toan bo tai lieu")
EndFunc

Func _PreviewParenthesizedPhrases()
    If Not _CheckConnection() Then Return

    Local $oRange = 0
    Local $sScopeLabel = "toan bo tai lieu"
    Local $oSel = $g_oWord.Selection
    If IsObj($oSel) And $oSel.Type <> 1 Then
        $oRange = $oSel.Range
        $sScopeLabel = "vung chon"
    Else
        $oRange = $g_oDoc.Content
    EndIf

    Local $aMatches = _CollectParenthesizedMatches($oRange.Text)
    If Not IsArray($aMatches) Or UBound($aMatches, 1) = 0 Then
        MsgBox($MB_ICONINFORMATION, "Xem truoc ngoac tieng Anh", "Khong tim thay cum ngoac tieng Anh nao trong " & $sScopeLabel & ".")
        Return
    EndIf

    Local $sPreview = "XEM TRUOC NGOAC TIENG ANH SE BI XOA - " & StringUpper($sScopeLabel) & @CRLF & @CRLF & _
        "Tong so: " & UBound($aMatches, 1) & @CRLF & _
        "Pham vi: " & $sScopeLabel & @CRLF & _
        "Giu lai: nam/trich dan va viet tat (KPI, KOL, ...)" & @CRLF & @CRLF

    Local $iLimit = 40
    For $i = 0 To UBound($aMatches, 1) - 1
        If $i = $iLimit Then
            $sPreview &= "... va " & (UBound($aMatches, 1) - $iLimit) & " cum khac"
            ExitLoop
        EndIf
        $sPreview &= ($i + 1) & ". " & $aMatches[$i][2] & " | " & _ParenthesesPreviewContext($oRange.Text, $aMatches[$i][0], $aMatches[$i][1]) & @CRLF
    Next

    _LogPreview($sPreview)
    MsgBox($MB_ICONINFORMATION, "Xem truoc ngoac tieng Anh", $sPreview)
EndFunc

Func _RemoveParenthesizedPhrasesInRange($oRange, $sScopeLabel)
    If Not IsObj($oRange) Then Return

    Local $aMatches = _CollectParenthesizedMatches($oRange.Text)
    If Not IsArray($aMatches) Or UBound($aMatches, 1) = 0 Then
        _UpdateProgress("Khong tim thay ngoac tieng Anh trong " & $sScopeLabel)
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong tim thay cum ngoac tieng Anh nao trong " & $sScopeLabel & ".")
        Return
    EndIf

    _Process_Start("Xoa ngoac tieng Anh", UBound($aMatches, 1), $sScopeLabel)
    _Process_Step("Dang xoa ngoac tieng Anh", 0)
    Local $aBatch = _Perf_BeginWordBatch("Xoa ngoac tieng Anh")
    Local $iRemoved = _DeleteParenthesizedMatchesInWordRange($oRange, $aMatches)
    _Perf_EndWordBatch($aBatch)
    If $iRemoved = 0 Then
        _UpdateProgress("Khong xoa duoc ngoac tieng Anh trong " & $sScopeLabel)
        MsgBox($MB_ICONWARNING, "Thong bao", "Da tim thay ngoac tieng Anh, nhung Word khong cho phep sua noi dung trong " & $sScopeLabel & ".")
        Return
    EndIf

    _Process_Done("Da xoa " & $iRemoved & " cum ngoac tieng Anh trong " & $sScopeLabel)
    MsgBox($MB_ICONINFORMATION, "Hoan tat", _
        "Da xoa " & $iRemoved & " cum ngoac tieng Anh trong " & $sScopeLabel & "." & @CRLF & @CRLF & _
        "Van giu lai:" & @CRLF & _
        "- Nam/trich dan (vi du: PR Smith (2011), (Smith, 2011))" & @CRLF & _
        "- Viet tat viet hoa (vi du: KPI, KOL)" & @CRLF & @CRLF & _
        "Ho tro 2 che do:" & @CRLF & _
        "- Xoa vung chon" & @CRLF & _
        "- Xoa toan bo tai lieu")
EndFunc

; Resize Images
Func _ResizeImages()
    If Not _CheckConnection() Then Return
    Local $oShapes = $g_oDoc.InlineShapes
    Local $iShapeCount = _Perf_CollectionCount($oShapes)
    If Not IsObj($oShapes) Or $iShapeCount = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co hinh!")
        Return
    EndIf

    Local $fMaxW = $g_oDoc.PageSetup.PageWidth - $g_oDoc.PageSetup.LeftMargin - $g_oDoc.PageSetup.RightMargin
    Local $n = 0

    Local $aBatch = _Perf_BeginWordBatch("Resize Images")
    _Process_Start("Resize Images", $iShapeCount, "inline shapes")
    _Process_Step("Dang resize hinh", 0)
    For $i = 1 To $iShapeCount
        Local $oS = $oShapes.Item($i)
        If Not IsObj($oS) Then
            _Process_AddSkipped()
            ContinueLoop
        EndIf
        If $g_bDomSkipEquations And _Dom_IsMathInlineShape($oS) Then
            _Process_AddSkipped()
            ContinueLoop
        EndIf
        
        If $oS.Width > $fMaxW Then
            Local $fRatio = $fMaxW / $oS.Width
            If Not $g_bDomDryRun Then
                $oS.Width = $fMaxW
                $oS.Height = $oS.Height * $fRatio
            EndIf
            $n += 1
            _Process_AddChanged()
        EndIf
        _Process_Step("Dang resize hinh", $i, $n, $g_iProcessSkipped, $g_iProcessErrors)
    Next
    _Perf_EndWordBatch($aBatch)
    _Process_Done("Da resize " & $n & "/" & $iShapeCount & " hinh")
EndFunc

; Center Images
Func _CenterImages()
    If Not _CheckConnection() Then Return
    Local $oShapes = $g_oDoc.InlineShapes
    Local $iShapeCount = _Perf_CollectionCount($oShapes)
    If Not IsObj($oShapes) Or $iShapeCount = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co hinh!")
        Return
    EndIf
    
    Local $aBatch = _Perf_BeginWordBatch("Center Images")
    _Process_Start("Center Images", $iShapeCount, "inline shapes")
    For $i = 1 To $iShapeCount
        Local $oShape = $oShapes.Item($i)
        If IsObj($oShape) And $g_bDomSkipEquations And _Dom_IsMathInlineShape($oShape) Then
            _Process_AddSkipped()
        ElseIf IsObj($oShape) And IsObj($oShape.Range) Then
            If Not $g_bDomDryRun Then $oShape.Range.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
            _Process_AddChanged()
        Else
            _Process_AddSkipped()
        EndIf
        _Process_Step("Dang can giua hinh", $i, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    Next
    _Perf_EndWordBatch($aBatch)
    _Process_Done("Da can giua " & $g_iProcessChanged & "/" & $iShapeCount & " hinh")
EndFunc

; Auto Caption Images
Func _AutoCaptionImg()
    _AutoNumberImages()
EndFunc

; Remove All Images
Func _RemoveAllImages()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", "Xoa tat ca hinh?") <> $IDYES Then Return
    
    Local $oShapes = $g_oDoc.InlineShapes
    Local $n = _Perf_CollectionCount($oShapes)
    Local $aBatch = _Perf_BeginWordBatch("Remove Images")
    _Process_Start("Remove Images", $n, "inline shapes")
    While $oShapes.Count > 0
        If _Dom_SafeDelete($oShapes.Item(1)) Then _Process_AddChanged()
        _Process_Step("Dang xoa hinh", $g_iProcessChanged + $g_iProcessSkipped, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
        If $g_bDomDryRun Then ExitLoop
    WEnd
    _Perf_EndWordBatch($aBatch)
    _Process_Done("Da xoa " & $g_iProcessChanged & "/" & $n & " hinh")
EndFunc

; AutoFit Tables
Func _AutoFitTables($iMode)
    If Not _CheckConnection() Then Return
    Local $oTables = $g_oDoc.Tables
    Local $iTableCount = _Perf_CollectionCount($oTables)
    If Not IsObj($oTables) Or $iTableCount = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bang!")
        Return
    EndIf

    Local $aBatch = _Perf_BeginWordBatch("AutoFit Tables")
    _Process_Start("AutoFit Tables", $iTableCount, "tables")
    Local $iSuccess = 0
    For $i = 1 To $iTableCount
        Local $oTable = $oTables.Item($i)
        If IsObj($oTable) Then
            If Not $g_bDomDryRun Then
                If $iMode = 1 Then
                    $oTable.AutoFitBehavior(1) ; wdAutoFitContent
                Else
                    $oTable.AutoFitBehavior(2) ; wdAutoFitWindow
                EndIf
            EndIf
            $iSuccess += 1
            _Process_AddChanged()
        Else
            _Process_AddSkipped()
        EndIf
        _Process_Step("Dang AutoFit bang", $i, $iSuccess, $g_iProcessSkipped, $g_iProcessErrors)
    Next
    _Perf_EndWordBatch($aBatch)
    _Process_Done("Da AutoFit " & $iSuccess & "/" & $iTableCount & " bang")
EndFunc

; Auto Caption Tables
Func _AutoCaptionTbl()
    _AutoNumberTables()
EndFunc

; Add Table Borders
Func _AddTableBorders()
    If Not _CheckConnection() Then Return
    Local $oTables = $g_oDoc.Tables
    Local $iTableCount = _Perf_CollectionCount($oTables)
    If Not IsObj($oTables) Or $iTableCount = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bang!")
        Return
    EndIf
    
    Local $aBatch = _Perf_BeginWordBatch("Add Table Borders")
    _Process_Start("Add Table Borders", $iTableCount, "tables")
    For $i = 1 To $iTableCount
        Local $oTable = $oTables.Item($i)
        If IsObj($oTable) Then
            If Not $g_bDomDryRun Then $oTable.Borders.Enable = True
            _Process_AddChanged()
        Else
            _Process_AddSkipped()
        EndIf
        _Process_Step("Dang them vien bang", $i, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    Next
    _Perf_EndWordBatch($aBatch)
    _Process_Done("Da them vien " & $g_iProcessChanged & "/" & $iTableCount & " bang")
EndFunc


; Word Count
Func _ShowWordCount()
    If Not _CheckConnection() Then Return
    Local $sMsg = _BuildDocumentStatsText("THONG KE TAI LIEU")
    
    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "Thong ke", $sMsg)
EndFunc

; Check Spelling
Func _CheckSpelling()
    If Not _CheckConnection() Then Return
    $g_oDoc.CheckSpelling()
    _UpdateProgress("Da kiem tra chinh ta!")
EndFunc

; Check Format
Func _CheckFormat()
    _CheckThesisFormat()
EndFunc

; Show Detailed Stats
Func _ShowDetailedStats()
    _ShowWordCount()
EndFunc

; Export Stats
Func _ExportStats()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Luu bao cao", @DesktopDir, "Text (*.txt)", 16, "ThongKe.txt")
    If @error Then Return

    Local $sStats = "THONG KE TAI LIEU: " & $g_oDoc.Name & @CRLF
    $sStats &= "Ngay: " & @MDAY & "/" & @MON & "/" & @YEAR & @CRLF & @CRLF
    $sStats &= _BuildDocumentStatsText("")

    FileWrite($sPath, $sStats)
    _UpdateProgress("Da xuat bao cao!")
EndFunc

Func _BuildDocumentStatsText($sTitle = "")
    Local $iPages = $g_oDoc.ComputeStatistics(2)
    Local $iWords = $g_oDoc.ComputeStatistics(0)
    Local $iChars = $g_oDoc.ComputeStatistics(3)
    Local $iParas = $g_oDoc.ComputeStatistics(4)
    Local $iTables = _Perf_CollectionCount($g_oDoc.Tables)
    Local $iImages = _Perf_CollectionCount($g_oDoc.InlineShapes)

    Local $sMsg = ""
    If $sTitle <> "" Then $sMsg &= $sTitle & @CRLF & @CRLF
    $sMsg &= "So trang: " & $iPages & @CRLF
    $sMsg &= "So tu: " & $iWords & @CRLF
    $sMsg &= "So ky tu: " & $iChars & @CRLF
    If $sTitle <> "" Then $sMsg &= "So doan van: " & $iParas & @CRLF
    $sMsg &= "So bang: " & $iTables & @CRLF
    $sMsg &= "So hinh: " & $iImages
    Return $sMsg
EndFunc
