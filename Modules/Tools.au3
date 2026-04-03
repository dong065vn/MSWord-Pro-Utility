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

    Local $oFind = $g_oDoc.Content.Find
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    
    Local $bCase = (GUICtrlRead($g_chkMatchCase) = $GUI_CHECKED)
    Local $bWord = (GUICtrlRead($g_chkWholeWord) = $GUI_CHECKED)
    
    $oFind.Execute($sFind, $bCase, $bWord, False, False, False, True, 1, False, $sReplace, $WD_REPLACE_ALL)
    _UpdateProgress("Da thay the xong!")
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
        If StringStripWS($sInner, 3) <> "" And Not StringInStr($sInner, "(") And Not StringInStr($sInner, ")") Then
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
    For $i = 0 To UBound($aMatches, 1) - 1
        If _DeleteFirstParenthesizedTextInScope($oScopeRange, $aMatches[$i][2]) Then
            $iRemoved += 1
        EndIf
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
        MsgBox($MB_ICONINFORMATION, "Xem truoc (...)", "Khong tim thay cum nao dang (...) trong " & $sScopeLabel & ".")
        Return
    EndIf

    Local $sPreview = "XEM TRUOC CUM (...) SE BI XOA - " & StringUpper($sScopeLabel) & @CRLF & @CRLF & _
        "Tong so: " & UBound($aMatches, 1) & @CRLF & _
        "Pham vi: " & $sScopeLabel & @CRLF & @CRLF

    Local $iLimit = 40
    For $i = 0 To UBound($aMatches, 1) - 1
        If $i = $iLimit Then
            $sPreview &= "... va " & (UBound($aMatches, 1) - $iLimit) & " cum khac"
            ExitLoop
        EndIf
        $sPreview &= ($i + 1) & ". " & $aMatches[$i][2] & " | " & _ParenthesesPreviewContext($oRange.Text, $aMatches[$i][0], $aMatches[$i][1]) & @CRLF
    Next

    _LogPreview($sPreview)
    MsgBox($MB_ICONINFORMATION, "Xem truoc (...)", $sPreview)
EndFunc

Func _RemoveParenthesizedPhrasesInRange($oRange, $sScopeLabel)
    If Not IsObj($oRange) Then Return

    Local $aMatches = _CollectParenthesizedMatches($oRange.Text)
    If Not IsArray($aMatches) Or UBound($aMatches, 1) = 0 Then
        _UpdateProgress("Khong tim thay cum (...) trong " & $sScopeLabel)
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong tim thay cum nao dang (...) trong " & $sScopeLabel & ".")
        Return
    EndIf

    _UpdateProgress("Dang xoa cum (...) trong " & $sScopeLabel & "...")
    Local $iRemoved = _DeleteParenthesizedMatchesInWordRange($oRange, $aMatches)
    If $iRemoved = 0 Then
        _UpdateProgress("Khong xoa duoc cum (...) trong " & $sScopeLabel)
        MsgBox($MB_ICONWARNING, "Thong bao", "Da tim thay cum (...), nhung Word khong cho phep sua noi dung trong " & $sScopeLabel & ".")
        Return
    EndIf

    _UpdateProgress("Da xoa " & $iRemoved & " cum (...) trong " & $sScopeLabel)
    MsgBox($MB_ICONINFORMATION, "Hoan tat", _
        "Da xoa " & $iRemoved & " cum (...) trong " & $sScopeLabel & "." & @CRLF & @CRLF & _
        "Ho tro 2 che do:" & @CRLF & _
        "- Xoa vung chon" & @CRLF & _
        "- Xoa toan bo tai lieu")
EndFunc

; Resize Images
Func _ResizeImages()
    If Not _CheckConnection() Then Return
    Local $oShapes = $g_oDoc.InlineShapes
    If Not IsObj($oShapes) Or $oShapes.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co hinh!")
        Return
    EndIf

    Local $fMaxW = $g_oDoc.PageSetup.PageWidth - $g_oDoc.PageSetup.LeftMargin - $g_oDoc.PageSetup.RightMargin
    Local $n = 0

    _UpdateProgress("Dang resize hinh...")
    For $i = 1 To $oShapes.Count
        Local $oS = $oShapes.Item($i)
        If Not IsObj($oS) Then ContinueLoop
        
        If $oS.Width > $fMaxW Then
            Local $fRatio = $fMaxW / $oS.Width
            $oS.Width = $fMaxW
            $oS.Height = $oS.Height * $fRatio
            $n += 1
        EndIf
    Next
    _UpdateProgress("Da resize " & $n & "/" & $oShapes.Count & " hinh!")
EndFunc

; Center Images
Func _CenterImages()
    If Not _CheckConnection() Then Return
    Local $oShapes = $g_oDoc.InlineShapes
    If Not IsObj($oShapes) Or $oShapes.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co hinh!")
        Return
    EndIf
    
    _UpdateProgress("Dang can giua hinh...")
    For $i = 1 To $oShapes.Count
        Local $oShape = $oShapes.Item($i)
        If IsObj($oShape) And IsObj($oShape.Range) Then
            $oShape.Range.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
        EndIf
    Next
    _UpdateProgress("Da can giua " & $oShapes.Count & " hinh!")
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
    Local $n = $oShapes.Count
    While $oShapes.Count > 0
        $oShapes.Item(1).Delete()
    WEnd
    _UpdateProgress("Da xoa " & $n & " hinh!")
EndFunc

; AutoFit Tables
Func _AutoFitTables($iMode)
    If Not _CheckConnection() Then Return
    Local $oTables = $g_oDoc.Tables
    If Not IsObj($oTables) Or $oTables.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bang!")
        Return
    EndIf

    _UpdateProgress("Dang AutoFit bang...")
    Local $iSuccess = 0
    For $i = 1 To $oTables.Count
        Local $oTable = $oTables.Item($i)
        If IsObj($oTable) Then
            If $iMode = 1 Then
                $oTable.AutoFitBehavior(1) ; wdAutoFitContent
            Else
                $oTable.AutoFitBehavior(2) ; wdAutoFitWindow
            EndIf
            $iSuccess += 1
        EndIf
    Next
    _UpdateProgress("Da AutoFit " & $iSuccess & "/" & $oTables.Count & " bang!")
EndFunc

; Auto Caption Tables
Func _AutoCaptionTbl()
    _AutoNumberTables()
EndFunc

; Add Table Borders
Func _AddTableBorders()
    If Not _CheckConnection() Then Return
    Local $oTables = $g_oDoc.Tables
    If Not IsObj($oTables) Or $oTables.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bang!")
        Return
    EndIf
    
    _UpdateProgress("Dang them vien bang...")
    For $i = 1 To $oTables.Count
        Local $oTable = $oTables.Item($i)
        If IsObj($oTable) Then
            $oTable.Borders.Enable = True
        EndIf
    Next
    _UpdateProgress("Da them vien " & $oTables.Count & " bang!")
EndFunc


; Word Count
Func _ShowWordCount()
    If Not _CheckConnection() Then Return
    Local $sMsg = "THONG KE TAI LIEU" & @CRLF & @CRLF
    $sMsg &= "So trang: " & $g_oDoc.ComputeStatistics(2) & @CRLF
    $sMsg &= "So tu: " & $g_oDoc.ComputeStatistics(0) & @CRLF
    $sMsg &= "So ky tu: " & $g_oDoc.ComputeStatistics(3) & @CRLF
    $sMsg &= "So doan van: " & $g_oDoc.ComputeStatistics(4) & @CRLF
    $sMsg &= "So bang: " & $g_oDoc.Tables.Count & @CRLF
    $sMsg &= "So hinh: " & $g_oDoc.InlineShapes.Count
    
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
    $sStats &= "So trang: " & $g_oDoc.ComputeStatistics(2) & @CRLF
    $sStats &= "So tu: " & $g_oDoc.ComputeStatistics(0) & @CRLF
    $sStats &= "So ky tu: " & $g_oDoc.ComputeStatistics(3) & @CRLF
    $sStats &= "So bang: " & $g_oDoc.Tables.Count & @CRLF
    $sStats &= "So hinh: " & $g_oDoc.InlineShapes.Count

    FileWrite($sPath, $sStats)
    _UpdateProgress("Da xuat bao cao!")
EndFunc
