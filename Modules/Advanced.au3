; ============================================
; ADVANCED.AU3 - Module Nang cao
; ============================================

#include-once

; Auto Detect Heading
Func _AutoDetectHeading()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang tu dong gan Heading...")
    
    Local $iH1 = 0, $iH2 = 0, $iH3 = 0
    Local $oParas = $g_oDoc.Paragraphs
    
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        Local $oRange = $oPara.Range
        Local $sText = StringStripWS($oRange.Text, 3)
        
        If StringLen($sText) < 3 Or StringLen($sText) > 100 Then ContinueLoop
        
        ; Get font properties - handle mixed formatting (9999999 = wdUndefined)
        Local $bBold = $oRange.Font.Bold
        Local $bItalic = $oRange.Font.Italic
        Local $fSize = $oRange.Font.Size
        
        ; Skip if mixed formatting (undefined values)
        If $bBold = 9999999 Or $fSize = 9999999 Then ContinueLoop
        
        ; Detect H1: Bold, Size >= 14
        If $bBold = True And $fSize >= 14 Then
            $oRange.Style = "Heading 1"
            $iH1 += 1
            ContinueLoop
        EndIf
        
        ; Detect H2: Bold, Size >= 13
        If $bBold = True And $fSize >= 13 Then
            $oRange.Style = "Heading 2"
            $iH2 += 1
            ContinueLoop
        EndIf
        
        ; Detect H3: Bold + Italic
        If $bBold = True And $bItalic = True Then
            $oRange.Style = "Heading 3"
            $iH3 += 1
        EndIf
    Next
    
    _UpdateProgress("Da gan: H1=" & $iH1 & ", H2=" & $iH2 & ", H3=" & $iH3)
    MsgBox($MB_ICONINFORMATION, "Ket qua", _
        "Da tu dong gan Heading:" & @CRLF & _
        "Heading 1: " & $iH1 & @CRLF & _
        "Heading 2: " & $iH2 & @CRLF & _
        "Heading 3: " & $iH3)
EndFunc

; Reset All Headings
Func _ResetAllHeadings()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", "Reset tat ca Heading ve Normal?") <> $IDYES Then Return
    
    _UpdateProgress("Dang reset Headings...")
    Local $oParas = $g_oDoc.Paragraphs
    Local $iReset = 0
    
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        
        Local $oStyle = $oPara.Style
        If Not IsObj($oStyle) Then ContinueLoop
        
        Local $sStyle = $oStyle.NameLocal
        If StringInStr($sStyle, "Heading") Or StringInStr($sStyle, "Tieu de") Then
            $oPara.Style = "Normal"
            $iReset += 1
        EndIf
    Next
    
    _UpdateProgress("Da reset " & $iReset & " Heading!")
EndFunc

; Heading to TOC
Func _HeadingToTOC()
    _CreateTOC()
EndFunc

; List All Headings
Func _ListAllHeadings()
    If Not _CheckConnection() Then Return
    
    Local $sMsg = "DANH SACH HEADING:" & @CRLF & @CRLF
    Local $oParas = $g_oDoc.Paragraphs
    Local $iCount = 0
    
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        
        Local $oStyle = $oPara.Style
        If Not IsObj($oStyle) Then ContinueLoop
        
        Local $sStyle = $oStyle.NameLocal
        If StringInStr($sStyle, "Heading") Or StringInStr($sStyle, "Tieu de") Then
            Local $sText = StringStripWS($oPara.Range.Text, 3)
            $sText = StringLeft($sText, 50)
            $sMsg &= $sStyle & ": " & $sText & @CRLF
            $iCount += 1
            If $iCount >= 50 Then
                $sMsg &= "... va con nua"
                ExitLoop
            EndIf
        EndIf
    Next
    
    If $iCount = 0 Then $sMsg &= "(Khong co Heading)"
    
    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "Danh sach Heading (" & $iCount & ")", $sMsg)
EndFunc

; Remove All Formatting
Func _RemoveAllFormatting()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", "Xoa tat ca dinh dang?") <> $IDYES Then Return
    $g_oDoc.Content.ClearFormatting()
    _UpdateProgress("Da xoa tat ca dinh dang!")
EndFunc

; Convert Text Case
Func _ConvertTextCase()
    If Not _CheckConnection() Then Return
    Local $sChoice = InputBox("Chuyen doi chu", _
        "1 - UPPER CASE" & @CRLF & _
        "2 - lower case" & @CRLF & _
        "3 - Title Case", "1")
    If @error Or $sChoice = "" Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon van ban truoc!")
        Return
    EndIf
    
    Switch $sChoice
        Case "1"
            $oSel.Range.Case = 1 ; wdUpperCase
        Case "2"
            $oSel.Range.Case = 2 ; wdLowerCase
        Case "3"
            $oSel.Range.Case = 4 ; wdTitleCase
        Case Else
            MsgBox($MB_ICONWARNING, "Loi", "Chon 1, 2 hoac 3!")
            Return
    EndSwitch
    _UpdateProgress("Da chuyen doi!")
EndFunc

; Remove All Hyperlinks
Func _RemoveAllHyperlinks()
    If Not _CheckConnection() Then Return
    Local $oLinks = $g_oDoc.Hyperlinks
    Local $n = $oLinks.Count
    While $oLinks.Count > 0
        $oLinks.Item(1).Delete()
    WEnd
    _UpdateProgress("Da xoa " & $n & " hyperlinks!")
EndFunc


; Remove All Comments
Func _RemoveAllComments()
    If Not _CheckConnection() Then Return
    Local $oComments = $g_oDoc.Comments
    Local $n = $oComments.Count
    While $oComments.Count > 0
        $oComments.Item(1).Delete()
    WEnd
    _UpdateProgress("Da xoa " & $n & " comments!")
EndFunc

; Accept All Changes
Func _AcceptAllChanges()
    If Not _CheckConnection() Then Return
    $g_oDoc.AcceptAllRevisions()
    _UpdateProgress("Da chap nhan tat ca thay doi!")
EndFunc

; Convert Numbering to Text
; FIX: Su dung while-loop de xu ly Live Collection, xoa list formatting sau khi convert
Func _ConvertNumberingToText()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", _
        "Chuyen Numbering/Bullet thanh text?" & @CRLF & @CRLF & _
        "Chuc nang nay se:" & @CRLF & _
        "1. Chuyen so thu tu (1. 2. 3.) thanh text" & @CRLF & _
        "2. Chuyen bullet thanh ky tu text" & @CRLF & _
        "3. Xoa dinh dang list (thut le, numbering)" & @CRLF & @CRLF & _
        "LUU Y: Nen Backup truoc!") <> $IDYES Then Return

    _UpdateProgress("Dang chuyen Numbering thanh text...")

    ; Buoc 1: Dem so luong lists ban dau
    Local $iTotal = $g_oDoc.Lists.Count
    If $iTotal = 0 Then
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong tim thay Numbering/Bullet nao!")
        _UpdateProgress("")
        Return
    EndIf

    ; Buoc 2: Convert numbers thanh text
    ; QUAN TRONG: Lists la Live Collection - sau moi ConvertNumbersToText(),
    ; list bi xoa khoi collection. Dung While loop de xu ly an toan.
    Local $iConverted = 0
    Local $iFailed = 0

    While $g_oDoc.Lists.Count > 0
        _UpdateProgress("Dang chuyen list " & ($iConverted + 1) & "/" & $iTotal & "...")
        Local $oList = $g_oDoc.Lists.Item(1)
        If IsObj($oList) Then
            ; Luu range truoc khi convert de xoa formatting sau
            Local $oRange = $oList.Range
            $oList.ConvertNumbersToText()
            If @error Then
                $iFailed += 1
            Else
                ; Buoc 3: Xoa list formatting con sot lai (thut le, ListTemplate)
                ; Neu khong xoa, paragraph van giu indent va co the bi re-number
                If IsObj($oRange) Then
                    $oRange.ListFormat.RemoveNumbers()
                EndIf
                $iConverted += 1
            EndIf
        Else
            ExitLoop ; Tranh vong lap vo han
        EndIf
    WEnd

    ; Buoc 4: Bao cao ket qua
    Local $sMsg = "CHUYEN NUMBERING THANH TEXT" & @CRLF & @CRLF
    $sMsg &= "Tong cong: " & $iTotal & " lists" & @CRLF
    $sMsg &= "Thanh cong: " & $iConverted & @CRLF
    If $iFailed > 0 Then $sMsg &= "Loi: " & $iFailed & @CRLF

    _LogPreview($sMsg)
    _UpdateProgress("Da chuyen " & $iConverted & "/" & $iTotal & " lists!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", $sMsg)
EndFunc

; Convert Bullet to Text
Func _ConvertBulletToText()
    _ConvertNumberingToText()
EndFunc

; Convert Numbering to Text (Selection)
; FIX: Them IsObj check, xu ly insertion point, xoa list formatting
Func _ConvertNumberingToTextSelection()
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then
        MsgBox($MB_ICONWARNING, "Loi", "Khong lay duoc Selection!")
        Return
    EndIf

    ; Kiem tra selection co text khong
    If $oSel.Type = 1 Then ; wdSelectionIP = insertion point
        MsgBox($MB_ICONWARNING, "Loi", "Vui long boi den (select) vung van ban co numbering!")
        Return
    EndIf

    Local $oRange = $oSel.Range
    If Not IsObj($oRange) Then Return

    If $oRange.ListFormat.ListType <> 0 Then ; 0 = wdListNoNumbering
        _UpdateProgress("Dang chuyen vung chon...")
        $oRange.ListFormat.ConvertNumbersToText()
        ; Xoa list formatting con sot lai
        $oRange.ListFormat.RemoveNumbers()
        _UpdateProgress("Da chuyen vung chon thanh text!")
    Else
        MsgBox($MB_ICONWARNING, "Thong bao", _
            "Vung chon khong co Numbering/Bullet!" & @CRLF & @CRLF & _
            "Huong dan:" & @CRLF & _
            "1. Boi den cac dong co so thu tu hoac bullet" & @CRLF & _
            "2. Nhan lai nut nay")
    EndIf
EndFunc

; Export to PDF
Func _ExportToPDF()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Xuat PDF", @DesktopDir, "PDF (*.pdf)", 16, _
        StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "") & ".pdf")
    If @error Then Return
    
    _UpdateProgress("Dang xuat PDF...")
    _ExportCurrentDocumentToPath($sPath, $WD_EXPORT_PDF)
    _UpdateProgress("Da xuat PDF!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da xuat: " & $sPath)
EndFunc

Func _ScanAndApplyThesisHeadingStyles()
    If Not _CheckConnection() Then Return

    Local $aMatches = _ScanThesisHeadingParagraphs()
    If @error Or Not IsArray($aMatches) Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong quet duoc danh sach de muc.")
        Return
    EndIf

    Local $aCounts[5] = [0, 0, 0, 0, 0]
    For $i = 1 To UBound($aMatches) - 1
        Local $iLevel = $aMatches[$i][0]
        If $iLevel >= 1 And $iLevel <= 4 Then $aCounts[$iLevel] += 1
    Next

    If $aCounts[1] + $aCounts[2] + $aCounts[3] + $aCounts[4] = 0 Then
        MsgBox($MB_ICONINFORMATION, "Khong tim thay", _
            "Khong tim thay chuong/de muc theo mau do an." & @CRLF & @CRLF & _
            "Mau dang duoc quet:" & @CRLF & _
            "- CHUONG 1 / Chuong 1" & @CRLF & _
            "- 1 Tieu de / 1. Tieu de" & @CRLF & _
            "- 1.1 Tieu de" & @CRLF & _
            "- 1.1.1 Tieu de" & @CRLF & _
            "- 1.1.1.1 Tieu de")
        Return
    EndIf

    _ShowThesisHeadingStyleDialog($aMatches, $aCounts)
EndFunc

Func _ScanThesisHeadingParagraphs()
    Local $oParas = $g_oDoc.Paragraphs
    If Not IsObj($oParas) Then Return SetError(1, 0, 0)

    Local $aFound[1][4]
    Local $iFound = 0

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sText = StringReplace($oPara.Range.Text, @CR, "")
        $sText = StringReplace($sText, @LF, "")
        $sText = StringStripWS($sText, 3)
        If $sText = "" Then ContinueLoop

        Local $iLevel = _DetectThesisHeadingLevel($sText)
        If $iLevel = 0 Then ContinueLoop

        $iFound += 1
        ReDim $aFound[$iFound + 1][4]
        $aFound[$iFound][0] = $iLevel
        $aFound[$iFound][1] = $i
        $aFound[$iFound][2] = $sText
        $aFound[$iFound][3] = ""
    Next

    Return $aFound
EndFunc

Func _DetectThesisHeadingLevel($sText)
    If StringRegExp($sText, "^(CHUONG|Chuong|CHƯƠNG|Chương)\s+[IVXLC0-9]+\b") Then Return 1
    If StringRegExp($sText, "^\d+\.?\s+\S+") Then Return 1
    If StringRegExp($sText, "^\d+\.\d+\.?\s+\S+") Then Return 2
    If StringRegExp($sText, "^\d+\.\d+\.\d+\.?\s+\S+") Then Return 3
    If StringRegExp($sText, "^\d+\.\d+\.\d+\.\d+\.?\s+\S+") Then Return 4
    Return 0
EndFunc

Func _GetParagraphStylesData()
    Return _GetParagraphStylesDataFromDoc($g_oDoc)
EndFunc

Func _GetParagraphStylesDataFromDoc($oDocSource)
    If Not IsObj($oDocSource) Then Return ""

    Local $oStyles = $oDocSource.Styles
    If Not IsObj($oStyles) Then Return ""

    Local $sData = ""
    For $i = 1 To $oStyles.Count
        Local $oStyle = $oStyles.Item($i)
        If Not IsObj($oStyle) Then ContinueLoop
        If $oStyle.Type <> $wdStyleTypeParagraph Then ContinueLoop

        Local $sName = $oStyle.NameLocal
        If $sName = "" Then ContinueLoop
        If StringLeft($sName, 1) = "_" Then ContinueLoop
        If StringInStr($sData, "|" & $sName & "|") > 0 Then ContinueLoop
        $sData &= $sName & "|"
    Next

    If $sData = "" Then Return ""
    Return StringTrimRight($sData, 1)
EndFunc

Func _GetOpenWordDocsListData($bIncludeCurrent = True)
    If Not IsObj($g_oWord) Then Return ""
    If $g_oWord.Documents.Count = 0 Then Return ""

    Local $sData = ""
    If $bIncludeCurrent Then $sData = "Chinh van ban dang sua|"

    For $i = 1 To $g_oWord.Documents.Count
        Local $oDocItem = $g_oWord.Documents.Item($i)
        If Not IsObj($oDocItem) Then ContinueLoop
        $sData &= $i & ". " & $oDocItem.Name & "|"
    Next

    If StringRight($sData, 1) = "|" Then $sData = StringTrimRight($sData, 1)
    Return $sData
EndFunc

Func _GetStyleSourceDocFromSelection($sSelection)
    If $sSelection = "" Or $sSelection = "Chinh van ban dang sua" Then Return $g_oDoc
    If Not IsObj($g_oWord) Then Return 0

    Local $iDotPos = StringInStr($sSelection, ".")
    If $iDotPos = 0 Then Return 0

    Local $iIndex = Int(StringLeft($sSelection, $iDotPos - 1))
    If $iIndex < 1 Or $iIndex > $g_oWord.Documents.Count Then Return 0
    Return $g_oWord.Documents.Item($iIndex)
EndFunc

Func _ResolvePreferredStyleName($sStyles, $sPreferred, $sFallback = "")
    Local $sWrapped = "|" & $sStyles & "|"
    If $sPreferred <> "" And StringInStr($sWrapped, "|" & $sPreferred & "|") > 0 Then Return $sPreferred
    If $sFallback <> "" And StringInStr($sWrapped, "|" & $sFallback & "|") > 0 Then Return $sFallback

    Local $aStyles = StringSplit($sStyles, "|", 2)
    If IsArray($aStyles) And UBound($aStyles) > 0 Then Return $aStyles[0]
    Return ""
EndFunc

Func _ShowThesisHeadingStyleDialog(ByRef $aMatches, ByRef $aCounts)
    Local $sDocList = _GetOpenWordDocsListData(True)
    If $sDocList = "" Then
        MsgBox($MB_ICONWARNING, "Khong co file", "Khong lay duoc danh sach file Word dang mo.")
        Return
    EndIf

    Local $oStyleSourceDoc = $g_oDoc
    Local $sStyles = _GetParagraphStylesDataFromDoc($oStyleSourceDoc)
    If $sStyles = "" Then
        MsgBox($MB_ICONWARNING, "Khong co style", "Khong doc duoc danh sach style doan van trong van ban hien tai.")
        Return
    EndIf

    Local $hPopup = GUICreate("Quet de muc do an va chon style", 720, 470, -1, -1, _
        BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xF7F9FB, $hPopup)

    GUICtrlCreateLabel("Da quet thay cac de muc sau. Chon style co san cho tung cap truoc khi ap dung:", 20, 15, 670, 20)
    GUICtrlSetFont(-1, 9, 600)

    GUICtrlCreateLabel("Nguon style:", 30, 48, 90, 20)
    Local $cboStyleSource = GUICtrlCreateCombo("", 120, 43, 360, 24, $CBS_DROPDOWNLIST)
    GUICtrlSetData($cboStyleSource, $sDocList, "Chinh van ban dang sua")
    Local $btnRefreshSource = GUICtrlCreateButton("Lam moi", 490, 42, 80, 26)
    Local $lblSourceHint = GUICtrlCreateLabel("Mac dinh lay style tu chinh file dang sua. Co the doi sang file nguon dang mo.", 30, 73, 620, 18)
    GUICtrlSetColor($lblSourceHint, 0x555555)

    GUICtrlCreateLabel("Cap chuong / muc 1: " & $aCounts[1] & " dong", 30, 108, 220, 20)
    Local $cboL1 = GUICtrlCreateCombo("", 260, 103, 220, 24, $CBS_DROPDOWNLIST)

    GUICtrlCreateLabel("Cap 1.1: " & $aCounts[2] & " dong", 30, 148, 220, 20)
    Local $cboL2 = GUICtrlCreateCombo("", 260, 143, 220, 24, $CBS_DROPDOWNLIST)

    GUICtrlCreateLabel("Cap 1.1.1: " & $aCounts[3] & " dong", 30, 188, 220, 20)
    Local $cboL3 = GUICtrlCreateCombo("", 260, 183, 220, 24, $CBS_DROPDOWNLIST)

    GUICtrlCreateLabel("Cap 1.1.1.1: " & $aCounts[4] & " dong", 30, 228, 220, 20)
    Local $cboL4 = GUICtrlCreateCombo("", 260, 223, 220, 24, $CBS_DROPDOWNLIST)

    _RefreshThesisStyleCombos($cboL1, $cboL2, $cboL3, $cboL4, $sStyles)

    GUICtrlCreateLabel("Xem nhanh 12 dong dau tien duoc nhan dien:", 20, 270, 260, 20)
    Local $sPreview = ""
    Local $iLimit = UBound($aMatches) - 1
    If $iLimit > 12 Then $iLimit = 12
    For $i = 1 To $iLimit
        $sPreview &= "[" & $aMatches[$i][0] & "] " & $aMatches[$i][2] & @CRLF
    Next
    GUICtrlCreateEdit($sPreview, 20, 295, 675, 105, BitOR($ES_READONLY, $WS_VSCROLL))

    Local $btnApply = GUICtrlCreateButton("Ap dung hang loat", 405, 415, 140, 34, $BS_DEFPUSHBUTTON)
    GUICtrlSetBkColor($btnApply, 0x27AE60)
    Local $btnCancel = GUICtrlCreateButton("Dong", 555, 415, 140, 34)

    GUISetState(@SW_SHOW, $hPopup)

    While 1
        Local $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                GUIDelete($hPopup)
                Return
            Case $cboStyleSource, $btnRefreshSource
                If $iMsg = $btnRefreshSource Then
                    Local $sCurrentSelection = GUICtrlRead($cboStyleSource)
                    $sDocList = _GetOpenWordDocsListData(True)
                    GUICtrlSetData($cboStyleSource, "", "")
                    GUICtrlSetData($cboStyleSource, $sDocList, $sCurrentSelection)
                    If GUICtrlRead($cboStyleSource) = "" Then GUICtrlSetData($cboStyleSource, $sDocList, "Chinh van ban dang sua")
                EndIf

                $oStyleSourceDoc = _GetStyleSourceDocFromSelection(GUICtrlRead($cboStyleSource))
                If Not IsObj($oStyleSourceDoc) Then
                    MsgBox($MB_ICONWARNING, "Loi", "Khong mo duoc file nguon style.")
                    ContinueLoop
                EndIf

                $sStyles = _GetParagraphStylesDataFromDoc($oStyleSourceDoc)
                If $sStyles = "" Then
                    MsgBox($MB_ICONWARNING, "Khong co style", "File nguon khong doc duoc danh sach style doan van.")
                    ContinueLoop
                EndIf

                _RefreshThesisStyleCombos($cboL1, $cboL2, $cboL3, $cboL4, $sStyles)
            Case $btnApply
                Local $aStyleMap[5]
                $aStyleMap[1] = GUICtrlRead($cboL1)
                $aStyleMap[2] = GUICtrlRead($cboL2)
                $aStyleMap[3] = GUICtrlRead($cboL3)
                $aStyleMap[4] = GUICtrlRead($cboL4)

                If $aStyleMap[1] = "" Or $aStyleMap[2] = "" Or $aStyleMap[3] = "" Or $aStyleMap[4] = "" Then
                    MsgBox($MB_ICONWARNING, "Chua chon du", "Vui long chon style cho tat ca cac cap.")
                    ContinueLoop
                EndIf

                GUIDelete($hPopup)
                _ApplyDetectedThesisHeadingStyles($aMatches, $aStyleMap)
                Return
        EndSwitch
    WEnd
EndFunc

Func _RefreshThesisStyleCombos($cboL1, $cboL2, $cboL3, $cboL4, $sStyles)
    Local $sDefaultL1 = _ResolvePreferredStyleName($sStyles, "Heading 1")
    Local $sDefaultL2 = _ResolvePreferredStyleName($sStyles, "Heading 2", $sDefaultL1)
    Local $sDefaultL3 = _ResolvePreferredStyleName($sStyles, "Heading 3", $sDefaultL2)
    Local $sDefaultL4 = _ResolvePreferredStyleName($sStyles, "Heading 4", $sDefaultL3)

    GUICtrlSetData($cboL1, "", "")
    GUICtrlSetData($cboL2, "", "")
    GUICtrlSetData($cboL3, "", "")
    GUICtrlSetData($cboL4, "", "")

    GUICtrlSetData($cboL1, $sStyles, $sDefaultL1)
    GUICtrlSetData($cboL2, $sStyles, $sDefaultL2)
    GUICtrlSetData($cboL3, $sStyles, $sDefaultL3)
    GUICtrlSetData($cboL4, $sStyles, $sDefaultL4)
EndFunc

Func _ApplyDetectedThesisHeadingStyles(ByRef $aMatches, ByRef $aStyleMap)
    _UpdateProgress("Dang ap dung style cho de muc do an...")

    Local $oParas = $g_oDoc.Paragraphs
    Local $iApplied = 0
    Local $iFailed = 0

    For $i = 1 To UBound($aMatches) - 1
        Local $iLevel = $aMatches[$i][0]
        Local $iParaIndex = $aMatches[$i][1]
        If $iLevel < 1 Or $iLevel > 4 Then ContinueLoop

        Local $oPara = $oParas.Item($iParaIndex)
        If Not IsObj($oPara) Then
            $iFailed += 1
            ContinueLoop
        EndIf

        $oPara.Range.Style = $aStyleMap[$iLevel]
        If @error Then
            $iFailed += 1
        Else
            $iApplied += 1
        EndIf
    Next

    _UpdateProgress("Da ap dung style cho " & $iApplied & " de muc")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", _
        "Da ap dung style cho " & $iApplied & " de muc." & @CRLF & _
        "Khong ap dung duoc: " & $iFailed)
EndFunc

; Export to HTML
Func _ExportToHTML()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Xuat HTML", @DesktopDir, "HTML (*.html)", 16, _
        StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "") & ".html")
    If @error Then Return
    
    _UpdateProgress("Dang xuat HTML...")
    _ExportCurrentDocumentToPath($sPath, $WD_FORMAT_HTML)
    _UpdateProgress("Da xuat HTML!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da xuat: " & $sPath)
EndFunc

; Export to TXT
Func _ExportToTXT()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Xuat TXT", @DesktopDir, "Text (*.txt)", 16, _
        StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "") & ".txt")
    If @error Then Return
    
    _UpdateProgress("Dang xuat TXT...")
    _ExportCurrentDocumentToPath($sPath, $WD_FORMAT_TEXT)
    _UpdateProgress("Da xuat TXT!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da xuat: " & $sPath)
EndFunc

; Export to RTF
Func _ExportToRTF()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Xuat RTF", @DesktopDir, "RTF (*.rtf)", 16, _
        StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "") & ".rtf")
    If @error Then Return
    
    _UpdateProgress("Dang xuat RTF...")
    _ExportCurrentDocumentToPath($sPath, $WD_FORMAT_RTF)
    _UpdateProgress("Da xuat RTF!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da xuat: " & $sPath)
EndFunc

Func _ExportCurrentDocumentToPath($sPath, $iFormat)
    If Not _CheckConnection() Then Return False
    If $sPath = "" Then Return False

    Local $sSourcePath = ""
    Local $bDeleteSource = False

    If $iFormat = $WD_EXPORT_PDF Then
        $g_oDoc.ExportAsFixedFormat($sPath, $iFormat)
        Return (Not @error And FileExists($sPath))
    EndIf

    If Not _PrepareDocumentPathForExport($g_oDoc, $sSourcePath, $bDeleteSource) Then Return False

    Local $oExportDoc = $g_oWord.Documents.Open($sSourcePath, False, False)
    If Not IsObj($oExportDoc) Then
        If $bDeleteSource And FileExists($sSourcePath) Then FileDelete($sSourcePath)
        Return False
    EndIf

    $oExportDoc.SaveAs2($sPath, $iFormat)
    Local $bOk = (Not @error And FileExists($sPath))
    $oExportDoc.Close(0)

    If IsObj($g_oDoc) Then $g_oDoc.Activate()

    If $bDeleteSource And FileExists($sSourcePath) Then FileDelete($sSourcePath)
    Return $bOk
EndFunc

Func _PrepareDocumentPathForExport($oDoc, ByRef $sSourcePath, ByRef $bDeleteSource)
    If Not IsObj($oDoc) Then Return False

    $sSourcePath = ""
    $bDeleteSource = False

    Local $sCurrentPath = $oDoc.FullName
    If $sCurrentPath <> "" And FileExists($sCurrentPath) Then
        $oDoc.Save()
        $sSourcePath = $sCurrentPath
        Return FileExists($sSourcePath)
    EndIf

    Local $sTempDir = @TempDir & "\PDFToWordFixer\Exports"
    DirCreate($sTempDir)
    $sSourcePath = $sTempDir & "\ExportSource_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & "_" & @MSEC & ".docx"

    Local $oTempDoc = $g_oWord.Documents.Add()
    If Not IsObj($oTempDoc) Then Return False

    $oDoc.Content.Copy()
    $oTempDoc.Range(0, 0).Paste()
    $oTempDoc.PageSetup.LeftMargin = $oDoc.PageSetup.LeftMargin
    $oTempDoc.PageSetup.RightMargin = $oDoc.PageSetup.RightMargin
    $oTempDoc.PageSetup.TopMargin = $oDoc.PageSetup.TopMargin
    $oTempDoc.PageSetup.BottomMargin = $oDoc.PageSetup.BottomMargin
    $oTempDoc.SaveAs2($sSourcePath, 16)
    $oTempDoc.Close(0)

    If @error Or Not FileExists($sSourcePath) Then Return False

    $bDeleteSource = True
    Return True
EndFunc

; Show Print Preview
Func _ShowPrintPreview()
    If Not _CheckConnection() Then Return
    $g_oDoc.PrintPreview()
EndFunc

; Compare Documents
; FIX: Them error handling cho CompareDocuments COM call
Func _CompareDocuments()
    If Not _CheckConnection() Then Return
    _ShowCompareDocumentsDialog()
EndFunc

Func _ShowCompareDocumentsDialog()
    Local $sDocList = _GetOpenWordDocsListData(True)
    If $sDocList = "" Then
        MsgBox($MB_ICONWARNING, "Khong co file", "Khong lay duoc danh sach file Word dang mo.")
        Return
    EndIf

    Local $sDefaultRevised = _GetDefaultCompareDocSelection()
    Local $hDlg = GUICreate("So sanh 2 ban Word", 930, 620, -1, -1, BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xF7F9FB, $hDlg)

    GUICtrlCreateLabel("Chon 2 nguon de so sanh. Co the lay tu van ban dang mo hoac chon file ben ngoai.", 20, 15, 820, 20)
    GUICtrlSetFont(-1, 9, 600)

    GUICtrlCreateGroup(" Ban goc ", 20, 45, 890, 95)
    GUICtrlCreateLabel("Nguon dang mo:", 35, 72, 90, 20)
    Local $cboOriginal = GUICtrlCreateCombo("", 130, 67, 360, 24, $CBS_DROPDOWNLIST)
    GUICtrlSetData($cboOriginal, $sDocList, "Chinh van ban dang sua")
    GUICtrlCreateLabel("Hoac file:", 510, 72, 55, 20)
    Local $inpOriginalPath = GUICtrlCreateInput("", 570, 67, 250, 24)
    Local $btnBrowseOriginal = GUICtrlCreateButton("Chon...", 830, 66, 60, 26)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    GUICtrlCreateGroup(" Ban sua ", 20, 145, 890, 95)
    GUICtrlCreateLabel("Nguon dang mo:", 35, 172, 90, 20)
    Local $cboRevised = GUICtrlCreateCombo("", 130, 167, 360, 24, $CBS_DROPDOWNLIST)
    GUICtrlSetData($cboRevised, $sDocList, $sDefaultRevised)
    GUICtrlCreateLabel("Hoac file:", 510, 172, 55, 20)
    Local $inpRevisedPath = GUICtrlCreateInput("", 570, 167, 250, 24)
    Local $btnBrowseRevised = GUICtrlCreateButton("Chon...", 830, 166, 60, 26)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    Local $btnCompare = GUICtrlCreateButton("Tao ban so sanh", 690, 250, 200, 34, $BS_DEFPUSHBUTTON)
    GUICtrlSetBkColor($btnCompare, 0x27AE60)
    Local $lblSummary = GUICtrlCreateLabel("Chua tao ban so sanh.", 20, 255, 640, 22)
    GUICtrlSetColor($lblSummary, 0x2C3E50)

    Local $listRevisions = GUICtrlCreateListView("STT|Loai thay doi|Noi dung|Tac gia|Thoi gian", 20, 295, 890, 230, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS))
    _GUICtrlListView_SetExtendedListViewStyle($listRevisions, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
    _GUICtrlListView_SetColumnWidth($listRevisions, 0, 45)
    _GUICtrlListView_SetColumnWidth($listRevisions, 1, 140)
    _GUICtrlListView_SetColumnWidth($listRevisions, 2, 420)
    _GUICtrlListView_SetColumnWidth($listRevisions, 3, 120)
    _GUICtrlListView_SetColumnWidth($listRevisions, 4, 140)

    Local $btnRefresh = GUICtrlCreateButton("Tai lai danh sach", 20, 535, 120, 30)
    Local $btnGoTo = GUICtrlCreateButton("Xem thay doi", 150, 535, 120, 30)
    Local $btnOpenCompare = GUICtrlCreateButton("Mo file so sanh", 280, 535, 120, 30)
    Local $btnShowMarkup = GUICtrlCreateButton("Hien markup", 410, 535, 100, 30)
    Local $btnHideMarkup = GUICtrlCreateButton("An markup", 520, 535, 100, 30)
    Local $btnClose = GUICtrlCreateButton("Dong", 810, 535, 100, 30)

    Local $idDetail = GUICtrlCreateEdit("", 20, 570, 890, 35, BitOR($ES_READONLY, $WS_VSCROLL))

    Local $oCompareDoc = 0
    Local $aRevisionMap[1] = [0]

    GUISetState(@SW_SHOW, $hDlg)

    While 1
        Local $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnClose
                GUIDelete($hDlg)
                Return

            Case $btnBrowseOriginal
                Local $sPath1 = FileOpenDialog("Chon ban goc", @ScriptDir, "Word (*.docx;*.doc)", 1)
                If Not @error Then GUICtrlSetData($inpOriginalPath, $sPath1)

            Case $btnBrowseRevised
                Local $sPath2 = FileOpenDialog("Chon ban sua", @ScriptDir, "Word (*.docx;*.doc)", 1)
                If Not @error Then GUICtrlSetData($inpRevisedPath, $sPath2)

            Case $btnCompare
                Local $oDocOriginal = 0, $oDocRevised = 0
                Local $bCloseOriginal = False, $bCloseRevised = False

                If Not _ResolveCompareSource(GUICtrlRead($cboOriginal), GUICtrlRead($inpOriginalPath), $oDocOriginal, $bCloseOriginal) Then
                    MsgBox($MB_ICONWARNING, "Loi", "Khong mo duoc ban goc.")
                    ContinueLoop
                EndIf
                If Not _ResolveCompareSource(GUICtrlRead($cboRevised), GUICtrlRead($inpRevisedPath), $oDocRevised, $bCloseRevised) Then
                    If $bCloseOriginal And IsObj($oDocOriginal) Then $oDocOriginal.Close(0)
                    MsgBox($MB_ICONWARNING, "Loi", "Khong mo duoc ban sua.")
                    ContinueLoop
                EndIf

                If _IsSameCompareSource($oDocOriginal, $oDocRevised) Then
                    If $bCloseOriginal And IsObj($oDocOriginal) Then $oDocOriginal.Close(0)
                    If $bCloseRevised And IsObj($oDocRevised) Then $oDocRevised.Close(0)
                    MsgBox($MB_ICONWARNING, "Trung nguon", "Ban goc va ban sua dang trung nhau.")
                    ContinueLoop
                EndIf

                _UpdateProgress("Dang tao ban so sanh...")
                $oCompareDoc = _CompareDocumentsObjects($oDocOriginal, $oDocRevised)

                If $bCloseOriginal And IsObj($oDocOriginal) Then $oDocOriginal.Close(0)
                If $bCloseRevised And IsObj($oDocRevised) Then $oDocRevised.Close(0)

                If Not IsObj($oCompareDoc) Then
                    GUICtrlSetData($lblSummary, "Khong tao duoc ban so sanh.")
                    _UpdateProgress("Loi khi so sanh!")
                    MsgBox($MB_ICONWARNING, "Loi", "Khong the so sanh 2 file da chon.")
                    ContinueLoop
                EndIf

                _ActivateCompareDocMarkup($oCompareDoc, True)
                _PopulateCompareRevisionList($listRevisions, $idDetail, $oCompareDoc, $aRevisionMap, $lblSummary)
                _UpdateProgress("Da tao file so sanh!")

            Case $btnRefresh
                If IsObj($oCompareDoc) Then _PopulateCompareRevisionList($listRevisions, $idDetail, $oCompareDoc, $aRevisionMap, $lblSummary)

            Case $btnGoTo
                _GoToSelectedCompareRevision($listRevisions, $oCompareDoc, $aRevisionMap, $idDetail)

            Case $btnOpenCompare
                If IsObj($oCompareDoc) Then $oCompareDoc.Activate()

            Case $btnShowMarkup
                If IsObj($oCompareDoc) Then _ActivateCompareDocMarkup($oCompareDoc, True)

            Case $btnHideMarkup
                If IsObj($oCompareDoc) Then _ActivateCompareDocMarkup($oCompareDoc, False)

            Case $listRevisions
                _UpdateCompareRevisionDetail($listRevisions, $oCompareDoc, $aRevisionMap, $idDetail)
        EndSwitch
    WEnd
EndFunc

Func _GetDefaultCompareDocSelection()
    If Not IsObj($g_oWord) Or $g_oWord.Documents.Count = 0 Then Return "Chinh van ban dang sua"
    For $i = 1 To $g_oWord.Documents.Count
        Local $oDocItem = $g_oWord.Documents.Item($i)
        If IsObj($oDocItem) Then
            If Not IsObj($g_oDoc) Or $oDocItem.FullName <> $g_oDoc.FullName Then Return $i & ". " & $oDocItem.Name
        EndIf
    Next
    Return "Chinh van ban dang sua"
EndFunc

Func _ResolveCompareSource($sSelection, $sPath, ByRef $oDocResolved, ByRef $bCloseAfter)
    $oDocResolved = 0
    $bCloseAfter = False

    Local $sTrimPath = StringStripWS($sPath, 3)
    If $sTrimPath <> "" Then
        If Not FileExists($sTrimPath) Then Return False
        $oDocResolved = $g_oWord.Documents.Open($sTrimPath, False, True)
        If Not IsObj($oDocResolved) Then Return False
        $bCloseAfter = True
        Return True
    EndIf

    $oDocResolved = _GetStyleSourceDocFromSelection($sSelection)
    If Not IsObj($oDocResolved) Then Return False
    Return True
EndFunc

Func _IsSameCompareSource($oDoc1, $oDoc2)
    If Not IsObj($oDoc1) Or Not IsObj($oDoc2) Then Return False

    Local $sFull1 = $oDoc1.FullName
    Local $sFull2 = $oDoc2.FullName
    If $sFull1 <> "" And $sFull2 <> "" And $sFull1 = $sFull2 Then Return True
    If $oDoc1.Name = $oDoc2.Name And $oDoc1.Path = $oDoc2.Path Then Return True
    Return False
EndFunc

Func _CompareDocumentsObjects($oOriginal, $oRevised)
    If Not _CheckConnection() Then Return 0
    If Not IsObj($oOriginal) Or Not IsObj($oRevised) Then Return 0

    Local Const $wdCompareTargetNew = 2
    Local Const $wdGranularityWordLevel = 1
    Local $oCompareDoc = $g_oWord.CompareDocuments($oOriginal, $oRevised, $wdCompareTargetNew, $wdGranularityWordLevel, True)
    If @error Or Not IsObj($oCompareDoc) Then Return 0
    Return $oCompareDoc
EndFunc

Func _GetRevisionTypeText($iType)
    Switch $iType
        Case 1
            Return "Insert"
        Case 2
            Return "Delete"
        Case 3
            Return "Property"
        Case 4
            Return "Paragraph #"
        Case 5
            Return "Display field"
        Case 6
            Return "Reconcile"
        Case 7
            Return "Conflict"
        Case 8
            Return "Style"
        Case 9
            Return "Replace"
        Case 10
            Return "Paragraph prop"
        Case 11
            Return "Table prop"
        Case 12
            Return "Section prop"
        Case 13
            Return "Style def"
        Case 14
            Return "Moved from"
        Case 15
            Return "Moved to"
        Case Else
            Return "Khac (" & $iType & ")"
    EndSwitch
EndFunc

Func _GetRevisionPreviewText($oRevision)
    If Not IsObj($oRevision) Then Return ""
    Local $sText = $oRevision.Range.Text
    $sText = StringReplace($sText, @CR, " ")
    $sText = StringReplace($sText, @LF, " ")
    $sText = StringStripWS($sText, 3)
    If $sText = "" Then $sText = "[" & _GetRevisionTypeText($oRevision.Type) & "]"
    If StringLen($sText) > 90 Then $sText = StringLeft($sText, 90) & "..."
    Return $sText
EndFunc

Func _GetRevisionDetailText($oRevision)
    If Not IsObj($oRevision) Then Return ""
    Local $sText = $oRevision.Range.Text
    $sText = StringReplace($sText, @CR, " ")
    $sText = StringReplace($sText, @LF, " ")
    $sText = StringStripWS($sText, 3)
    If $sText = "" Then $sText = "(Khong co text hien thi)"

    Return "Loai: " & _GetRevisionTypeText($oRevision.Type) & _
        " | Tac gia: " & $oRevision.Author & _
        " | Thoi gian: " & $oRevision.Date & _
        " | Noi dung: " & $sText
EndFunc

Func _PopulateCompareRevisionList($listRevisions, $idDetail, $oCompareDoc, ByRef $aRevisionMap, $lblSummary)
    If Not IsObj($oCompareDoc) Then Return

    _GUICtrlListView_DeleteAllItems($listRevisions)
    GUICtrlSetData($idDetail, "")
    ReDim $aRevisionMap[1]
    $aRevisionMap[0] = 0

    Local $oRevisions = $oCompareDoc.Revisions
    If Not IsObj($oRevisions) Then
        GUICtrlSetData($lblSummary, "Khong doc duoc danh sach thay doi.")
        Return
    EndIf

    Local $iTotal = $oRevisions.Count
    Local $iInsert = 0, $iDelete = 0, $iReplace = 0

    For $i = 1 To $iTotal
        Local $oRevision = $oRevisions.Item($i)
        If Not IsObj($oRevision) Then ContinueLoop

        ReDim $aRevisionMap[UBound($aRevisionMap) + 1]
        $aRevisionMap[UBound($aRevisionMap) - 1] = $i

        GUICtrlCreateListViewItem($i & "|" & _GetRevisionTypeText($oRevision.Type) & "|" & _
            _GetRevisionPreviewText($oRevision) & "|" & $oRevision.Author & "|" & $oRevision.Date, $listRevisions)

        Switch $oRevision.Type
            Case 1, 15
                $iInsert += 1
            Case 2, 14
                $iDelete += 1
            Case 9
                $iReplace += 1
        EndSwitch
    Next

    GUICtrlSetData($lblSummary, "Tong thay doi: " & $iTotal & " | Them: " & $iInsert & " | Xoa: " & $iDelete & " | Thay the: " & $iReplace)
    If $iTotal > 0 Then _UpdateCompareRevisionDetail($listRevisions, $oCompareDoc, $aRevisionMap, $idDetail)
EndFunc

Func _GetSelectedRevisionMapIndex($listRevisions)
    Local $vSel = _GUICtrlListView_GetSelectedIndices($listRevisions, False)
    If $vSel = "" Or $vSel = -1 Then Return -1
    Return Number($vSel) + 1
EndFunc

Func _UpdateCompareRevisionDetail($listRevisions, $oCompareDoc, ByRef $aRevisionMap, $idDetail)
    If Not IsObj($oCompareDoc) Then Return
    Local $iMapIndex = _GetSelectedRevisionMapIndex($listRevisions)
    If $iMapIndex < 1 Or $iMapIndex >= UBound($aRevisionMap) Then Return

    Local $oRevision = $oCompareDoc.Revisions.Item($aRevisionMap[$iMapIndex])
    If Not IsObj($oRevision) Then Return
    GUICtrlSetData($idDetail, _GetRevisionDetailText($oRevision))
EndFunc

Func _GoToSelectedCompareRevision($listRevisions, $oCompareDoc, ByRef $aRevisionMap, $idDetail)
    If Not IsObj($oCompareDoc) Then Return
    Local $iMapIndex = _GetSelectedRevisionMapIndex($listRevisions)
    If $iMapIndex < 1 Or $iMapIndex >= UBound($aRevisionMap) Then
        MsgBox($MB_ICONWARNING, "Chua chon", "Vui long chon 1 thay doi trong danh sach.")
        Return
    EndIf

    Local $oRevision = $oCompareDoc.Revisions.Item($aRevisionMap[$iMapIndex])
    If Not IsObj($oRevision) Then Return

    $oCompareDoc.Activate()
    $oRevision.Range.Select()
    _UpdateCompareRevisionDetail($listRevisions, $oCompareDoc, $aRevisionMap, $idDetail)
EndFunc

Func _ActivateCompareDocMarkup($oCompareDoc, $bShowMarkup = True)
    If Not IsObj($oCompareDoc) Then Return
    $oCompareDoc.Activate()
    If Not IsObj($g_oWord) Or Not IsObj($g_oWord.ActiveWindow) Then Return
    $g_oWord.ActiveWindow.View.ShowRevisionsAndComments = $bShowMarkup
EndFunc

; Merge Documents
Func _MergeDocuments()
    If Not _CheckConnection() Then Return

    Local $sChoice = InputBox("Gop file", "Chon vi tri gop:" & @CRLF & _
        "1 - Chen vao cuoi file" & @CRLF & _
        "2 - Chen vao dau file" & @CRLF & _
        "3 - Chen tai vi tri con tro", "1")
    If @error Then Return

    Local $sPath = FileOpenDialog("Chon file de gop", @ScriptDir, "Word (*.docx;*.doc)", 1)
    If @error Then Return

    _UpdateProgress("Dang gop file...")
    If Not _MergeDocumentFromPath($sPath, $sChoice) Then
        MsgBox($MB_ICONWARNING, "Gop file", "Khong the gop file da chon!")
        Return
    EndIf
    _UpdateProgress("Da gop file!")
    MsgBox($MB_ICONINFORMATION, "Gop file", "Da gop file thanh cong!")
EndFunc

; Split Document
; FIX: Re-activate file goc sau khi tach, dung Selection moi thay vi $oSel cu
Func _SplitDocument()
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Huong dan", "Chon phan van ban can tach truoc!" & @CRLF & @CRLF & _
            "Cach dung:" & @CRLF & _
            "1. Boi den phan van ban can tach" & @CRLF & _
            "2. Nhan nut 'Tach file'" & @CRLF & _
            "3. Chon noi luu file moi")
        Return
    EndIf

    Local $sBaseName = StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "")
    Local $sPath = FileSaveDialog("Luu phan tach", @ScriptDir, "Word (*.docx)", 16, $sBaseName & "_tach.docx")
    If @error Then Return

    _UpdateProgress("Dang tach file...")

    Local $bDeleteFromSource = (MsgBox($MB_YESNO, "Xoa phan da tach?", "Xoa phan van ban da tach khoi file goc?") = $IDYES)
    If Not _SplitRangeToPath($oSel.Range, $sPath, $bDeleteFromSource) Then
        MsgBox($MB_ICONWARNING, "Tach file", "Khong the tach vung da chon!")
        Return
    EndIf

    _UpdateProgress("Da tach file!")
    MsgBox($MB_ICONINFORMATION, "Tach file thanh cong", "Da luu phan tach tai:" & @CRLF & @CRLF & $sPath)
EndFunc

; Protect Document
; FIX: Them error handling cho Unprotect (sai mat khau throw COM error)
Func _ProtectDocument()
    If Not _CheckConnection() Then Return
    
    ; Kiem tra trang thai bao ve hien tai
    Local $bProtected = $g_oDoc.ProtectionType <> -1 ; -1 = wdNoProtection

    If $bProtected Then
        ; File dang duoc bao ve -> hoi bo bao ve
        Local $sPassword = InputBox("Bo bao ve", "Nhap mat khau de bo bao ve:" & @CRLF & _
            "(De trong neu khong co mat khau)", "", "*")
        If @error Then Return
        
        ; QUAN TRONG: Unprotect() throw COM error neu sai mat khau
        ; _ComErrorHandler se bat va set @error = 1
        If Not _RemoveDocumentProtection($sPassword) Then
            MsgBox($MB_ICONWARNING, "Loi", "Mat khau khong dung hoac khong the bo bao ve!")
            Return
        EndIf
        
        If $g_oDoc.ProtectionType = -1 Then
            _UpdateProgress("Da bo bao ve!")
            MsgBox($MB_ICONINFORMATION, "Bo bao ve", "Da bo bao ve file thanh cong!")
        Else
            MsgBox($MB_ICONWARNING, "Loi", "Mat khau khong dung!")
        EndIf
    Else
        ; File chua bao ve -> hoi bao ve
        Local $sChoice = InputBox("Bao ve file", "Chon loai bao ve:" & @CRLF & _
            "1 - Chi doc (Read Only)" & @CRLF & _
            "2 - Chi cho dien form" & @CRLF & _
            "3 - Chi cho comment" & @CRLF & _
            "4 - Chi cho track changes", "1")
        If @error Then Return

        Local $sPassword = InputBox("Mat khau", "Nhap mat khau bao ve (co the de trong):", "", "*")
        If @error Then Return

        Local $iProtectType = 0
        Switch $sChoice
            Case "1"
                $iProtectType = 3 ; wdAllowOnlyReading
            Case "2"
                $iProtectType = 2 ; wdAllowOnlyFormFields
            Case "3"
                $iProtectType = 1 ; wdAllowOnlyComments
            Case "4"
                $iProtectType = 0 ; wdAllowOnlyRevisions
            Case Else
                $iProtectType = 3
        EndSwitch

        If Not _SetDocumentProtection($iProtectType, $sPassword) Then
            MsgBox($MB_ICONWARNING, "Bao ve", "Khong the bao ve file!")
            Return
        EndIf
        _UpdateProgress("Da bao ve file!")
        MsgBox($MB_ICONINFORMATION, "Bao ve", "Da bao ve file thanh cong!")
    EndIf
EndFunc

Func _CompareDocumentsByPath($sPath)
    If Not _CheckConnection() Then Return 0
    If $sPath = "" Or Not FileExists($sPath) Then Return 0

    Local $oDoc2 = $g_oWord.Documents.Open($sPath, False, True)
    If Not IsObj($oDoc2) Then Return 0

    Local $oCompareDoc = _CompareDocumentsObjects($g_oDoc, $oDoc2)
    $oDoc2.Close(0)

    If @error Or Not IsObj($oCompareDoc) Then Return 0
    Return $oCompareDoc
EndFunc

Func _MergeDocumentFromPath($sPath, $sChoice = "1")
    If Not _CheckConnection() Then Return False
    If $sPath = "" Or Not FileExists($sPath) Then Return False

    Local $oRange = 0
    Switch $sChoice
        Case "1"
            $oRange = $g_oDoc.Content
            $oRange.Collapse($WD_COLLAPSE_END)
            $oRange.InsertParagraphAfter()
            $oRange.Collapse($WD_COLLAPSE_END)
        Case "2"
            $oRange = $g_oDoc.Range(0, 0)
        Case "3"
            $oRange = $g_oWord.Selection.Range
        Case Else
            $oRange = $g_oDoc.Content
            $oRange.Collapse($WD_COLLAPSE_END)
    EndSwitch

    If Not IsObj($oRange) Then Return False
    $oRange.InsertFile($sPath)
    Return (Not @error)
EndFunc

Func _SplitRangeToPath($oRange, $sPath, $bDeleteFromSource = False)
    If Not _CheckConnection() Then Return False
    If Not IsObj($oRange) Or $sPath = "" Then Return False

    Local $iStart = $oRange.Start
    Local $iEnd = $oRange.End

    $oRange.Copy()
    Local $oNewDoc = $g_oWord.Documents.Add()
    If Not IsObj($oNewDoc) Then Return False

    $oNewDoc.Content.Paste()
    $oNewDoc.SaveAs2($sPath)
    Local $bOk = (Not @error And FileExists($sPath))
    $oNewDoc.Close(0)

    If IsObj($g_oDoc) Then $g_oDoc.Activate()
    If Not $bOk Then Return False

    If $bDeleteFromSource Then
        Local $oNewRange = $g_oDoc.Range($iStart, $iEnd)
        $oNewRange.Delete()
    EndIf

    Return True
EndFunc

Func _SetDocumentProtection($iProtectType, $sPassword = "")
    If Not _CheckConnection() Then Return False
    $g_oDoc.Protect($iProtectType, False, $sPassword)
    Return (Not @error And $g_oDoc.ProtectionType <> -1)
EndFunc

Func _RemoveDocumentProtection($sPassword = "")
    If Not _CheckConnection() Then Return False
    $g_oDoc.Unprotect($sPassword)
    Return (Not @error And $g_oDoc.ProtectionType = -1)
EndFunc

; Show Document Properties
Func _ShowDocProperties()
    If Not _CheckConnection() Then Return
    
    Local $sMsg = "THUOC TINH FILE:" & @CRLF & @CRLF
    $sMsg &= "Ten: " & $g_oDoc.Name & @CRLF
    $sMsg &= "Duong dan: " & $g_oDoc.Path & @CRLF
    $sMsg &= "So trang: " & $g_oDoc.ComputeStatistics(2) & @CRLF
    $sMsg &= "So tu: " & $g_oDoc.ComputeStatistics(0) & @CRLF
    
    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "Thuoc tinh", $sMsg)
EndFunc

; Clean Document
Func _CleanDocument()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Don dep file", _
        "Se thuc hien:" & @CRLF & _
        "- Xoa comments" & @CRLF & _
        "- Xoa hyperlinks" & @CRLF & _
        "- Chap nhan thay doi" & @CRLF & _
        "Tiep tuc?") <> $IDYES Then Return
    
    _UpdateProgress("Dang don dep...")

    _CleanDocumentCore()
    _UpdateProgress("Da don dep file!")
EndFunc

Func _CleanDocumentCore()
    If Not _CheckConnection() Then Return False

    While $g_oDoc.Comments.Count > 0
        $g_oDoc.Comments.Item(1).Delete()
    WEnd

    While $g_oDoc.Hyperlinks.Count > 0
        $g_oDoc.Hyperlinks.Item(1).Delete()
    WEnd

    $g_oDoc.AcceptAllRevisions()
    Return (Not @error)
EndFunc
