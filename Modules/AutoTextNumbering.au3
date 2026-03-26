; ============================================
; AUTOTEXTNUMBERING.AU3 - Module Danh so theo mau text
; Tach rieng de tranh xung dot voi Tools.au3
; ============================================

#include-once

Func _ShowAutoTextNumberingDialog()
    If Not _CheckConnection() Then Return

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

    Local $hPopup = GUICreate("Danh so caption theo mau text", 760, 430, -1, -1, BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xF7F9FB, $hPopup)

    GUICtrlCreateLabel("Nhap 1 caption mau de danh so lai hang loat theo text va style.", 20, 15, 700, 20)
    GUICtrlSetFont(-1, 9, 600)

    GUICtrlCreateLabel("Loai doi tuong:", 30, 48, 100, 20)
    Local $cboType = GUICtrlCreateCombo("", 140, 43, 150, 24, $CBS_DROPDOWNLIST)
    GUICtrlSetData($cboType, "Hinh|Bang|Bieu do|So do", "Hinh")

    GUICtrlCreateLabel("Mau caption:", 30, 83, 100, 20)
    Local $inpSample = GUICtrlCreateInput("Hinh 1.1 Anh edge computing", 140, 80, 560, 24)

    GUICtrlCreateLabel("Nguon style:", 30, 120, 100, 20)
    Local $cboStyleSource = GUICtrlCreateCombo("", 140, 115, 360, 24, $CBS_DROPDOWNLIST)
    GUICtrlSetData($cboStyleSource, $sDocList, "Chinh van ban dang sua")
    Local $btnRefreshSource = GUICtrlCreateButton("Lam moi", 510, 114, 80, 26)

    GUICtrlCreateLabel("Loc style:", 30, 156, 100, 20)
    Local $inpStyleFilter = GUICtrlCreateInput("", 140, 152, 250, 24)
    Local $btnApplyFilter = GUICtrlCreateButton("Loc", 400, 151, 55, 26)
    Local $btnClearFilter = GUICtrlCreateButton("Bo loc", 462, 151, 70, 26)

    GUICtrlCreateLabel("Style caption:", 30, 192, 100, 20)
    Local $cboStyle = GUICtrlCreateCombo("", 140, 188, 360, 24, $CBS_DROPDOWNLIST)
    Local $sPreferredStyle = _ResolvePreferredStyleName($sStyles, "Caption")
    GUICtrlSetData($cboStyle, $sStyles, $sPreferredStyle)
    GUICtrlCreateLabel("Co the chon style tu file hien tai hoac tu 1 file Word dang mo khac.", 140, 217, 520, 18)

    Local $chkStyleOnly = GUICtrlCreateCheckbox(" Chi quet cac paragraph dang dung style da chon", 140, 248, 330, 22)
    GUICtrlSetState($chkStyleOnly, $GUI_CHECKED)
    Local $chkApplyStyle = GUICtrlCreateCheckbox(" Gan style da chon cho caption tim thay sau khi danh so", 140, 274, 360, 22)
    GUICtrlSetState($chkApplyStyle, $GUI_CHECKED)

    Local $lblHint = GUICtrlCreateLabel( _
        "Vi du: mau 'Hinh 1.1 Anh edge computing' se quet nhom 'Hinh 1.x' va danh lai thanh Hinh 1.1, Hinh 1.2, Hinh 1.3..." & _
        @CRLF & "Tieu de tung caption duoc giu nguyen, chi thay block so o dau dong.", 30, 308, 690, 42)
    GUICtrlSetColor($lblHint, 0x555555)

    Local $btnApply = GUICtrlCreateButton("Danh so hang loat", 430, 370, 140, 34, $BS_DEFPUSHBUTTON)
    GUICtrlSetBkColor($btnApply, 0x27AE60)
    Local $btnCancel = GUICtrlCreateButton("Dong", 580, 370, 120, 34)

    GUISetState(@SW_SHOW, $hPopup)

    While 1
        Local $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                GUIDelete($hPopup)
                Return
            Case $cboType
                GUICtrlSetData($inpSample, _GetAutoTextNumberingSampleForType(GUICtrlRead($cboType)))
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

                _RefreshAutoTextNumberingStyleCombo($cboStyle, $sStyles, GUICtrlRead($inpStyleFilter), GUICtrlRead($cboStyle))
            Case $btnApplyFilter, $inpStyleFilter, $btnClearFilter
                If $iMsg = $btnClearFilter Then GUICtrlSetData($inpStyleFilter, "")
                _RefreshAutoTextNumberingStyleCombo($cboStyle, $sStyles, GUICtrlRead($inpStyleFilter), GUICtrlRead($cboStyle))
            Case $btnApply
                Local $sSample = GUICtrlRead($inpSample)
                Local $sStyleName = GUICtrlRead($cboStyle)
                If StringStripWS($sSample, 3) = "" Then
                    MsgBox($MB_ICONWARNING, "Thieu mau", "Vui long nhap caption mau can danh so.")
                    ContinueLoop
                EndIf
                If StringStripWS($sStyleName, 3) = "" Then
                    MsgBox($MB_ICONWARNING, "Chua chon style", "Vui long chon style caption de quet.")
                    ContinueLoop
                EndIf

                GUIDelete($hPopup)
                _ApplyAutoTextNumberingFromSample($sSample, $oStyleSourceDoc, $sStyleName, _
                    GUICtrlRead($chkStyleOnly) = $GUI_CHECKED, GUICtrlRead($chkApplyStyle) = $GUI_CHECKED)
                Return
        EndSwitch
    WEnd
EndFunc

Func _RefreshAutoTextNumberingStyleCombo($cboStyle, $sStyles, $sFilter, $sCurrentSelection = "")
    Local $sFiltered = _FilterStyleList($sStyles, $sFilter, "Caption")
    If $sFiltered = "" Then $sFiltered = $sStyles

    Local $sPreferred = _ResolvePreferredStyleName($sFiltered, $sCurrentSelection, "Caption")
    GUICtrlSetData($cboStyle, "", "")
    GUICtrlSetData($cboStyle, $sFiltered, $sPreferred)
EndFunc

Func _GetAutoTextNumberingSampleForType($sType)
    Switch StringLower(StringStripWS($sType, 3))
        Case "bang"
            Return "Bang 1.1 Bang ket qua thuc nghiem"
        Case "bieu do"
            Return "Bieu do 1.1 Bieu do tai luong"
        Case "so do"
            Return "So do 1.1 So do kien truc he thong"
        Case Else
            Return "Hinh 1.1 Anh edge computing"
    EndSwitch
EndFunc

Func _ApplyAutoTextNumberingFromSample($sSample, $oStyleSourceDoc, $sStyleName, $bStyleOnly = True, $bApplyStyle = True)
    If Not _CheckConnection() Then Return

    Local $sLabel = "", $sFixedPrefix = "", $sSeparator = "", $sSampleTitle = ""
    Local $iStartNumber = 0
    If Not _ParseAutoTextCaptionSample($sSample, $sLabel, $sFixedPrefix, $iStartNumber, $sSeparator, $sSampleTitle) Then
        MsgBox($MB_ICONWARNING, "Mau khong hop le", _
            "Khong phan tich duoc caption mau." & @CRLF & @CRLF & _
            "Vi du hop le:" & @CRLF & _
            "- Hinh 1.1 Anh edge computing" & @CRLF & _
            "- Bang 2.3 Bang so lieu" & @CRLF & _
            "- Bieu do 3.1: Luu luong mang")
        Return
    EndIf

    Local $bNeedStyle = $bStyleOnly Or $bApplyStyle
    If $bNeedStyle And Not _EnsureAutoTextNumberingStyleAvailable($oStyleSourceDoc, $g_oDoc, $sStyleName) Then
        MsgBox($MB_ICONWARNING, "Khong co style", "Khong copy/khong tim thay style da chon trong van ban hien tai.")
        Return
    EndIf

    Local $oParas = $g_oDoc.Paragraphs
    If Not IsObj($oParas) Or $oParas.Count = 0 Then Return

    Local $iCount = 0
    Local $sGroupDisplay = _GetAutoTextCaptionGroupDisplay($sFixedPrefix)
    Local $sLog = "DANH SO CAPTION THEO MAU" & @CRLF & @CRLF
    $sLog &= "Mau: " & $sSample & @CRLF
    $sLog &= "Nhan dien nhom: " & $sLabel & " " & $sGroupDisplay & @CRLF
    $sLog &= "Style: " & $sStyleName & @CRLF & @CRLF

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        If $bStyleOnly And Not _ParagraphUsesStyleName($oPara, $sStyleName) Then ContinueLoop

        Local $sLeading = "", $sParaLabel = "", $sNumberBlock = "", $sParaSeparator = "", $sTitle = "", $sEol = ""
        If Not _ParseAutoTextCaptionLine($oPara.Range.Text, $sLeading, $sParaLabel, $sNumberBlock, $sParaSeparator, $sTitle, $sEol) Then ContinueLoop
        If Not _AutoTextCaptionMatchesGroup($sParaLabel, $sLabel, $sNumberBlock, $sFixedPrefix) Then ContinueLoop

        $iCount += 1
        Local $sNewNumber = _BuildAutoTextCaptionNumber($sFixedPrefix, $iCount)
        Local $sNewText = $sLeading & $sLabel & " " & $sNewNumber & $sSeparator & $sTitle & $sEol
        $oPara.Range.Text = $sNewText
        If $bApplyStyle Then $oPara.Range.Style = $sStyleName
    Next

    If $iCount = 0 Then
        MsgBox($MB_ICONINFORMATION, "Khong tim thay", _
            "Khong tim thay caption nao phu hop." & @CRLF & @CRLF & _
            "Meo:" & @CRLF & _
            "- Kiem tra lai mau caption" & @CRLF & _
            "- Kiem tra style da chon" & @CRLF & _
            "- Neu caption chua gan style, bo tick 'Chi quet cac paragraph dang dung style da chon'")
        Return
    EndIf

    $sLog &= "Da danh so lai: " & $iCount & " caption" & @CRLF
    $sLog &= "Dinh dang so moi: " & $sLabel & " " & _BuildAutoTextCaptionNumber($sFixedPrefix, 1)
    _LogPreview($sLog)
    _UpdateProgress("Da danh so lai " & $iCount & " caption theo mau!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", _
        "Da danh so lai " & $iCount & " caption." & @CRLF & @CRLF & _
        "Nhom: " & $sLabel & " " & $sGroupDisplay & @CRLF & _
        "Style: " & $sStyleName)
EndFunc

Func _EnsureAutoTextNumberingStyleAvailable($oSourceDoc, $oTargetDoc, $sStyleName)
    If Not IsObj($oTargetDoc) Then Return False
    If _DocumentHasParagraphStyle($oTargetDoc, $sStyleName) Then Return True
    If Not IsObj($oSourceDoc) Then Return False
    If $oSourceDoc.FullName = $oTargetDoc.FullName Then Return False
    Return _CopySingleStyle($oSourceDoc, $oTargetDoc, $sStyleName)
EndFunc

Func _DocumentHasParagraphStyle($oDoc, $sStyleName)
    If Not IsObj($oDoc) Then Return False
    If StringStripWS($sStyleName, 3) = "" Then Return False

    Local $oStyle = 0
    Local $bPrevMute = $g_bMuteComErrors
    $g_bMuteComErrors = True
    $oStyle = $oDoc.Styles($sStyleName)
    $g_bMuteComErrors = $bPrevMute
    If @error Or Not IsObj($oStyle) Then Return False
    Return True
EndFunc

Func _ParagraphUsesStyleName($oPara, $sStyleName)
    If Not IsObj($oPara) Then Return False
    If StringStripWS($sStyleName, 3) = "" Then Return False

    Local $sExpected = StringLower(StringStripWS($sStyleName, 3))
    Local $vStyle = 0
    Local $bPrevMute = $g_bMuteComErrors
    $g_bMuteComErrors = True
    $vStyle = $oPara.Range.Style
    $g_bMuteComErrors = $bPrevMute
    If @error Then Return False

    If IsObj($vStyle) Then
        Return _AutoTextStyleMatchesName($vStyle, $sExpected)
    EndIf

    Return (StringLower(StringStripWS($vStyle, 3)) = $sExpected)
EndFunc

Func _AutoTextStyleMatchesName($oStyle, $sExpectedLower)
    If Not IsObj($oStyle) Then Return False

    Local $bPrevMute = $g_bMuteComErrors
    Local $sNameLocal = ""
    Local $sName = ""

    $g_bMuteComErrors = True
    $sNameLocal = $oStyle.NameLocal
    $sName = $oStyle.Name
    $g_bMuteComErrors = $bPrevMute

    If StringLower(StringStripWS($sNameLocal, 3)) = $sExpectedLower Then Return True
    If StringLower(StringStripWS($sName, 3)) = $sExpectedLower Then Return True
    Return False
EndFunc

Func _ParseAutoTextCaptionSample($sSample, ByRef $sLabel, ByRef $sFixedPrefix, ByRef $iStartNumber, ByRef $sSeparator, ByRef $sTitle)
    Local $sLeading = "", $sNumberBlock = "", $sEol = ""
    If Not _ParseAutoTextCaptionLine($sSample, $sLeading, $sLabel, $sNumberBlock, $sSeparator, $sTitle, $sEol) Then Return False

    Local $aParts = StringSplit($sNumberBlock, ".", 2)
    If Not IsArray($aParts) Or UBound($aParts) = 0 Then Return False

    $iStartNumber = Number($aParts[UBound($aParts) - 1])
    If $iStartNumber <= 0 Then Return False

    $sFixedPrefix = ""
    For $i = 0 To UBound($aParts) - 2
        If $sFixedPrefix <> "" Then $sFixedPrefix &= "."
        $sFixedPrefix &= $aParts[$i]
    Next

    Return True
EndFunc

Func _ParseAutoTextCaptionLine($sText, ByRef $sLeading, ByRef $sLabel, ByRef $sNumberBlock, ByRef $sSeparator, ByRef $sTitle, ByRef $sEol)
    $sLeading = ""
    $sLabel = ""
    $sNumberBlock = ""
    $sSeparator = ""
    $sTitle = ""
    $sEol = ""

    Local $sBody = $sText
    If StringRight($sBody, 2) = @CRLF Then
        $sEol = @CRLF
        $sBody = StringTrimRight($sBody, 2)
    ElseIf StringRight($sBody, 1) = @CR Then
        $sEol = @CR
        $sBody = StringTrimRight($sBody, 1)
    ElseIf StringRight($sBody, 1) = @LF Then
        $sEol = @LF
        $sBody = StringTrimRight($sBody, 1)
    EndIf

    Local $aMatch = StringRegExp($sBody, "^(\s*)(.+?)\s+((?:\d+\s*\.\s*)*\d+)(\s*(?:[\.\):\-])?\s+)(.+?)\s*$", 1)
    If @error Or Not IsArray($aMatch) Or UBound($aMatch) < 5 Then Return False

    $sLeading = $aMatch[0]
    $sLabel = _NormalizeAutoTextCaptionLabel($aMatch[1])
    $sNumberBlock = StringRegExpReplace(StringStripWS($aMatch[2], 3), "\s*\.\s*", ".")
    $sSeparator = _NormalizeAutoTextCaptionSeparator($aMatch[3])
    $sTitle = StringStripWS($aMatch[4], 3)
    Return ($sLabel <> "" And $sNumberBlock <> "" And $sSeparator <> "" And $sTitle <> "")
EndFunc

Func _NormalizeAutoTextCaptionLabel($sLabel)
    Local $sResult = StringStripWS($sLabel, 3)
    $sResult = StringRegExpReplace($sResult, "\s+", " ")
    Return $sResult
EndFunc

Func _NormalizeAutoTextCaptionSeparator($sSeparator)
    Local $sClean = StringRegExpReplace($sSeparator, "\s+", " ")
    $sClean = StringStripWS($sClean, 3)
    If $sClean = "" Then Return " "
    If StringRight($sClean, 1) <> " " Then $sClean &= " "
    Return $sClean
EndFunc

Func _AutoTextCaptionMatchesGroup($sActualLabel, $sExpectedLabel, $sNumberBlock, $sFixedPrefix)
    If StringLower(_NormalizeAutoTextCaptionLabel($sActualLabel)) <> StringLower(_NormalizeAutoTextCaptionLabel($sExpectedLabel)) Then Return False

    If $sFixedPrefix = "" Then
        Return StringRegExp($sNumberBlock, "^\d+$")
    EndIf

    Return StringLeft($sNumberBlock, StringLen($sFixedPrefix) + 1) = ($sFixedPrefix & ".")
EndFunc

Func _BuildAutoTextCaptionNumber($sFixedPrefix, $iSequence)
    If $sFixedPrefix = "" Then Return String($iSequence)
    Return $sFixedPrefix & "." & $iSequence
EndFunc

Func _GetAutoTextCaptionGroupDisplay($sFixedPrefix)
    If $sFixedPrefix = "" Then Return "x"
    Return $sFixedPrefix & ".x"
EndFunc
