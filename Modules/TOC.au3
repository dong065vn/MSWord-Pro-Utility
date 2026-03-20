; ============================================
; TOC.AU3 - Module Muc luc & TLTK
; ============================================

#include-once

; Create TOC
Func _CreateTOC()
    If Not _CheckConnection() Then Return
    
    Local $iLevels = Number(GUICtrlRead($g_cboTOCLevels))
    Local $bHyperlink = (GUICtrlRead($g_chkTOCHyperlink) = $GUI_CHECKED)
    
    _UpdateProgress("Dang tao muc luc...")
    
    ; Chen tieu de
    Local $oRange = $g_oDoc.Range(0, 0)
    $oRange.InsertBefore("MUC LUC" & @CR & @CR)
    $oRange.Font.Bold = True
    $oRange.Font.Size = 14
    $oRange.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
    
    ; Tao TOC
    Local $oTOCRange = $g_oDoc.Range($oRange.End, $oRange.End)
    $g_oDoc.TablesOfContents.Add($oTOCRange, True, 1, $iLevels, False, "", True, True, "", $bHyperlink, False, True)
    
    _UpdateProgress("Da tao muc luc!")
EndFunc

; Update TOC
Func _UpdateTOC()
    If Not _CheckConnection() Then Return
    If $g_oDoc.TablesOfContents.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co muc luc!")
        Return
    EndIf
    $g_oDoc.TablesOfContents.Item(1).Update()
    _UpdateProgress("Da cap nhat muc luc!")
EndFunc

; Delete TOC
Func _DeleteTOC()
    If Not _CheckConnection() Then Return
    If $g_oDoc.TablesOfContents.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co muc luc!")
        Return
    EndIf
    $g_oDoc.TablesOfContents.Item(1).Delete()
    _UpdateProgress("Da xoa muc luc!")
EndFunc

; Fix All TOC Styles
Func _FixAllTOCStyles()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang fix TOC styles...")
    
    For $i = 1 To 3
        _FixTOCStyle($i)
    Next
    
    If $g_oDoc.TablesOfContents.Count > 0 Then
        $g_oDoc.TablesOfContents.Item(1).Update()
    EndIf
    
    _UpdateProgress("Da fix TOC styles!")
EndFunc

; Fix TOC Style by level
Func _FixTOCStyle($iLevel)
    If Not _CheckConnection() Then Return
    
    Local $sStyleName = "TOC " & $iLevel
    Local $oStyle = $g_oDoc.Styles($sStyleName)
    If Not IsObj($oStyle) Then
        ; Try Vietnamese style name
        $oStyle = $g_oDoc.Styles("Muc luc " & $iLevel)
        If Not IsObj($oStyle) Then
            _UpdateProgress("Khong tim thay style TOC " & $iLevel)
            Return
        EndIf
    EndIf
    
    ; Get tab leader
    Local $iLeader = $WD_TAB_LEADER_DOTS
    Local $sLeader = GUICtrlRead($g_cboTabLeader)
    If StringInStr($sLeader, "Dashes") Then $iLeader = 3
    If StringInStr($sLeader, "Line") Then $iLeader = 4
    If StringInStr($sLeader, "Khong") Then $iLeader = 0
    
    ; Clear existing tabs
    $oStyle.ParagraphFormat.TabStops.ClearAll()
    
    ; Add right-aligned tab with leader
    Local $fTabPos = $g_oDoc.PageSetup.PageWidth - $g_oDoc.PageSetup.LeftMargin - $g_oDoc.PageSetup.RightMargin
    $oStyle.ParagraphFormat.TabStops.Add($fTabPos, $WD_TAB_ALIGN_RIGHT, $iLeader)
    
    _UpdateProgress("Da fix TOC " & $iLevel)
EndFunc

; Preview TOC Styles
Func _PreviewTOCStyles()
    If Not _CheckConnection() Then Return
    
    Local $sMsg = "TOC STYLES HIEN TAI:" & @CRLF & @CRLF
    
    For $i = 1 To 3
        Local $sStyleName = "TOC " & $i
        Local $oStyle = $g_oDoc.Styles($sStyleName)
        
        ; Try Vietnamese name if English not found
        If Not IsObj($oStyle) Then
            $oStyle = $g_oDoc.Styles("Muc luc " & $i)
        EndIf
        
        If IsObj($oStyle) Then
            $sMsg &= $sStyleName & ":" & @CRLF
            $sMsg &= "  Font: " & $oStyle.Font.Name & " " & $oStyle.Font.Size & "pt" & @CRLF
            $sMsg &= "  Indent: " & Round($oStyle.ParagraphFormat.LeftIndent / $CM_TO_POINTS, 2) & " cm" & @CRLF
        Else
            $sMsg &= $sStyleName & ": (khong tim thay)" & @CRLF
        EndIf
    Next
    
    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "TOC Styles", $sMsg)
EndFunc


; === TLTK (References) ===

; Add Reference
Func _AddReference()
    If Not _CheckConnection() Then Return
    
    Local $sAuthor = GUICtrlRead($g_inputAuthor)
    Local $sYear = GUICtrlRead($g_inputYear)
    Local $sTitle = GUICtrlRead($g_inputTitle)
    Local $sSource = GUICtrlRead($g_inputSource)
    Local $sURL = GUICtrlRead($g_inputURL)
    
    If $sAuthor = "" Or $sTitle = "" Then
        MsgBox($MB_ICONWARNING, "Loi", "Nhap Tac gia va Tieu de!")
        Return
    EndIf
    
    ; Format theo style
    Local $sStyle = GUICtrlRead($g_cboCitationStyle)
    Local $sRef = ""
    
    Switch $sStyle
        Case "APA 7th"
            $sRef = $sAuthor & " (" & $sYear & "). " & $sTitle & ". "
            If $sSource <> "" Then $sRef &= $sSource & ". "
            If $sURL <> "" Then $sRef &= "Retrieved from " & $sURL
        Case "IEEE"
            $sRef = "[#] " & $sAuthor & ", """ & $sTitle & ","" "
            If $sSource <> "" Then $sRef &= $sSource & ", "
            $sRef &= $sYear & "."
        Case Else
            $sRef = $sAuthor & ". " & $sTitle & ". " & $sSource & ", " & $sYear & "."
    EndSwitch
    
    ; Chen vao cuoi document
    Local $oRange = $g_oDoc.Content
    $oRange.Collapse($WD_COLLAPSE_END)
    $oRange.InsertParagraphAfter()
    $oRange.Collapse($WD_COLLAPSE_END)
    $oRange.Text = $sRef
    
    _UpdateProgress("Da them TLTK!")
EndFunc

; Insert Citation
Func _InsertCitation()
    If Not _CheckConnection() Then Return
    
    Local $sAuthor = GUICtrlRead($g_inputAuthor)
    Local $sYear = GUICtrlRead($g_inputYear)
    
    If $sAuthor = "" Then
        MsgBox($MB_ICONWARNING, "Loi", "Nhap Tac gia!")
        Return
    EndIf
    
    Local $sCite = "(" & $sAuthor
    If $sYear <> "" Then $sCite &= ", " & $sYear
    $sCite &= ")"
    
    $g_oWord.Selection.TypeText($sCite)
    _UpdateProgress("Da chen trich dan!")
EndFunc

; Sort References (A-Z)
Func _SortReferences()
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Huong dan", "Chon danh sach TLTK can sap xep!" & @CRLF & @CRLF & _
            "Cach dung:" & @CRLF & _
            "1. Boi den toan bo danh sach TLTK" & @CRLF & _
            "2. Nhan nut 'A-Z' de sap xep")
        Return
    EndIf
    ; Sap xep A-Z theo doan van
    $oSel.Sort(True, 1, "", 0, 0, "", 0, 0, "", False, False, 0, 1, 1)
    _UpdateProgress("Da sap xep TLTK theo A-Z!")
    MsgBox($MB_ICONINFORMATION, "Sap xep", "Da sap xep danh sach TLTK theo thu tu A-Z!")
EndFunc

; Clear Reference Form
Func _ClearRefForm()
    GUICtrlSetData($g_inputAuthor, "")
    GUICtrlSetData($g_inputYear, "")
    GUICtrlSetData($g_inputTitle, "")
    GUICtrlSetData($g_inputSource, "")
    GUICtrlSetData($g_inputURL, "")
    _UpdateProgress("Da xoa form!")
EndFunc

; Format References
Func _FormatReferences()
    If Not _CheckConnection() Then Return
    Local $sStyle = GUICtrlRead($g_cboCitationStyle)
    _UpdateProgress("Dang dinh dang TLTK theo " & $sStyle & "...")

    ; Tim va dinh dang phan TLTK
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Huong dan", "Chon phan Tai lieu tham khao can dinh dang!" & @CRLF & @CRLF & _
            "Cach dung:" & @CRLF & _
            "1. Boi den phan TLTK trong Word" & @CRLF & _
            "2. Chon kieu dinh dang (APA, IEEE...)" & @CRLF & _
            "3. Nhan nut 'Dinh dang TLTK'")
        Return
    EndIf

    ; Ap dung dinh dang cho vung chon
    Local $oRange = $oSel.Range
    $oRange.Font.Name = "Times New Roman"
    $oRange.Font.Size = 13
    $oRange.ParagraphFormat.Alignment = $WD_ALIGN_JUSTIFY
    $oRange.ParagraphFormat.LineSpacingRule = $WD_LINE_SPACE_MULTIPLE
    $oRange.ParagraphFormat.LineSpacing = 18
    $oRange.ParagraphFormat.SpaceBefore = 0
    $oRange.ParagraphFormat.SpaceAfter = 6

    ; Thut le treo (hanging indent) cho TLTK
    $oRange.ParagraphFormat.FirstLineIndent = -0.5 * $CM_TO_POINTS
    $oRange.ParagraphFormat.LeftIndent = 0.5 * $CM_TO_POINTS

    _UpdateProgress("Da dinh dang TLTK theo " & $sStyle & "!")
    MsgBox($MB_ICONINFORMATION, "Dinh dang TLTK", "Da ap dung dinh dang " & $sStyle & @CRLF & @CRLF & _
        "- Font: Times New Roman 13pt" & @CRLF & _
        "- Gian dong: 1.5" & @CRLF & _
        "- Thut le treo: 0.5cm")
EndFunc
