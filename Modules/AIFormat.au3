; ============================================
; AIFORMAT.AU3 - Module AI Format
; Chuyen noi dung tu ChatGPT/Gemini/Claude/Copilot
; sang chuan do an dai hoc VN
; ============================================

#include-once

; === SETTINGS (co the tuy chinh qua dialog) ===
; Chuan do an dai hoc VN - mac dinh
Global $AI_FONT_NAME = "Times New Roman"
Global $AI_FONT_SIZE = 13
Global $AI_FONT_SIZE_H1 = 14
Global $AI_FONT_SIZE_H2 = 13
Global $AI_FONT_SIZE_H3 = 13
Global $AI_LINE_SPACING = 1.5
Global $AI_FIRST_INDENT = 1.27 ; cm
Global $AI_MARGIN_LEFT = 3    ; cm
Global $AI_MARGIN_RIGHT = 2   ; cm
Global $AI_MARGIN_TOP = 2     ; cm
Global $AI_MARGIN_BOTTOM = 2  ; cm
Global $AI_PARA_SPACING_BEFORE = 0
Global $AI_PARA_SPACING_AFTER = 6 ; pt

; INI file path
Global $AI_SETTINGS_FILE = @ScriptDir & "\AIFormat_Settings.ini"

; Load settings tu INI (goi khi khoi dong)
_AI_LoadSettings()

Func _AI_LoadSettings()
    If Not FileExists($AI_SETTINGS_FILE) Then Return

    $AI_FONT_NAME = IniRead($AI_SETTINGS_FILE, "Format", "FontName", $AI_FONT_NAME)
    $AI_FONT_SIZE = Number(IniRead($AI_SETTINGS_FILE, "Format", "FontSize", $AI_FONT_SIZE))
    $AI_FONT_SIZE_H1 = Number(IniRead($AI_SETTINGS_FILE, "Format", "H1Size", $AI_FONT_SIZE_H1))
    $AI_FONT_SIZE_H2 = Number(IniRead($AI_SETTINGS_FILE, "Format", "H2Size", $AI_FONT_SIZE_H2))
    $AI_FONT_SIZE_H3 = Number(IniRead($AI_SETTINGS_FILE, "Format", "H3Size", $AI_FONT_SIZE_H3))
    $AI_LINE_SPACING = Number(IniRead($AI_SETTINGS_FILE, "Format", "LineSpacing", $AI_LINE_SPACING))
    $AI_FIRST_INDENT = Number(IniRead($AI_SETTINGS_FILE, "Format", "FirstIndent", $AI_FIRST_INDENT))
    $AI_MARGIN_LEFT = Number(IniRead($AI_SETTINGS_FILE, "Format", "MarginLeft", $AI_MARGIN_LEFT))
    $AI_MARGIN_RIGHT = Number(IniRead($AI_SETTINGS_FILE, "Format", "MarginRight", $AI_MARGIN_RIGHT))
    $AI_MARGIN_TOP = Number(IniRead($AI_SETTINGS_FILE, "Format", "MarginTop", $AI_MARGIN_TOP))
    $AI_MARGIN_BOTTOM = Number(IniRead($AI_SETTINGS_FILE, "Format", "MarginBottom", $AI_MARGIN_BOTTOM))
    $AI_PARA_SPACING_BEFORE = Number(IniRead($AI_SETTINGS_FILE, "Format", "SpaceBefore", $AI_PARA_SPACING_BEFORE))
    $AI_PARA_SPACING_AFTER = Number(IniRead($AI_SETTINGS_FILE, "Format", "SpaceAfter", $AI_PARA_SPACING_AFTER))
EndFunc

; ========================================
; SETTINGS DIALOG
; ========================================
Func _AI_ShowSettings()
    Local $hSettings = GUICreate("Cai dat AI Format - Chuan do an", 450, 520, -1, -1)
    GUISetBkColor(0xF0F4F8, $hSettings)

    ; Title
    GUICtrlCreateLabel("CAI DAT DINH DANG DO AN", 20, 10, 410, 25, $SS_CENTER)
    GUICtrlSetFont(-1, 14, 800, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x2C3E50)

    ; === Font ===
    GUICtrlCreateGroup(" Font chu ", 15, 40, 420, 80)
    GUICtrlCreateLabel("Font:", 30, 65, 40, 20)
    Local $inpFont = GUICtrlCreateInput($AI_FONT_NAME, 75, 62, 180, 22)
    GUICtrlCreateLabel("Size body:", 270, 65, 60, 20)
    Local $inpSize = GUICtrlCreateInput($AI_FONT_SIZE, 335, 62, 45, 22)
    GUICtrlCreateLabel("H1:", 30, 92, 25, 20)
    Local $inpH1 = GUICtrlCreateInput($AI_FONT_SIZE_H1, 55, 89, 40, 22)
    GUICtrlCreateLabel("H2:", 110, 92, 25, 20)
    Local $inpH2 = GUICtrlCreateInput($AI_FONT_SIZE_H2, 135, 89, 40, 22)
    GUICtrlCreateLabel("H3:", 190, 92, 25, 20)
    Local $inpH3 = GUICtrlCreateInput($AI_FONT_SIZE_H3, 215, 89, 40, 22)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Spacing ===
    GUICtrlCreateGroup(" Khoang cach ", 15, 125, 420, 70)
    GUICtrlCreateLabel("Line spacing:", 30, 150, 75, 20)
    Local $inpLineSpacing = GUICtrlCreateInput($AI_LINE_SPACING, 110, 147, 50, 22)
    GUICtrlCreateLabel("Before (pt):", 180, 150, 70, 20)
    Local $inpBefore = GUICtrlCreateInput($AI_PARA_SPACING_BEFORE, 255, 147, 45, 22)
    GUICtrlCreateLabel("After (pt):", 315, 150, 60, 20)
    Local $inpAfter = GUICtrlCreateInput($AI_PARA_SPACING_AFTER, 380, 147, 45, 22)
    GUICtrlCreateLabel("Thut dau dong (cm):", 30, 172, 110, 20)
    Local $inpIndent = GUICtrlCreateInput($AI_FIRST_INDENT, 145, 169, 50, 22)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Margins ===
    GUICtrlCreateGroup(" Le trang (cm) ", 15, 200, 420, 60)
    GUICtrlCreateLabel("Trai:", 30, 225, 30, 20)
    Local $inpLeft = GUICtrlCreateInput($AI_MARGIN_LEFT, 62, 222, 40, 22)
    GUICtrlCreateLabel("Phai:", 120, 225, 35, 20)
    Local $inpRight = GUICtrlCreateInput($AI_MARGIN_RIGHT, 157, 222, 40, 22)
    GUICtrlCreateLabel("Tren:", 215, 225, 35, 20)
    Local $inpTop = GUICtrlCreateInput($AI_MARGIN_TOP, 252, 222, 40, 22)
    GUICtrlCreateLabel("Duoi:", 310, 225, 35, 20)
    Local $inpBottom = GUICtrlCreateInput($AI_MARGIN_BOTTOM, 347, 222, 40, 22)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Preset buttons ===
    GUICtrlCreateGroup(" Mau co san ", 15, 265, 420, 55)
    Local $btnPresetDH = GUICtrlCreateButton("Dai hoc (13pt)", 30, 283, 120, 28)
    Local $btnPresetTHS = GUICtrlCreateButton("Thac si (14pt)", 160, 283, 120, 28)
    Local $btnPresetDefault = GUICtrlCreateButton("Mac dinh", 290, 283, 100, 28)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Preview ===
    GUICtrlCreateGroup(" Xem truoc ", 15, 325, 420, 120)
    Local $lblPreview = GUICtrlCreateLabel("", 25, 345, 400, 90)
    GUICtrlSetFont($lblPreview, 9, 400, 0, "Segoe UI")
    _AI_UpdatePreviewLabel($lblPreview, $inpFont, $inpSize, $inpLineSpacing, $inpIndent, $inpLeft, $inpRight, $inpTop, $inpBottom)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Buttons ===
    Local $btnSave = GUICtrlCreateButton("Luu cai dat", 50, 455, 150, 35)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    Local $btnCancel = GUICtrlCreateButton("Huy", 250, 455, 150, 35)
    GUICtrlSetFont(-1, 11, 400, 0, "Segoe UI")

    GUISetState(@SW_SHOW, $hSettings)

    ; Event loop
    While 1
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case -3 ; GUI_EVENT_CLOSE
                ExitLoop

            Case $btnCancel
                ExitLoop

            Case $btnPresetDH
                ; Dai hoc preset
                GUICtrlSetData($inpFont, "Times New Roman")
                GUICtrlSetData($inpSize, "13")
                GUICtrlSetData($inpH1, "14")
                GUICtrlSetData($inpH2, "13")
                GUICtrlSetData($inpH3, "13")
                GUICtrlSetData($inpLineSpacing, "1.5")
                GUICtrlSetData($inpIndent, "1.27")
                GUICtrlSetData($inpLeft, "3")
                GUICtrlSetData($inpRight, "2")
                GUICtrlSetData($inpTop, "2")
                GUICtrlSetData($inpBottom, "2")
                GUICtrlSetData($inpBefore, "0")
                GUICtrlSetData($inpAfter, "6")
                _AI_UpdatePreviewLabel($lblPreview, $inpFont, $inpSize, $inpLineSpacing, $inpIndent, $inpLeft, $inpRight, $inpTop, $inpBottom)

            Case $btnPresetTHS
                ; Thac si preset
                GUICtrlSetData($inpFont, "Times New Roman")
                GUICtrlSetData($inpSize, "14")
                GUICtrlSetData($inpH1, "16")
                GUICtrlSetData($inpH2, "14")
                GUICtrlSetData($inpH3, "14")
                GUICtrlSetData($inpLineSpacing, "1.5")
                GUICtrlSetData($inpIndent, "1.27")
                GUICtrlSetData($inpLeft, "3.5")
                GUICtrlSetData($inpRight, "2")
                GUICtrlSetData($inpTop, "2.5")
                GUICtrlSetData($inpBottom, "2.5")
                GUICtrlSetData($inpBefore, "0")
                GUICtrlSetData($inpAfter, "6")
                _AI_UpdatePreviewLabel($lblPreview, $inpFont, $inpSize, $inpLineSpacing, $inpIndent, $inpLeft, $inpRight, $inpTop, $inpBottom)

            Case $btnPresetDefault
                ; Default
                GUICtrlSetData($inpFont, "Times New Roman")
                GUICtrlSetData($inpSize, "13")
                GUICtrlSetData($inpH1, "14")
                GUICtrlSetData($inpH2, "13")
                GUICtrlSetData($inpH3, "13")
                GUICtrlSetData($inpLineSpacing, "1.5")
                GUICtrlSetData($inpIndent, "1.27")
                GUICtrlSetData($inpLeft, "3")
                GUICtrlSetData($inpRight, "2")
                GUICtrlSetData($inpTop, "2")
                GUICtrlSetData($inpBottom, "2")
                GUICtrlSetData($inpBefore, "0")
                GUICtrlSetData($inpAfter, "6")
                _AI_UpdatePreviewLabel($lblPreview, $inpFont, $inpSize, $inpLineSpacing, $inpIndent, $inpLeft, $inpRight, $inpTop, $inpBottom)

            Case $btnSave
                ; Luu vao variables
                $AI_FONT_NAME = GUICtrlRead($inpFont)
                $AI_FONT_SIZE = Number(GUICtrlRead($inpSize))
                $AI_FONT_SIZE_H1 = Number(GUICtrlRead($inpH1))
                $AI_FONT_SIZE_H2 = Number(GUICtrlRead($inpH2))
                $AI_FONT_SIZE_H3 = Number(GUICtrlRead($inpH3))
                $AI_LINE_SPACING = Number(GUICtrlRead($inpLineSpacing))
                $AI_FIRST_INDENT = Number(GUICtrlRead($inpIndent))
                $AI_MARGIN_LEFT = Number(GUICtrlRead($inpLeft))
                $AI_MARGIN_RIGHT = Number(GUICtrlRead($inpRight))
                $AI_MARGIN_TOP = Number(GUICtrlRead($inpTop))
                $AI_MARGIN_BOTTOM = Number(GUICtrlRead($inpBottom))
                $AI_PARA_SPACING_BEFORE = Number(GUICtrlRead($inpBefore))
                $AI_PARA_SPACING_AFTER = Number(GUICtrlRead($inpAfter))

                ; Luu vao INI
                IniWrite($AI_SETTINGS_FILE, "Format", "FontName", $AI_FONT_NAME)
                IniWrite($AI_SETTINGS_FILE, "Format", "FontSize", $AI_FONT_SIZE)
                IniWrite($AI_SETTINGS_FILE, "Format", "H1Size", $AI_FONT_SIZE_H1)
                IniWrite($AI_SETTINGS_FILE, "Format", "H2Size", $AI_FONT_SIZE_H2)
                IniWrite($AI_SETTINGS_FILE, "Format", "H3Size", $AI_FONT_SIZE_H3)
                IniWrite($AI_SETTINGS_FILE, "Format", "LineSpacing", $AI_LINE_SPACING)
                IniWrite($AI_SETTINGS_FILE, "Format", "FirstIndent", $AI_FIRST_INDENT)
                IniWrite($AI_SETTINGS_FILE, "Format", "MarginLeft", $AI_MARGIN_LEFT)
                IniWrite($AI_SETTINGS_FILE, "Format", "MarginRight", $AI_MARGIN_RIGHT)
                IniWrite($AI_SETTINGS_FILE, "Format", "MarginTop", $AI_MARGIN_TOP)
                IniWrite($AI_SETTINGS_FILE, "Format", "MarginBottom", $AI_MARGIN_BOTTOM)
                IniWrite($AI_SETTINGS_FILE, "Format", "SpaceBefore", $AI_PARA_SPACING_BEFORE)
                IniWrite($AI_SETTINGS_FILE, "Format", "SpaceAfter", $AI_PARA_SPACING_AFTER)

                MsgBox($MB_ICONINFORMATION, "Luu thanh cong", _
                    "Da luu cai dat!" & @CRLF & @CRLF & _
                    "Font: " & $AI_FONT_NAME & " " & $AI_FONT_SIZE & "pt" & @CRLF & _
                    "Line spacing: " & $AI_LINE_SPACING & @CRLF & _
                    "Margins: " & $AI_MARGIN_LEFT & "/" & $AI_MARGIN_RIGHT & "/" & $AI_MARGIN_TOP & "/" & $AI_MARGIN_BOTTOM & " cm" & @CRLF & _
                    "Indent: " & $AI_FIRST_INDENT & "cm")
                ExitLoop
        EndSwitch
    WEnd

    GUIDelete($hSettings)
EndFunc

; Helper: Update preview label
Func _AI_UpdatePreviewLabel($lblPreview, $inpFont, $inpSize, $inpSpacing, $inpIndent, $inpL, $inpR, $inpT, $inpB)
    Local $sPreview = "Font: " & GUICtrlRead($inpFont) & " " & GUICtrlRead($inpSize) & "pt" & @CRLF
    $sPreview &= "Line spacing: " & GUICtrlRead($inpSpacing) & @CRLF
    $sPreview &= "Thut dau dong: " & GUICtrlRead($inpIndent) & "cm" & @CRLF
    $sPreview &= "Le: Trai=" & GUICtrlRead($inpL) & "cm, Phai=" & GUICtrlRead($inpR) & "cm" & @CRLF
    $sPreview &= "     Tren=" & GUICtrlRead($inpT) & "cm, Duoi=" & GUICtrlRead($inpB) & "cm"
    GUICtrlSetData($lblPreview, $sPreview)
EndFunc

; ========================================
; BEAUTIFY SETTINGS
; ========================================
; Cac thong so lam dep (co the tuy chinh)
Global $AI_H1_SPACE_BEFORE = 12
Global $AI_H1_SPACE_AFTER = 6
Global $AI_H2_SPACE_BEFORE = 6
Global $AI_H2_SPACE_AFTER = 3
Global $AI_H3_SPACE_BEFORE = 6
Global $AI_H3_SPACE_AFTER = 3
Global $AI_LIST_INDENT = 1.27 ; cm
Global $AI_LIST_SPACE_AFTER = 2 ; pt
Global $AI_PARA_ALIGNMENT = 3 ; 0=Left, 1=Center, 2=Right, 3=Justify
Global $AI_REMOVE_ORPHAN_BULLETS = True
Global $AI_REMOVE_EMPTY_DOUBLES = True
Global $AI_REMOVE_MD_SEPARATORS = True

_AI_LoadBeautifySettings()

Func _AI_LoadBeautifySettings()
    If Not FileExists($AI_SETTINGS_FILE) Then Return
    $AI_H1_SPACE_BEFORE = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "H1SpaceBefore", $AI_H1_SPACE_BEFORE))
    $AI_H1_SPACE_AFTER = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "H1SpaceAfter", $AI_H1_SPACE_AFTER))
    $AI_H2_SPACE_BEFORE = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "H2SpaceBefore", $AI_H2_SPACE_BEFORE))
    $AI_H2_SPACE_AFTER = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "H2SpaceAfter", $AI_H2_SPACE_AFTER))
    $AI_H3_SPACE_BEFORE = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "H3SpaceBefore", $AI_H3_SPACE_BEFORE))
    $AI_H3_SPACE_AFTER = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "H3SpaceAfter", $AI_H3_SPACE_AFTER))
    $AI_LIST_INDENT = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "ListIndent", $AI_LIST_INDENT))
    $AI_LIST_SPACE_AFTER = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "ListSpaceAfter", $AI_LIST_SPACE_AFTER))
    $AI_PARA_ALIGNMENT = Number(IniRead($AI_SETTINGS_FILE, "Beautify", "ParaAlignment", $AI_PARA_ALIGNMENT))
    $AI_REMOVE_ORPHAN_BULLETS = (IniRead($AI_SETTINGS_FILE, "Beautify", "RemoveOrphanBullets", "True") = "True")
    $AI_REMOVE_EMPTY_DOUBLES = (IniRead($AI_SETTINGS_FILE, "Beautify", "RemoveEmptyDoubles", "True") = "True")
    $AI_REMOVE_MD_SEPARATORS = (IniRead($AI_SETTINGS_FILE, "Beautify", "RemoveMDSeparators", "True") = "True")
EndFunc

Func _AI_ShowBeautifySettings()
    Local $hDlg = GUICreate("Thiet dat Lam dep van ban", 420, 460, -1, -1)
    GUISetBkColor(0xF0F4F8, $hDlg)

    ; Title
    GUICtrlCreateLabel("THIET DAT LAM DEP", 20, 10, 380, 25, $SS_CENTER)
    GUICtrlSetFont(-1, 14, 800, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x2C3E50)

    ; === Heading spacing ===
    GUICtrlCreateGroup(" Heading spacing (pt) ", 15, 40, 390, 90)
    GUICtrlCreateLabel("H1 - Before:", 30, 62, 70, 20)
    Local $inpH1B = GUICtrlCreateInput($AI_H1_SPACE_BEFORE, 105, 59, 40, 22)
    GUICtrlCreateLabel("After:", 155, 62, 35, 20)
    Local $inpH1A = GUICtrlCreateInput($AI_H1_SPACE_AFTER, 195, 59, 40, 22)
    GUICtrlCreateLabel("H2 - Before:", 30, 87, 70, 20)
    Local $inpH2B = GUICtrlCreateInput($AI_H2_SPACE_BEFORE, 105, 84, 40, 22)
    GUICtrlCreateLabel("After:", 155, 87, 35, 20)
    Local $inpH2A = GUICtrlCreateInput($AI_H2_SPACE_AFTER, 195, 84, 40, 22)
    GUICtrlCreateLabel("H3 - Before:", 30, 107, 70, 20)
    Local $inpH3B = GUICtrlCreateInput($AI_H3_SPACE_BEFORE, 105, 104, 40, 22)
    GUICtrlCreateLabel("After:", 155, 107, 35, 20)
    Local $inpH3A = GUICtrlCreateInput($AI_H3_SPACE_AFTER, 195, 104, 40, 22)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === List ===
    GUICtrlCreateGroup(" Bullet/Number list ", 15, 135, 390, 55)
    GUICtrlCreateLabel("Indent (cm):", 30, 158, 70, 20)
    Local $inpListIndent = GUICtrlCreateInput($AI_LIST_INDENT, 105, 155, 50, 22)
    GUICtrlCreateLabel("Space after (pt):", 175, 158, 90, 20)
    Local $inpListAfter = GUICtrlCreateInput($AI_LIST_SPACE_AFTER, 270, 155, 40, 22)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Alignment ===
    GUICtrlCreateGroup(" Text alignment ", 15, 195, 390, 50)
    Local $radLeft = GUICtrlCreateRadio("Left", 30, 215, 55, 20)
    Local $radCenter = GUICtrlCreateRadio("Center", 95, 215, 65, 20)
    Local $radRight = GUICtrlCreateRadio("Right", 170, 215, 55, 20)
    Local $radJustify = GUICtrlCreateRadio("Justify", 235, 215, 65, 20)
    ; Set checked
    Switch $AI_PARA_ALIGNMENT
        Case 0
            GUICtrlSetState($radLeft, $GUI_CHECKED)
        Case 1
            GUICtrlSetState($radCenter, $GUI_CHECKED)
        Case 2
            GUICtrlSetState($radRight, $GUI_CHECKED)
        Case 3
            GUICtrlSetState($radJustify, $GUI_CHECKED)
    EndSwitch
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Cleanup options ===
    GUICtrlCreateGroup(" Don dep tu dong ", 15, 250, 390, 80)
    Local $chkOrphan = GUICtrlCreateCheckbox("Xoa dong bullet rong (chi co bullet, khong co text)", 30, 270, 350, 20)
    If $AI_REMOVE_ORPHAN_BULLETS Then GUICtrlSetState($chkOrphan, $GUI_CHECKED)
    Local $chkEmpty = GUICtrlCreateCheckbox("Xoa dong trong lien tiep (giu lai 1)", 30, 290, 350, 20)
    If $AI_REMOVE_EMPTY_DOUBLES Then GUICtrlSetState($chkEmpty, $GUI_CHECKED)
    Local $chkSep = GUICtrlCreateCheckbox("Xoa markdown separator (---, ***, ===)", 30, 310, 350, 20)
    If $AI_REMOVE_MD_SEPARATORS Then GUICtrlSetState($chkSep, $GUI_CHECKED)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Preset ===
    GUICtrlCreateGroup(" Mau ", 15, 335, 390, 45)
    Local $btnPreDoan = GUICtrlCreateButton("Do an", 30, 352, 80, 22)
    Local $btnPreBaocao = GUICtrlCreateButton("Bao cao", 120, 352, 80, 22)
    Local $btnPreSach = GUICtrlCreateButton("Sach", 210, 352, 80, 22)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; === Buttons ===
    Local $btnSave = GUICtrlCreateButton("Luu", 60, 395, 130, 35)
    GUICtrlSetFont(-1, 11, 700, 0, "Segoe UI")
    Local $btnCancel = GUICtrlCreateButton("Huy", 230, 395, 130, 35)

    GUISetState(@SW_SHOW, $hDlg)

    While 1
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case -3
                ExitLoop
            Case $btnCancel
                ExitLoop

            Case $btnPreDoan
                GUICtrlSetData($inpH1B, "12")
                GUICtrlSetData($inpH1A, "6")
                GUICtrlSetData($inpH2B, "6")
                GUICtrlSetData($inpH2A, "3")
                GUICtrlSetData($inpH3B, "6")
                GUICtrlSetData($inpH3A, "3")
                GUICtrlSetData($inpListIndent, "1.27")
                GUICtrlSetData($inpListAfter, "2")
                GUICtrlSetState($radJustify, $GUI_CHECKED)

            Case $btnPreBaocao
                GUICtrlSetData($inpH1B, "18")
                GUICtrlSetData($inpH1A, "12")
                GUICtrlSetData($inpH2B, "12")
                GUICtrlSetData($inpH2A, "6")
                GUICtrlSetData($inpH3B, "6")
                GUICtrlSetData($inpH3A, "6")
                GUICtrlSetData($inpListIndent, "1.5")
                GUICtrlSetData($inpListAfter, "3")
                GUICtrlSetState($radJustify, $GUI_CHECKED)

            Case $btnPreSach
                GUICtrlSetData($inpH1B, "24")
                GUICtrlSetData($inpH1A, "12")
                GUICtrlSetData($inpH2B, "12")
                GUICtrlSetData($inpH2A, "6")
                GUICtrlSetData($inpH3B, "6")
                GUICtrlSetData($inpH3A, "3")
                GUICtrlSetData($inpListIndent, "1")
                GUICtrlSetData($inpListAfter, "2")
                GUICtrlSetState($radLeft, $GUI_CHECKED)

            Case $btnSave
                $AI_H1_SPACE_BEFORE = Number(GUICtrlRead($inpH1B))
                $AI_H1_SPACE_AFTER = Number(GUICtrlRead($inpH1A))
                $AI_H2_SPACE_BEFORE = Number(GUICtrlRead($inpH2B))
                $AI_H2_SPACE_AFTER = Number(GUICtrlRead($inpH2A))
                $AI_H3_SPACE_BEFORE = Number(GUICtrlRead($inpH3B))
                $AI_H3_SPACE_AFTER = Number(GUICtrlRead($inpH3A))
                $AI_LIST_INDENT = Number(GUICtrlRead($inpListIndent))
                $AI_LIST_SPACE_AFTER = Number(GUICtrlRead($inpListAfter))
                $AI_REMOVE_ORPHAN_BULLETS = (GUICtrlRead($chkOrphan) = $GUI_CHECKED)
                $AI_REMOVE_EMPTY_DOUBLES = (GUICtrlRead($chkEmpty) = $GUI_CHECKED)
                $AI_REMOVE_MD_SEPARATORS = (GUICtrlRead($chkSep) = $GUI_CHECKED)

                ; Alignment
                If GUICtrlRead($radLeft) = $GUI_CHECKED Then $AI_PARA_ALIGNMENT = 0
                If GUICtrlRead($radCenter) = $GUI_CHECKED Then $AI_PARA_ALIGNMENT = 1
                If GUICtrlRead($radRight) = $GUI_CHECKED Then $AI_PARA_ALIGNMENT = 2
                If GUICtrlRead($radJustify) = $GUI_CHECKED Then $AI_PARA_ALIGNMENT = 3

                ; Save to INI
                IniWrite($AI_SETTINGS_FILE, "Beautify", "H1SpaceBefore", $AI_H1_SPACE_BEFORE)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "H1SpaceAfter", $AI_H1_SPACE_AFTER)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "H2SpaceBefore", $AI_H2_SPACE_BEFORE)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "H2SpaceAfter", $AI_H2_SPACE_AFTER)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "H3SpaceBefore", $AI_H3_SPACE_BEFORE)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "H3SpaceAfter", $AI_H3_SPACE_AFTER)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "ListIndent", $AI_LIST_INDENT)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "ListSpaceAfter", $AI_LIST_SPACE_AFTER)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "ParaAlignment", $AI_PARA_ALIGNMENT)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "RemoveOrphanBullets", $AI_REMOVE_ORPHAN_BULLETS)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "RemoveEmptyDoubles", $AI_REMOVE_EMPTY_DOUBLES)
                IniWrite($AI_SETTINGS_FILE, "Beautify", "RemoveMDSeparators", $AI_REMOVE_MD_SEPARATORS)

                MsgBox($MB_ICONINFORMATION, "Luu thanh cong", "Da luu thiet dat lam dep!")
                ExitLoop
        EndSwitch
    WEnd

    GUIDelete($hDlg)
EndFunc

; ========================================
; HELPER: Lay Range xu ly (Selection hoac Content)
; ========================================
Func _AI_GetRange()
    If Not _CheckConnection() Then Return SetError(1, 0, 0)

    Local $bSelection = (GUICtrlRead($g_chkAIScopeSelection) = $GUI_CHECKED)
    Local $bAll = (GUICtrlRead($g_chkAIScopeAll) = $GUI_CHECKED)

    If $bSelection Then
        Local $oSel = $g_oWord.Selection
        If IsObj($oSel) And $oSel.Type <> 1 Then ; Khong phai insertion point
            Return $oSel.Range
        EndIf
    EndIf

    If $bAll Then
        Return $g_oDoc.Content
    EndIf

    ; Mac dinh: thu selection, neu khong co thi lay Content
    Local $oSel2 = $g_oWord.Selection
    If IsObj($oSel2) And $oSel2.Type <> 1 Then
        Return $oSel2.Range
    EndIf
    Return $g_oDoc.Content
EndFunc

Func _AI_ExecuteReplaceAll($oFind)
    If Not IsObj($oFind) Then Return False
    Return $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
EndFunc

Func _AI_ReplaceSmartQuotesInFind($oFind)
    If Not IsObj($oFind) Then Return

    Local $bPrevAutoType = False, $bPrevAutoReplace = False
    Local $bHasOptions = IsObj($g_oWord) And IsObj($g_oWord.Options)

    If $bHasOptions Then
        $bPrevAutoType = $g_oWord.Options.AutoFormatAsYouTypeReplaceQuotes
        $bPrevAutoReplace = $g_oWord.Options.AutoFormatReplaceQuotes
        $g_oWord.Options.AutoFormatAsYouTypeReplaceQuotes = False
        $g_oWord.Options.AutoFormatReplaceQuotes = False
    EndIf

    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    $oFind.MatchWildcards = False

    $oFind.Text = ChrW(8220)
    $oFind.Replacement.Text = '"'
    _AI_ExecuteReplaceAll($oFind)

    $oFind.Text = ChrW(8221)
    $oFind.Replacement.Text = '"'
    _AI_ExecuteReplaceAll($oFind)

    $oFind.Text = ChrW(8216)
    $oFind.Replacement.Text = "'"
    _AI_ExecuteReplaceAll($oFind)

    $oFind.Text = ChrW(8217)
    $oFind.Replacement.Text = "'"
    _AI_ExecuteReplaceAll($oFind)

    If $bHasOptions Then
        $g_oWord.Options.AutoFormatAsYouTypeReplaceQuotes = $bPrevAutoType
        $g_oWord.Options.AutoFormatReplaceQuotes = $bPrevAutoReplace
    EndIf
EndFunc

; ========================================
; MARKDOWN CLEANUP FUNCTIONS
; ========================================

; ## Heading -> Word Heading Style
Func _AI_ConvertHeadings()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang chuyen Markdown Heading...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $oFind, $iCount = 0

    ; H1: # text (chi 1 dau #)
    ; Tim dong bat dau bang "# " (nhung khong phai "## ")
    For $iLevel = 3 To 1 Step -1
        Local $sPrefix = ""
        For $j = 1 To $iLevel
            $sPrefix &= "#"
        Next
        $sPrefix &= " "

        Local $sStyleName = "Heading " & $iLevel

        ; Duyet tung paragraph
        Local $oParas = $oRange.Paragraphs
        For $i = 1 To $oParas.Count
            Local $oPara = $oParas.Item($i)
            If Not IsObj($oPara) Then ContinueLoop

            Local $sText = $oPara.Range.Text
            $sText = StringStripWS($sText, 1) ; Strip leading

            If StringLeft($sText, StringLen($sPrefix)) = $sPrefix Then
                ; Xoa ky tu # va khoang trang
                Local $oParaRange = $oPara.Range
                ; Tim va xoa prefix
                Local $oFindInPara = $oParaRange.Find
                $oFindInPara.Text = $sPrefix
                $oFindInPara.Replacement.Text = ""
                _AI_ExecuteReplaceAll($oFindInPara)

                ; Gan style
                $oPara.Style = $sStyleName
                $iCount += 1
            EndIf
        Next
    Next

    _UpdateProgress("Da chuyen " & $iCount & " Heading!")
    If $iCount = 0 Then
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong tim thay Markdown Heading (## text)")
    EndIf
EndFunc

; **text** -> Bold
; FIX: Dung paragraph-based approach thay vi Word Find wildcards (khong tin cay)
Func _AI_ConvertBold()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang chuyen **Bold**...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $iCount = 0
    Local $oParas = $oRange.Paragraphs

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        ; Xu ly ** markers trong paragraph
        Local $iSafe = 0
        While $iSafe < 20 ; Tranh vong lap vo han
            $iSafe += 1
            Local $sText = $oPara.Range.Text
            Local $iParaStart = $oPara.Range.Start

            ; Tim cap ** ... **
            Local $iOpen = StringInStr($sText, "**")
            If $iOpen = 0 Then ExitLoop

            Local $iClose = StringInStr($sText, "**", 0, 1, $iOpen + 2)
            If $iClose = 0 Then ExitLoop

            ; Tinh vi tri trong Word Range (0-based tu dau document)
            ; StringInStr tra ve 1-based, nen tru 1
            Local $iWordOpen = $iParaStart + $iOpen - 1
            Local $iWordClose = $iParaStart + $iClose - 1

            ; Bold phan text giua 2 cap **
            Local $oBoldRange = $g_oDoc.Range($iWordOpen + 2, $iWordClose)
            If IsObj($oBoldRange) Then
                $oBoldRange.Font.Bold = True
            EndIf

            ; Xoa ** dong truoc (de vi tri khong bi lech)
            Local $oCloseMarker = $g_oDoc.Range($iWordClose, $iWordClose + 2)
            If IsObj($oCloseMarker) Then $oCloseMarker.Delete()

            ; Xoa ** mo
            Local $oOpenMarker = $g_oDoc.Range($iWordOpen, $iWordOpen + 2)
            If IsObj($oOpenMarker) Then $oOpenMarker.Delete()

            $iCount += 1
        WEnd

        ; Xu ly __ markers (it gap hon)
        $iSafe = 0
        While $iSafe < 20
            $iSafe += 1
            Local $sText2 = $oPara.Range.Text
            Local $iParaStart2 = $oPara.Range.Start

            Local $iOpen2 = StringInStr($sText2, "__")
            If $iOpen2 = 0 Then ExitLoop

            Local $iClose2 = StringInStr($sText2, "__", 0, 1, $iOpen2 + 2)
            If $iClose2 = 0 Then ExitLoop

            Local $iWordOpen2 = $iParaStart2 + $iOpen2 - 1
            Local $iWordClose2 = $iParaStart2 + $iClose2 - 1

            Local $oBoldRange2 = $g_oDoc.Range($iWordOpen2 + 2, $iWordClose2)
            If IsObj($oBoldRange2) Then $oBoldRange2.Font.Bold = True

            Local $oClose2 = $g_oDoc.Range($iWordClose2, $iWordClose2 + 2)
            If IsObj($oClose2) Then $oClose2.Delete()
            Local $oOpen2 = $g_oDoc.Range($iWordOpen2, $iWordOpen2 + 2)
            If IsObj($oOpen2) Then $oOpen2.Delete()

            $iCount += 1
        WEnd
    Next

    _UpdateProgress("Da chuyen " & $iCount & " bold!")
EndFunc

; *text* -> Italic
Func _AI_ConvertItalic()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang chuyen *Italic*...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $iCount = 0
    Local $oParas = $oRange.Paragraphs

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $iSafe = 0
        While $iSafe < 30
            $iSafe += 1
            Local $sText = $oPara.Range.Text
            Local $iParaStart = $oPara.Range.Start
            Local $iMarkerLen = 0
            Local $iOpen = 0, $iClose = 0

            ; Uu tien xu ly _text_ truoc, sau do moi den *text*.
            ; Dieu nay tranh viec cat nham dau * trong **bold**.
            $iOpen = StringInStr($sText, "_")
            If $iOpen > 0 Then
                $iClose = StringInStr($sText, "_", 0, 1, $iOpen + 1)
                If $iClose > ($iOpen + 1) Then
                    $iMarkerLen = 1
                Else
                    $iOpen = 0
                    $iClose = 0
                EndIf
            EndIf

            If $iOpen = 0 Then
                $iOpen = _AI_FindSingleAsteriskOpen($sText)
                If $iOpen > 0 Then
                    $iClose = _AI_FindSingleAsteriskClose($sText, $iOpen + 1)
                    If $iClose > ($iOpen + 1) Then
                        $iMarkerLen = 1
                    Else
                        $iOpen = 0
                        $iClose = 0
                    EndIf
                EndIf
            EndIf

            If $iOpen = 0 Or $iClose = 0 Then ExitLoop

            Local $iWordOpen = $iParaStart + $iOpen - 1
            Local $iWordClose = $iParaStart + $iClose - 1
            Local $oItalicRange = $g_oDoc.Range($iWordOpen + $iMarkerLen, $iWordClose)
            If IsObj($oItalicRange) Then $oItalicRange.Font.Italic = True

            Local $oCloseMarker = $g_oDoc.Range($iWordClose, $iWordClose + $iMarkerLen)
            If IsObj($oCloseMarker) Then $oCloseMarker.Delete()
            Local $oOpenMarker = $g_oDoc.Range($iWordOpen, $iWordOpen + $iMarkerLen)
            If IsObj($oOpenMarker) Then $oOpenMarker.Delete()

            $iCount += 1
        WEnd
    Next

    _UpdateProgress("Da chuyen " & $iCount & " italic!")
EndFunc

Func _AI_FindSingleAsteriskOpen($sText)
    Local $iLen = StringLen($sText)
    For $i = 1 To $iLen
        If StringMid($sText, $i, 1) <> "*" Then ContinueLoop
        If $i < $iLen And StringMid($sText, $i + 1, 1) = "*" Then
            $i += 1
            ContinueLoop
        EndIf

        Local $sPrev = ""
        If $i > 1 Then $sPrev = StringMid($sText, $i - 1, 1)
        Local $sNext = ""
        If $i < $iLen Then $sNext = StringMid($sText, $i + 1, 1)

        If $sNext = "" Or StringIsSpace($sNext) Then ContinueLoop
        If $sPrev <> "" And StringRegExp($sPrev, "[A-Za-z0-9_]") Then ContinueLoop
        Return $i
    Next
    Return 0
EndFunc

Func _AI_FindSingleAsteriskClose($sText, $iStartPos)
    Local $iLen = StringLen($sText)
    For $i = $iStartPos To $iLen
        If StringMid($sText, $i, 1) <> "*" Then ContinueLoop
        If $i < $iLen And StringMid($sText, $i + 1, 1) = "*" Then
            $i += 1
            ContinueLoop
        EndIf

        Local $sPrev = ""
        If $i > 1 Then $sPrev = StringMid($sText, $i - 1, 1)
        Local $sNext = ""
        If $i < $iLen Then $sNext = StringMid($sText, $i + 1, 1)

        If $sPrev = "" Or StringIsSpace($sPrev) Then ContinueLoop
        If $sNext <> "" And StringRegExp($sNext, "[A-Za-z0-9_]") Then ContinueLoop
        Return $i
    Next
    Return 0
EndFunc

; ```code``` -> Code block (Consolas + border)
Func _AI_ConvertCodeBlocks()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang xu ly Code blocks...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $iCount = 0
    Local $bInCodeBlock = False
    Local $iStartPara = 0

    ; Duyet tung paragraph tim ``` markers
    Local $oParas = $oRange.Paragraphs
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sText = StringStripWS($oPara.Range.Text, 3)

        If StringLeft($sText, 3) = "```" Then
            If Not $bInCodeBlock Then
                ; Bat dau code block
                $bInCodeBlock = True
                $iStartPara = $i
                ; Xoa dong ``` mo
                $oPara.Range.Text = ""
            Else
                ; Ket thuc code block
                $bInCodeBlock = False
                ; Xoa dong ``` dong
                $oPara.Range.Text = ""

                ; Format cac dong giua thanh code
                For $j = $iStartPara To $i - 1
                    If $j > $oParas.Count Then ExitLoop
                    Local $oCodePara = $oParas.Item($j)
                    If IsObj($oCodePara) And StringLen(StringStripWS($oCodePara.Range.Text, 3)) > 0 Then
                        $oCodePara.Range.Font.Name = "Consolas"
                        $oCodePara.Range.Font.Size = 10
                        $oCodePara.Range.Shading.BackgroundPatternColor = 0xF5F5F5 ; Light gray
                        $oCodePara.Format.SpaceBefore = 0
                        $oCodePara.Format.SpaceAfter = 0
                        $oCodePara.Format.LineSpacingRule = 0 ; wdLineSpaceSingle
                    EndIf
                Next
                $iCount += 1
            EndIf
        EndIf
    Next

    _UpdateProgress("Da xu ly " & $iCount & " code blocks!")
EndFunc

; `code` -> Inline code (Consolas)
Func _AI_ConvertInlineCode()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang chuyen `inline code`...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $oFind = $oRange.Find
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    $oFind.Text = "`(*{1,})`"
    $oFind.Replacement.Text = "\1"
    $oFind.Replacement.Font.Name = "Consolas"
    $oFind.Replacement.Font.Size = 11
    $oFind.MatchWildcards = True
    _AI_ExecuteReplaceAll($oFind)

    _UpdateProgress("Da chuyen inline code!")
EndFunc

; - item / • item -> Word bullet list
; FIX: Them xu ly ky tu • (Unicode 2022) tu ChatGPT paste
Func _AI_ConvertBullets()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang chuyen Bullet lists...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $iCount = 0
    Local $oParas = $oRange.Paragraphs
    Local $aRuns[1][2]
    Local $iRunCount = 0
    Local $iRunStart = 0, $iRunEnd = 0
    ; Bullet characters: • (2022), ◦ (25E6), ▪ (25AA), ▸ (25B8), ● (25CF)
    Local $sBulletChars = ChrW(8226) & ChrW(9702) & ChrW(9642) & ChrW(9656) & ChrW(9679)

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sText = $oPara.Range.Text
        Local $sStripped = StringStripWS($sText, 1) ; Strip leading whitespace

        ; Kiem tra markdown bullet: - hoac *
        If StringRegExp($sStripped, "^[-\*]\s") Then
            Local $sClean = StringRegExpReplace($sText, "^[\s]*[-\*]\s+", "")
            If StringRight($sText, 1) = Chr(13) And StringRight($sClean, 1) <> Chr(13) Then $sClean &= Chr(13)
            $oPara.Range.Text = $sClean
            _AI_AddListRun($aRuns, $iRunCount, $iRunStart, $iRunEnd, $i)
            $iCount += 1
            ContinueLoop
        EndIf

        ; Kiem tra Unicode bullet characters: •, ◦, ▪ v.v.
        Local $sFirstChar = StringLeft($sStripped, 1)
        If StringInStr($sBulletChars, $sFirstChar) Then
            ; Xoa bullet char va khoang trang sau no
            Local $sClean2 = StringStripWS(StringTrimLeft($sStripped, 1), 1)
            ; Xoa ky tu CR cuoi
            $sClean2 = StringReplace($sClean2, Chr(13), "")
            $sClean2 = StringReplace($sClean2, Chr(10), "")
            If StringLen($sClean2) > 0 Then
                If StringRight($sText, 1) = Chr(13) And StringRight($sClean2, 1) <> Chr(13) Then $sClean2 &= Chr(13)
                $oPara.Range.Text = $sClean2
                _AI_AddListRun($aRuns, $iRunCount, $iRunStart, $iRunEnd, $i)
                $iCount += 1
            Else
                ; Dong chi co bullet char, khong co text -> XOA
                $oPara.Range.Delete()
            EndIf
            ContinueLoop
        EndIf

        ; Kiem tra tab + text (ChatGPT thuong chen tab truoc bullet)
        If StringLeft($sText, 1) = @TAB Then
            Local $sAfterTab = StringStripWS(StringTrimLeft($sText, 1), 1)
            Local $sFirstAfterTab = StringLeft($sAfterTab, 1)
            If StringInStr($sBulletChars, $sFirstAfterTab) Or StringRegExp($sAfterTab, "^[-\*]\s") Then
                ; Xoa tab + bullet
                Local $sClean3 = StringRegExpReplace($sAfterTab, "^[" & $sBulletChars & "\-\*]\s*", "")
                If StringRight($sText, 1) = Chr(13) And StringRight($sClean3, 1) <> Chr(13) Then $sClean3 &= Chr(13)
                $oPara.Range.Text = $sClean3
                _AI_AddListRun($aRuns, $iRunCount, $iRunStart, $iRunEnd, $i)
                $iCount += 1
            EndIf
        EndIf
    Next

    _AI_FlushListRun($aRuns, $iRunCount, $iRunStart, $iRunEnd)
    For $r = 1 To $iRunCount
        _AI_ApplyBulletToParagraphRange($oParas, $aRuns[$r][0], $aRuns[$r][1])
    Next

    _UpdateProgress("Da chuyen " & $iCount & " bullets!")
EndFunc

; 1. item -> Word numbered list
Func _AI_ConvertNumberedLists()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang chuyen Numbered lists...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $iCount = 0
    Local $oParas = $oRange.Paragraphs
    Local $aRuns[1][2]
    Local $iRunCount = 0
    Local $iRunStart = 0, $iRunEnd = 0

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sText = $oPara.Range.Text
        ; Tim dong bat dau bang "1. " hoac "2. " v.v.
        If StringRegExp($sText, "^[\s]*\d+[\.\)]\s") Then
            ; Xoa so thu tu dau dong
            Local $oParaRange = $oPara.Range
            Local $sClean = StringRegExpReplace($sText, "^[\s]*\d+[\.\)]\s+", "")
            If StringRight($sText, 1) = Chr(13) And StringRight($sClean, 1) <> Chr(13) Then $sClean &= Chr(13)
            $oParaRange.Text = $sClean

            ; Ap dung numbered list
            _AI_AddListRun($aRuns, $iRunCount, $iRunStart, $iRunEnd, $i)
            $iCount += 1
        EndIf
    Next

    _AI_FlushListRun($aRuns, $iRunCount, $iRunStart, $iRunEnd)
    For $r = 1 To $iRunCount
        _AI_ApplyNumberingToParagraphRange($oParas, $aRuns[$r][0], $aRuns[$r][1])
    Next

    _UpdateProgress("Da chuyen " & $iCount & " numbered items!")
EndFunc

; | col | col | -> Word table
Func _AI_ConvertTables()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang chuyen Markdown tables...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $oParas = $oRange.Paragraphs
    Local $aParas[1][3]
    Local $iParaCount = 0
    Local $iTablesConverted = 0

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        $iParaCount += 1
        ReDim $aParas[$iParaCount + 1][3]
        $aParas[$iParaCount][0] = $oPara.Range.Start
        $aParas[$iParaCount][1] = $oPara.Range.End
        $aParas[$iParaCount][2] = StringStripWS($oPara.Range.Text, 3)
    Next

    Local $i = $iParaCount
    While $i >= 1
        Local $sText = $aParas[$i][2]
        If StringLeft($sText, 1) = "|" And StringRight($sText, 1) = "|" Then
            Local $iBlockEnd = $i
            Local $iBlockStart = $i

            While $iBlockStart > 1
                Local $sPrev = $aParas[$iBlockStart - 1][2]
                If StringLeft($sPrev, 1) = "|" And StringRight($sPrev, 1) = "|" Then
                    $iBlockStart -= 1
                Else
                    ExitLoop
                EndIf
            WEnd

            Local $aTableRows[1] = [0]
            For $r = $iBlockStart To $iBlockEnd
                Local $sRow = $aParas[$r][2]
                If StringRegExp($sRow, "^\|[\s\-\|:]+\|$") Then ContinueLoop
                $aTableRows[0] += 1
                ReDim $aTableRows[$aTableRows[0] + 1]
                $aTableRows[$aTableRows[0]] = $sRow
            Next

            If $aTableRows[0] > 0 Then
                Local $iInsertStart = $aParas[$iBlockStart][0]
                Local $iDeleteEnd = $aParas[$iBlockEnd][1]
                Local $oDeleteRange = $g_oDoc.Range($iInsertStart, $iDeleteEnd)
                If IsObj($oDeleteRange) Then $oDeleteRange.Delete()
                _AI_CreateWordTableAtPosition($aTableRows, $iInsertStart)
                $iTablesConverted += 1
            EndIf

            $i = $iBlockStart - 1
            ContinueLoop
        EndIf
        $i -= 1
    WEnd

    _UpdateProgress("Da chuyen " & $iTablesConverted & " tables!")
    If $iTablesConverted = 0 Then
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong tim thay Markdown table (| col | col |)")
    EndIf
EndFunc

; Helper: Tao Word table tu array
Func _AI_CreateWordTable($aRows, $oInsertPara)
    If $aRows[0] < 1 Or Not IsObj($oInsertPara) Then Return

    ; Dem so cot tu dong dau
    Local $aCols = StringSplit(StringMid($aRows[1], 2, StringLen($aRows[1]) - 2), "|")
    Local $iCols = $aCols[0]
    Local $iRows = $aRows[0]

    ; Chen table tai vi tri paragraph
    Local $oTableRange = $oInsertPara.Range
    $oTableRange.Collapse($WD_COLLAPSE_START)

    _AI_CreateWordTableAtRange($aRows, $oTableRange)
EndFunc

Func _AI_CreateWordTableAtPosition($aRows, $iInsertStart)
    If $aRows[0] < 1 Then Return
    Local $oTableRange = $g_oDoc.Range($iInsertStart, $iInsertStart)
    If Not IsObj($oTableRange) Then Return
    _AI_CreateWordTableAtRange($aRows, $oTableRange)
EndFunc

Func _AI_CreateWordTableAtRange($aRows, $oTableRange)
    If $aRows[0] < 1 Or Not IsObj($oTableRange) Then Return

    Local $aCols = StringSplit(StringMid($aRows[1], 2, StringLen($aRows[1]) - 2), "|")
    Local $iCols = $aCols[0]
    Local $iRows = $aRows[0]

    Local $oTable = $g_oDoc.Tables.Add($oTableRange, $iRows, $iCols)
    If Not IsObj($oTable) Then Return

    ; Dien du lieu
    For $r = 1 To $iRows
        Local $aCells = StringSplit(StringMid($aRows[$r], 2, StringLen($aRows[$r]) - 2), "|")
        For $c = 1 To $iCols
            If $c <= $aCells[0] Then
                $oTable.Cell($r, $c).Range.Text = StringStripWS($aCells[$c], 3)
            EndIf
        Next
    Next

    ; Format table
    $oTable.AutoFitBehavior(1) ; wdAutoFitContent
    $oTable.Borders.Enable = True
    $oTable.Range.Font.Name = $AI_FONT_NAME
    $oTable.Range.Font.Size = $AI_FONT_SIZE

    ; Bold header row
    If $iRows > 0 Then
        $oTable.Rows.Item(1).Range.Font.Bold = True
    EndIf
EndFunc

Func _AI_ApplyBulletToParagraph($oPara)
    If Not IsObj($oPara) Then Return
    Local $oTemplate = $g_oWord.ListGalleries.Item(1).ListTemplates.Item(1)
    If IsObj($oTemplate) Then $oPara.Range.ListFormat.ApplyListTemplate($oTemplate)
EndFunc

Func _AI_ApplyNumberingToParagraph($oPara)
    If Not IsObj($oPara) Then Return
    Local $oTemplate = $g_oWord.ListGalleries.Item(2).ListTemplates.Item(1)
    If IsObj($oTemplate) Then $oPara.Range.ListFormat.ApplyListTemplate($oTemplate)
EndFunc

Func _AI_ApplyBulletToParagraphRange(ByRef $oParas, $iStart, $iEnd)
    If Not IsObj($oParas) Then Return
    Local $oStartPara = $oParas.Item($iStart)
    Local $oEndPara = $oParas.Item($iEnd)
    If Not IsObj($oStartPara) Or Not IsObj($oEndPara) Then Return
    Local $oRange = $g_oDoc.Range($oStartPara.Range.Start, $oEndPara.Range.End)
    If Not IsObj($oRange) Then Return
    Local $oTemplate = $g_oWord.ListGalleries.Item(1).ListTemplates.Item(1)
    If IsObj($oTemplate) Then $oRange.ListFormat.ApplyListTemplate($oTemplate)
EndFunc

Func _AI_ApplyNumberingToParagraphRange(ByRef $oParas, $iStart, $iEnd)
    If Not IsObj($oParas) Then Return
    Local $oStartPara = $oParas.Item($iStart)
    Local $oEndPara = $oParas.Item($iEnd)
    If Not IsObj($oStartPara) Or Not IsObj($oEndPara) Then Return
    Local $oRange = $g_oDoc.Range($oStartPara.Range.Start, $oEndPara.Range.End)
    If Not IsObj($oRange) Then Return
    Local $oTemplate = $g_oWord.ListGalleries.Item(2).ListTemplates.Item(1)
    If IsObj($oTemplate) Then $oRange.ListFormat.ApplyListTemplate($oTemplate)
EndFunc

Func _AI_AddListRun(ByRef $aRuns, ByRef $iRunCount, ByRef $iRunStart, ByRef $iRunEnd, $iParagraph)
    If $iRunStart = 0 Then
        $iRunStart = $iParagraph
        $iRunEnd = $iParagraph
        Return
    EndIf

    If $iParagraph = $iRunEnd + 1 Then
        $iRunEnd = $iParagraph
        Return
    EndIf

    _AI_FlushListRun($aRuns, $iRunCount, $iRunStart, $iRunEnd)
    $iRunStart = $iParagraph
    $iRunEnd = $iParagraph
EndFunc

Func _AI_FlushListRun(ByRef $aRuns, ByRef $iRunCount, ByRef $iRunStart, ByRef $iRunEnd)
    If $iRunStart = 0 Then Return
    $iRunCount += 1
    ReDim $aRuns[$iRunCount + 1][2]
    $aRuns[$iRunCount][0] = $iRunStart
    $aRuns[$iRunCount][1] = $iRunEnd
    $iRunStart = 0
    $iRunEnd = 0
EndFunc

; [text](url) -> text only
Func _AI_ConvertLinks()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang xu ly Markdown links...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $oParas = $oRange.Paragraphs
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sText = $oPara.Range.Text
        Local $sNew = StringRegExpReplace($sText, "!\[([^\]]+)\]\([^)]+\)", "\1")
        $sNew = StringRegExpReplace($sNew, "\[([^\]]+)\]\([^)]+\)", "\1")

        If $sNew <> $sText Then
            $oPara.Range.Text = $sNew
        EndIf
    Next

    _UpdateProgress("Da xu ly links!")
EndFunc

; Xoa tat ca Markdown
Func _AI_CleanAllMarkdown()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", _
        "Se xu ly TAT CA markdown:" & @CRLF & @CRLF & _
        "1. ## -> Heading styles" & @CRLF & _
        "2. **text** -> Bold" & @CRLF & _
        "3. *text* -> Italic" & @CRLF & _
        "4. ```code``` -> Code blocks" & @CRLF & _
        "5. `code` -> Inline code" & @CRLF & _
        "6. - items -> Bullet list" & @CRLF & _
        "7. 1. items -> Numbered list" & @CRLF & _
        "8. [text](url) -> text" & @CRLF & @CRLF & _
        "LUU Y: Nen Backup truoc!") <> $IDYES Then Return

    _UpdateProgress("Dang xu ly tat ca Markdown...")

    ; Thu tu quan trong: code blocks truoc, roi headings, roi inline
    _AI_ConvertCodeBlocks()
    _AI_ConvertHeadings()
    _AI_ConvertBold()
    _AI_ConvertItalic()
    _AI_ConvertInlineCode()
    _AI_ConvertLinks()
    _AI_ConvertBullets()
    _AI_ConvertNumberedLists()

    _UpdateProgress("Da xu ly tat ca Markdown!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da xoa va chuyen doi tat ca Markdown!")
EndFunc

; ========================================
; CHUAN HOA DO AN FUNCTIONS
; ========================================

; Font Times New Roman 13pt
Func _AI_ApplyThesisFont()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang ap dung font chuan do an...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    $oRange.Font.Name = $AI_FONT_NAME
    $oRange.Font.Size = $AI_FONT_SIZE
    $oRange.Font.Color = 0 ; Black

    _UpdateProgress("Da ap dung " & $AI_FONT_NAME & " " & $AI_FONT_SIZE & "pt!")
EndFunc

; Line spacing 1.5
Func _AI_ApplyThesisSpacing()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang ap dung line spacing...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $oParas = $oRange.Paragraphs
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        $oPara.Format.LineSpacingRule = $WD_LINE_SPACE_MULTIPLE
        $oPara.Format.LineSpacing = $AI_LINE_SPACING * 12 ; 1.5 * 12pt
        $oPara.Format.SpaceBefore = $AI_PARA_SPACING_BEFORE
        $oPara.Format.SpaceAfter = $AI_PARA_SPACING_AFTER
    Next

    _UpdateProgress("Da ap dung line spacing 1.5!")
EndFunc

; Margins chuan
Func _AI_ApplyThesisMargins()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang ap dung margins chuan do an...")

    $g_oDoc.PageSetup.LeftMargin = $AI_MARGIN_LEFT * $CM_TO_POINTS
    $g_oDoc.PageSetup.RightMargin = $AI_MARGIN_RIGHT * $CM_TO_POINTS
    $g_oDoc.PageSetup.TopMargin = $AI_MARGIN_TOP * $CM_TO_POINTS
    $g_oDoc.PageSetup.BottomMargin = $AI_MARGIN_BOTTOM * $CM_TO_POINTS

    _UpdateProgress("Da ap dung margins: L=" & $AI_MARGIN_LEFT & "cm, R=" & $AI_MARGIN_RIGHT & "cm, T=" & $AI_MARGIN_TOP & "cm, B=" & $AI_MARGIN_BOTTOM & "cm")
EndFunc

; Thut dau dong 1.27cm
Func _AI_ApplyFirstLineIndent()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang thut dau dong...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $oParas = $oRange.Paragraphs
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        ; Chi thut dau dong cho style Normal (khong thut Heading, Caption...)
        Local $oStyle = $oPara.Style
        If IsObj($oStyle) Then
            Local $sStyle = $oStyle.NameLocal
            If StringInStr($sStyle, "Heading") Or StringInStr($sStyle, "Tieu de") Or _
               StringInStr($sStyle, "Caption") Or StringInStr($sStyle, "TOC") Then
                ContinueLoop
            EndIf
        EndIf

        $oPara.Format.FirstLineIndent = $AI_FIRST_INDENT * $CM_TO_POINTS
    Next

    _UpdateProgress("Da thut dau dong " & $AI_FIRST_INDENT & "cm!")
EndFunc

; Fix khoang trang thua
; FIX: Them xu ly tab, khoang trang nhieu hon
Func _AI_FixExtraSpaces()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang fix khoang trang...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    ; 0. Chuyen TAB -> space (ChatGPT hay chen tab)
    Local $oFindTab = _AI_GetRange().Find
    $oFindTab.ClearFormatting()
    $oFindTab.Replacement.ClearFormatting()
    $oFindTab.Text = "^t"
    $oFindTab.Replacement.Text = " "
    $oFindTab.MatchWildcards = False
    _AI_ExecuteReplaceAll($oFindTab)

    ; 1. Xoa nhieu khoang trang lien tiep -> 1 khoang trang
    Local $iMax = 20
    While $iMax > 0
        Local $oFind = _AI_GetRange().Find
        $oFind.ClearFormatting()
        $oFind.Replacement.ClearFormatting()
        $oFind.Text = "  "
        $oFind.Replacement.Text = " "
        $oFind.MatchWildcards = False
        If Not _AI_ExecuteReplaceAll($oFind) Then ExitLoop
        $iMax -= 1
    WEnd

    ; 2. Xoa khoang trang dau dong va cuoi dong
    Local $oRange2 = _AI_GetRange()
    Local $oParas = $oRange2.Paragraphs
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        ; Bo qua list items (da co indent rieng)
        If $oPara.Range.ListFormat.ListType <> 0 Then ContinueLoop
        Local $sText = $oPara.Range.Text
        Local $sClean = StringStripWS($sText, 3) ; Strip leading AND trailing
        ; Giu lai paragraph mark
        If StringRight($sText, 1) = Chr(13) And StringRight($sClean, 1) <> Chr(13) Then
            $sClean = $sClean & Chr(13)
        EndIf
        If $sText <> $sClean And StringLen(StringReplace($sClean, Chr(13), "")) > 0 Then
            $oPara.Range.Text = $sClean
        EndIf
    Next

    _UpdateProgress("Da fix khoang trang!")
EndFunc

; Fix dau cau tieng Viet
Func _AI_FixVietnamesePunctuation()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang fix dau cau VN...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    ; 1. Khong co khoang trang truoc dau cham, phay, hai cham
    Local $aPatterns[4][2] = [ _
        [" .", "."], _
        [" ,", ","], _
        [" :", ":"], _
        [" ;", ";"] _
    ]

    For $p = 0 To UBound($aPatterns) - 1
        Local $oFind = _AI_GetRange().Find
        $oFind.ClearFormatting()
        $oFind.Replacement.ClearFormatting()
        $oFind.Text = $aPatterns[$p][0]
        $oFind.Replacement.Text = $aPatterns[$p][1]
        _AI_ExecuteReplaceAll($oFind)
    Next

    _UpdateProgress("Da fix dau cau VN!")
EndFunc

; Xoa dong trong thua
; FIX: Manh tay hon - ^p^p -> ^p (khong giu dong trong)
Func _AI_RemoveEmptyLines()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang xoa dong trong...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    ; Buoc 1: Giam 3+ dong trong -> 1 dong trong
    Local $iMax = 20
    While $iMax > 0
        Local $oFind = _AI_GetRange().Find
        $oFind.ClearFormatting()
        $oFind.Replacement.ClearFormatting()
        $oFind.Text = "^p^p^p"
        $oFind.Replacement.Text = "^p"
        $oFind.MatchWildcards = False
        If Not _AI_ExecuteReplaceAll($oFind) Then ExitLoop
        $iMax -= 1
    WEnd

    ; Buoc 2: Giam 2 dong trong -> 1 (loai bo dong trong thua giua cac paragraph)
    $iMax = 20
    While $iMax > 0
        Local $oFind2 = _AI_GetRange().Find
        $oFind2.ClearFormatting()
        $oFind2.Replacement.ClearFormatting()
        $oFind2.Text = "^p^p"
        $oFind2.Replacement.Text = "^p"
        $oFind2.MatchWildcards = False
        If Not _AI_ExecuteReplaceAll($oFind2) Then ExitLoop
        $iMax -= 1
    WEnd

    _UpdateProgress("Da xoa dong trong!")
EndFunc

; FIX TAT CA - Chuan do an
; ENHANCED: Them buoc 0 (cleanup) va buoc 4 (beautify)
Func _AI_FixAllThesis()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Chuan hoa do an", _
        "Se ap dung TOAN BO chuan do an dai hoc VN:" & @CRLF & @CRLF & _
        "0. Don dep ky tu rac (bullet rong, markdown sot)" & @CRLF & _
        "1. Xu ly tat ca Markdown" & @CRLF & _
        "2. Font: " & $AI_FONT_NAME & " " & $AI_FONT_SIZE & "pt" & @CRLF & _
        "3. Line spacing: " & $AI_LINE_SPACING & @CRLF & _
        "4. Margins: L=" & $AI_MARGIN_LEFT & "cm, R=" & $AI_MARGIN_RIGHT & "cm" & @CRLF & _
        "5. Thut dau dong: " & $AI_FIRST_INDENT & "cm" & @CRLF & _
        "6. Fix khoang trang, dau cau, dong trong" & @CRLF & _
        "7. Lam dep van ban" & @CRLF & @CRLF & _
        "LUU Y: Nen Backup truoc! Tiep tuc?") <> $IDYES Then Return

    _UpdateProgress("DANG CHUAN HOA DO AN...")

    ; Buoc 0: Don dep truoc
    _UpdateProgress("[0/7] Don dep ky tu rac...")
    _AI_CleanupVisual()

    ; Buoc 1: Xu ly Markdown
    _UpdateProgress("[1/7] Xu ly Markdown...")
    _AI_ConvertCodeBlocks()
    _AI_ConvertHeadings()
    _AI_ConvertBold()
    _AI_ConvertItalic()
    _AI_ConvertInlineCode()
    _AI_ConvertLinks()
    _AI_ConvertBullets()
    _AI_ConvertNumberedLists()

    ; Buoc 2: Chuan hoa format
    _UpdateProgress("[2/7] Ap dung font...")
    _AI_ApplyThesisFont()

    _UpdateProgress("[3/7] Ap dung spacing...")
    _AI_ApplyThesisSpacing()

    _UpdateProgress("[4/7] Ap dung margins...")
    _AI_ApplyThesisMargins()

    _UpdateProgress("[5/7] Thut dau dong...")
    _AI_ApplyFirstLineIndent()

    ; Buoc 3: Fix loi
    _UpdateProgress("[6/7] Fix khoang trang, dau cau...")
    _AI_FixExtraSpaces()
    _AI_FixVietnamesePunctuation()
    _AI_RemoveEmptyLines()

    ; Buoc 4: Lam dep
    _UpdateProgress("[7/7] Lam dep van ban...")
    _AI_BeautifyDocument()

    _UpdateProgress("DA HOAN TAT CHUAN HOA DO AN!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", _
        "Da chuan hoa do an thanh cong!" & @CRLF & @CRLF & _
        "Kiem tra lai:" & @CRLF & _
        "- Heading dung style?" & @CRLF & _
        "- Code block dung font?" & @CRLF & _
        "- Margins, line spacing dung?" & @CRLF & _
        "- Bullet list dep?" & @CRLF & _
        "- Dau cau hop ly?")
EndFunc

; ========================================
; CLEANUP & BEAUTIFY FUNCTIONS
; ========================================

; Don dep ky tu rac TRUOC khi xu ly markdown
Func _AI_CleanupVisual()
    If Not _CheckConnection() Then Return

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $oParas = $oRange.Paragraphs
    Local $sBulletChars = ChrW(8226) & ChrW(9702) & ChrW(9642) & ChrW(9656) & ChrW(9679)
    Local $iCleaned = 0

    ; Duyet nguoc de xoa paragraph khong bi lech index
    For $i = $oParas.Count To 1 Step -1
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sText = $oPara.Range.Text
        Local $sStripped = StringStripWS($sText, 3)
        $sStripped = StringReplace($sStripped, Chr(13), "")
        $sStripped = StringReplace($sStripped, Chr(10), "")

        ; 1. Xoa dong chi co bullet char (theo settings)
        If $AI_REMOVE_ORPHAN_BULLETS And StringLen($sStripped) <= 2 Then
            Local $bOnlyBullet = True
            For $c = 1 To StringLen($sStripped)
                Local $ch = StringMid($sStripped, $c, 1)
                If Not StringInStr($sBulletChars & "-*" & @TAB & " ", $ch) Then
                    $bOnlyBullet = False
                    ExitLoop
                EndIf
            Next
            If $bOnlyBullet And StringLen($sStripped) > 0 Then
                $oPara.Range.Delete()
                $iCleaned += 1
                ContinueLoop
            EndIf
        EndIf

        ; 2. Xoa dong trong lien tiep (theo settings)
        If $AI_REMOVE_EMPTY_DOUBLES And StringLen($sStripped) = 0 Then
            If $i > 1 Then
                Local $oPrevPara = $oParas.Item($i - 1)
                If IsObj($oPrevPara) Then
                    Local $sPrevText = StringStripWS($oPrevPara.Range.Text, 3)
                    $sPrevText = StringReplace($sPrevText, Chr(13), "")
                    If StringLen($sPrevText) = 0 Then
                        $oPara.Range.Delete()
                        $iCleaned += 1
                        ContinueLoop
                    EndIf
                EndIf
            EndIf
        EndIf

        ; 3. Xoa markdown separators (theo settings)
        If $AI_REMOVE_MD_SEPARATORS And StringRegExp($sStripped, "^[-=\*]{3,}$") Then
            $oPara.Range.Delete()
            $iCleaned += 1
            ContinueLoop
        EndIf
    Next

    If $iCleaned > 0 Then
        _UpdateProgress("Da don dep " & $iCleaned & " dong rac!")
    EndIf
EndFunc

; Lam dep van ban SAU khi da format
Func _AI_BeautifyDocument()
    If Not _CheckConnection() Then Return

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $oParas = $oRange.Paragraphs

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $oStyle = $oPara.Style
        If Not IsObj($oStyle) Then ContinueLoop
        Local $sStyle = $oStyle.NameLocal

        ; === Fix Heading formatting ===
        If StringInStr($sStyle, "Heading") Or StringInStr($sStyle, "Tieu de") Then
            ; Heading khong thut dau dong
            $oPara.Format.FirstLineIndent = 0
            $oPara.Format.LeftIndent = 0

            ; Heading spacing tu settings
            If StringInStr($sStyle, "1") Then
                $oPara.Range.Font.Size = $AI_FONT_SIZE_H1
                $oPara.Range.Font.Bold = True
                $oPara.Format.SpaceBefore = $AI_H1_SPACE_BEFORE
                $oPara.Format.SpaceAfter = $AI_H1_SPACE_AFTER
            ElseIf StringInStr($sStyle, "2") Then
                $oPara.Range.Font.Size = $AI_FONT_SIZE_H2
                $oPara.Range.Font.Bold = True
                $oPara.Format.SpaceBefore = $AI_H2_SPACE_BEFORE
                $oPara.Format.SpaceAfter = $AI_H2_SPACE_AFTER
            ElseIf StringInStr($sStyle, "3") Then
                $oPara.Range.Font.Size = $AI_FONT_SIZE_H3
                $oPara.Range.Font.Bold = True
                $oPara.Range.Font.Italic = True
                $oPara.Format.SpaceBefore = $AI_H3_SPACE_BEFORE
                $oPara.Format.SpaceAfter = $AI_H3_SPACE_AFTER
            EndIf

            ; Heading luon font chuan
            $oPara.Range.Font.Name = $AI_FONT_NAME
            $oPara.Range.Font.Color = 0 ; Black
            ContinueLoop
        EndIf

        ; === Fix List formatting ===
        If $oPara.Range.ListFormat.ListType <> 0 Then ; Co list format
            $oPara.Format.FirstLineIndent = 0
            $oPara.Format.LeftIndent = $AI_LIST_INDENT * $CM_TO_POINTS
            $oPara.Format.SpaceBefore = 0
            $oPara.Format.SpaceAfter = $AI_LIST_SPACE_AFTER
            ContinueLoop
        EndIf

        ; === Fix Normal paragraph ===
        $oPara.Format.Alignment = $AI_PARA_ALIGNMENT
        $oPara.Range.Font.Color = 0
    Next

    ; === Xoa sot bullet chars sau khi xu ly ===
    If $AI_REMOVE_ORPHAN_BULLETS Then
        Local $sBulletChars = ChrW(8226) & ChrW(9702) & ChrW(9642) & ChrW(9656) & ChrW(9679)
        For $b = 1 To StringLen($sBulletChars)
            Local $sBChar = StringMid($sBulletChars, $b, 1)
            Local $oFind = _AI_GetRange().Find
            $oFind.ClearFormatting()
            $oFind.Replacement.ClearFormatting()
            $oFind.Text = $sBChar
            $oFind.Replacement.Text = ""
            $oFind.MatchWildcards = False
            _AI_ExecuteReplaceAll($oFind)
        Next
    EndIf

    _UpdateProgress("Da lam dep van ban!")
EndFunc

; ========================================
; NANG CAO FUNCTIONS
; ========================================

; LaTeX -> Word Equation
Func _AI_ConvertLaTeX()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang xu ly LaTeX...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $iCount = 0
    Local $oParas = $oRange.Paragraphs

    ; Tim cac dong co $...$ hoac $$...$$
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sText = $oPara.Range.Text
        Local $sTrimmed = StringStripWS(StringReplace(StringReplace($sText, @CR, ""), @LF, ""), 3)

        ; Tim $$ block math $$
        If StringLeft($sTrimmed, 2) = "$$" And StringRight($sTrimmed, 2) = "$$" And StringLen($sTrimmed) > 4 Then
            ; Xoa $$ markers theo paragraph de tranh lech regex/cuoi dong
            Local $iParaStart = $oPara.Range.Start
            Local $sClean = StringTrimLeft(StringTrimRight($sTrimmed, 2), 2)
            $oPara.Range.Text = $sClean & @CR
            Local $oBlockRange = $g_oDoc.Range($iParaStart, $iParaStart + StringLen($sClean))
            If IsObj($oBlockRange) Then
                $oBlockRange.Font.Name = "Cambria Math"
                $oBlockRange.Font.Italic = True
            EndIf

            Local $oBlockPara = 0
            If $i <= $g_oDoc.Paragraphs.Count Then $oBlockPara = $g_oDoc.Paragraphs.Item($i)
            If IsObj($oBlockPara) Then $oBlockPara.Format.Alignment = $WD_ALIGN_CENTER
            $iCount += 1
        ; Tim $ inline math $
        ElseIf StringRegExp($sText, "\$[^\$]+\$") Then
            Local $iSafe = 0
            While $iSafe < 20
                $iSafe += 1
                Local $sText2 = $oPara.Range.Text
                Local $iParaStart = $oPara.Range.Start
                Local $iOpen = _AI_FindSingleDollar($sText2, 1)
                If $iOpen = 0 Then ExitLoop

                Local $iClose = _AI_FindSingleDollar($sText2, $iOpen + 1)
                If $iClose <= ($iOpen + 1) Then ExitLoop

                Local $iWordOpen = $iParaStart + $iOpen - 1
                Local $iWordClose = $iParaStart + $iClose - 1

                Local $oMathRange = $g_oDoc.Range($iWordOpen + 1, $iWordClose)
                If IsObj($oMathRange) Then
                    $oMathRange.Font.Name = "Cambria Math"
                    $oMathRange.Font.Italic = True
                EndIf

                Local $oCloseMarker = $g_oDoc.Range($iWordClose, $iWordClose + 1)
                If IsObj($oCloseMarker) Then $oCloseMarker.Delete()
                Local $oOpenMarker = $g_oDoc.Range($iWordOpen, $iWordOpen + 1)
                If IsObj($oOpenMarker) Then $oOpenMarker.Delete()

                $iCount += 1
            WEnd
        EndIf
    Next

    _UpdateProgress("Da xu ly " & $iCount & " cong thuc LaTeX!")
    If $iCount = 0 Then
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong tim thay cong thuc LaTeX ($...$)")
    EndIf
EndFunc

Func _AI_FindSingleDollar($sText, $iStartPos)
    Local $iLen = StringLen($sText)
    If $iStartPos < 1 Then $iStartPos = 1

    For $i = $iStartPos To $iLen
        If StringMid($sText, $i, 1) <> "$" Then ContinueLoop

        Local $sPrev = ""
        If $i > 1 Then $sPrev = StringMid($sText, $i - 1, 1)
        Local $sNext = ""
        If $i < $iLen Then $sNext = StringMid($sText, $i + 1, 1)

        If $sPrev = "$" Or $sNext = "$" Then ContinueLoop
        Return $i
    Next

    Return 0
EndFunc

; Xoa Emoji
Func _AI_RemoveEmoji()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang xoa emoji...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $iCount = 0
    Local $oParas = $oRange.Paragraphs

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sText = $oPara.Range.Text
        ; Xoa cac ky tu Unicode ngoai BMP (emoji thường > 0xFFFF)
        ; Va cac ky tu symbol thuong gap
        Local $sClean = StringRegExpReplace($sText, "[\x{1F300}-\x{1F9FF}]", "")
        $sClean = StringRegExpReplace($sClean, "[\x{2600}-\x{27BF}]", "")
        $sClean = StringRegExpReplace($sClean, "[\x{FE00}-\x{FE0F}]", "")

        If $sClean <> $sText Then
            $oPara.Range.Text = $sClean
            $iCount += 1
        EndIf
    Next

    _UpdateProgress("Da xoa emoji tu " & $iCount & " dong!")
EndFunc

; Fix encoding VN
Func _AI_FixEncoding()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang fix encoding VN...")

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    ; Chuyen smart quotes thanh straight quotes (chuan do an)
    ; ChrW(8220) = left double quote, ChrW(8221) = right double quote
    ; ChrW(8216) = left single quote, ChrW(8217) = right single quote
    ; ChrW(8211) = en-dash, ChrW(8212) = em-dash

    Local $oFind = _AI_GetRange().Find
    _AI_ReplaceSmartQuotesInFind($oFind)

    Local $aFixes[2][2] = [ _
        [ChrW(8212), " - "], _
        [ChrW(8211), "-"] _
    ]

    For $f = 0 To UBound($aFixes) - 1
        Local $oFindDash = _AI_GetRange().Find
        $oFindDash.ClearFormatting()
        $oFindDash.Replacement.ClearFormatting()
        $oFindDash.Text = $aFixes[$f][0]
        $oFindDash.Replacement.Text = $aFixes[$f][1]
        $oFindDash.MatchWildcards = False
        _AI_ExecuteReplaceAll($oFindDash)
    Next

    _UpdateProgress("Da fix encoding VN!")
EndFunc

; Preview truoc/sau
Func _AI_PreviewChanges()
    If Not _CheckConnection() Then Return

    Local $oRange = _AI_GetRange()
    If Not IsObj($oRange) Then Return

    Local $sText = StringLeft($oRange.Text, 500) ; Lay 500 ky tu dau
    Local $sMsg = "PREVIEW NOI DUNG:" & @CRLF & @CRLF
    $sMsg &= "Font: " & $oRange.Font.Name & " " & $oRange.Font.Size & "pt" & @CRLF

    Local $sBold = "Khong"
    If $oRange.Font.Bold = True Then $sBold = "Co"
    $sMsg &= "Bold: " & $sBold & @CRLF

    ; Dem markdown patterns
    Local $iHeadings = 0
    Local $aH = StringRegExp($sText, "(?m)^#{1,6}\s+[^\r\n]+", 3)
    If Not @error Then $iHeadings = UBound($aH)
    Local $iBold = 0
    Local $aB = StringRegExp($sText, "\*\*[^\*]+\*\*", 3)
    If Not @error Then $iBold = UBound($aB)
    Local $iCode = 0
    Local $aC = StringRegExp($sText, "```", 3)
    If Not @error Then $iCode = UBound($aC)
    Local $iLinks = 0
    Local $aL = StringRegExp($sText, "\[[^\]]+\]\([^\)]+\)", 3)
    If Not @error Then $iLinks = UBound($aL)

    $sMsg &= @CRLF & "MARKDOWN TIM THAY:" & @CRLF
    $sMsg &= "Headings (##): ~" & $iHeadings & @CRLF
    $sMsg &= "Bold (**): ~" & $iBold & @CRLF
    $sMsg &= "Code blocks (```): ~" & Int($iCode / 2) & @CRLF
    $sMsg &= "Links [...]: ~" & $iLinks & @CRLF
    $sMsg &= @CRLF & "NOI DUNG (500 ky tu dau):" & @CRLF & @CRLF
    $sMsg &= $sText

    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "Preview", $sMsg)
EndFunc
