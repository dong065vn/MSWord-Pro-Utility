; ============================================
; TABFORMAT.AU3 - Tab 2: Dinh dang
; ============================================

#include-once

Func _CreateTabFormat()
    GUICtrlCreateTabItem("Dinh dang")

    ; Thong so dinh dang
    _CreateGroup(" Thong so dinh dang ", 35, 160, 710, 100)
    _CreateLabel("Font:", 50, 188, 40, 20)
    $g_cboFont = _CreateCombo(95, 185, 140, 25, "Times New Roman|Arial|Calibri|Tahoma", "Times New Roman")
    _CreateLabel("Co chu:", 250, 188, 50, 20)
    $g_cboFontSize = _CreateCombo(305, 185, 55, 25, "11|12|13|14|15|16", "13")
    _CreateLabel("Gian dong:", 375, 188, 65, 20)
    $g_cboLineSpacing = _CreateCombo(445, 185, 60, 25, "1.0|1.15|1.5|2.0", "1.5")
    _CreateLabel("Canh le:", 520, 188, 50, 20)
    $g_cboAlignment = _CreateCombo(575, 185, 120, 25, "Justify|Left|Center|Right", "Justify")

    _CreateLabel("Le trai:", 50, 223, 45, 20)
    $g_inputLeftMargin = _CreateInput("3.5", 100, 220, 45, 22)
    _CreateLabel("Le phai:", 160, 223, 50, 20)
    $g_inputRightMargin = _CreateInput("2", 215, 220, 45, 22)
    _CreateLabel("Le tren:", 275, 223, 50, 20)
    $g_inputTopMargin = _CreateInput("2.5", 330, 220, 45, 22)
    _CreateLabel("Le duoi:", 390, 223, 50, 20)
    $g_inputBottomMargin = _CreateInput("2.5", 445, 220, 45, 22)
    _CreateLabel("cm", 495, 223, 25, 20)
    $g_chkAutoFirstLine = _CreateCheckbox(" Thut dau dong (1.27cm)", 540, 220, 170, 22, True)
    _EndGroup()

    ; Buttons row 1
    $g_btnApplyFormat = _CreateButton("Ap dung toan bo", 35, 270, 130, 35)
    $g_btnApplySelection = _CreateButton("Ap dung vung chon", 175, 270, 130, 35)
    $g_btnPresetVN = _CreateButton("Chuan VN", 315, 270, 100, 35, 0x3498DB)
    $g_btnPresetUS = _CreateButton("Chuan US/APA", 425, 270, 100, 35)
    $g_btnCheckThesis = _CreateButton("Kiem tra do an", 535, 270, 120, 35, 0xF39C12)

    ; Heading
    _CreateGroup(" Dinh dang tieu de (Heading) ", 35, 315, 340, 70)
    $g_btnFormatH1 = _CreateButton("H1-Chuong", 50, 340, 95, 30)
    $g_btnFormatH2 = _CreateButton("H2-Muc", 155, 340, 80, 30)
    $g_btnFormatH3 = _CreateButton("H3-Tieu muc", 245, 340, 95, 30)
    _EndGroup()

    ; Dinh dang khac
    _CreateGroup(" Dinh dang khac ", 385, 315, 360, 70)
    $g_btnFormatCaption = _CreateButton("Caption", 400, 340, 80, 30)
    $g_btnFormatNormal = _CreateButton("Normal", 490, 340, 80, 30)
    $g_btnClearFormat = _CreateButton("Xoa dinh dang", 580, 340, 110, 30)
    _EndGroup()

    ; Thao tac nhanh
    _CreateGroup(" Thao tac nhanh ", 35, 395, 710, 60)
    $g_btnRemoveHighlight = _CreateButton("Xoa highlight", 50, 418, 105, 28)
    $g_btnUnifyFont = _CreateButton("Thong nhat font", 165, 418, 110, 28)
    $g_btnFixAllSpacing = _CreateButton("Sua gian dong", 285, 418, 105, 28)
    $g_btnAddPageNum = _CreateButton("Them so trang", 400, 418, 105, 28)
    $g_btnAddHeader = _CreateButton("Them header", 515, 418, 100, 28)
    $g_btnRemovePageNum = _CreateButton("Xoa so trang", 625, 418, 100, 28)
    _EndGroup()

    ; Danh so tu dong
    _CreateGroup(" Danh so tu dong (Caption) ", 35, 465, 710, 60)
    $g_btnAutoNumImg = _CreateButton("Danh so Hinh", 50, 488, 110, 28)
    $g_btnAutoNumTbl = _CreateButton("Danh so Bang", 165, 488, 110, 28)
    $g_btnAutoNumEq = _CreateButton("Danh so CT", 280, 488, 100, 28)
    $g_btnRemoveNumEq = _CreateButton("Xoa so CT", 385, 488, 80, 28)
    _CreateLabel("Tien to:", 475, 492, 40, 20)
    $g_inputCaptionPrefix = _CreateInput("Hinh", 520, 488, 60, 24)
    _CreateLabel("Chuong:", 590, 492, 50, 20)
    $g_inputChapterNum = _CreateInput("1", 645, 488, 35, 24)
    _EndGroup()
EndFunc
