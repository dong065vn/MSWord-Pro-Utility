; ============================================
; TABPDFFIX.AU3 - Tab 1: Sua loi PDF
; ============================================

#include-once

Func _CreateTabPDFFix()
    GUICtrlCreateTabItem("Sua loi PDF")

    ; Cac loi pho bien
    _CreateGroup(" Cac loi pho bien tu PDF ", 35, 160, 710, 120)
    $g_chkLineBreaks = _CreateCheckbox(" Xoa xuong dong thua (^l)", 50, 185, 200, 22, True)
    $g_chkExtraSpaces = _CreateCheckbox(" Xoa khoang trang thua", 260, 185, 200, 22, True)
    $g_chkHyphenation = _CreateCheckbox(" Noi tu bi ngat (hyphen)", 470, 185, 200, 22, True)
    $g_chkSpecialChars = _CreateCheckbox(" Sua ky tu dac biet", 50, 210, 200, 22, True)
    $g_chkParagraphs = _CreateCheckbox(" Chuan hoa doan van", 260, 210, 200, 22, True)
    $g_chkTabs = _CreateCheckbox(" Xoa tab thua", 470, 210, 200, 22, True)
    $g_chkVietnamese = _CreateCheckbox(" Sua loi font tieng Viet", 50, 235, 150, 22)
    $g_chkPageNumbers = _CreateCheckbox(" Xoa so trang loi", 220, 235, 150, 22)
    $g_chkFixQuotes = _CreateCheckbox(" Sua dau ngoac kep", 390, 235, 150, 22, True)
    $g_chkRemoveFakeNumbering = _CreateCheckbox(" Bo so dau dong thua", 560, 235, 150, 22)
    _EndGroup()

    ; Xu ly cach dong
    _CreateGroup(" Xu ly cach dong ", 35, 285, 710, 60)
    $g_chkFixLineSpacing = _CreateCheckbox(" Fix cach dong (1.5)", 50, 308, 160, 22, True)
    $g_chkResetSpacing = _CreateCheckbox(" Reset cach dong (1.0)", 220, 308, 160, 22)
    $g_chkRemoveEmptyLines = _CreateCheckbox(" Xoa dong trong thua", 390, 308, 160, 22, True)
    $g_chkFixSpacingBefore = _CreateCheckbox(" Xoa spacing thua", 560, 308, 140, 22, True)
    _EndGroup()

    ; Buttons - Hang 1
    $g_btnFixSelected = _CreateButton("Sua vung chon", 35, 360, 100, 40, -1, 9, 600)
    $g_btnFixAll = _CreateButton("Sua toan bo", 145, 360, 100, 40, -1, 9, 600)
    $g_btnQuickFix = _CreateButton("Quick Fix", 255, 360, 90, 40, 0x27AE60, 9, 700)
    $g_btnCleanUp = _CreateButton("Clean Up", 355, 360, 90, 40, 0x9B59B6, 9, 700)
    $g_btnFixLayout = _CreateButton("Fix Layout", 455, 360, 90, 40, 0x3498DB, 9, 700)
    $g_btnUndoFix = _CreateButton("Undo", 555, 360, 70, 40, -1, 9, 600)
    $g_btnPDFFixHelp = _CreateButton("?", 635, 360, 40, 40, 0xF39C12, 12, 700)

    ; Preview
    _CreateGroup(" Xem truoc / Log ", 35, 410, 710, 255)
    $g_editPreview = GUICtrlCreateEdit("", 50, 435, 680, 215, _
        BitOR($ES_MULTILINE, $ES_READONLY, $WS_VSCROLL))
    GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
    _EndGroup()
EndFunc
