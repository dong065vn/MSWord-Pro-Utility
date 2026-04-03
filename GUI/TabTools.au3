; ============================================
; TABTOOLS.AU3 - Tab 3: Cong cu
; ============================================

#include-once

Func _CreateTabTools()
    GUICtrlCreateTabItem("Cong cu")

    ; Tim va thay the
    _CreateGroup(" Tim va thay the ", 35, 160, 710, 115)
    _CreateLabel("Tim:", 50, 188, 35, 20)
    $g_inputFind = _CreateInput("", 90, 185, 280, 24)
    _CreateLabel("Thay:", 385, 188, 40, 20)
    $g_inputReplace = _CreateInput("", 430, 185, 290, 24)
    $g_chkMatchCase = _CreateCheckbox(" Phan biet hoa/thuong", 50, 215, 160, 22)
    $g_chkWholeWord = _CreateCheckbox(" Nguyen tu", 220, 215, 100, 22)
    $g_btnFindNext = _CreateButton("Tim tiep", 430, 213, 90, 26)
    $g_btnFindReplace = _CreateButton("Thay the tat ca", 530, 213, 110, 26)
    $g_btnPreviewParentheses = _CreateButton("Xem truoc (...)", 430, 245, 110, 26)
    $g_btnRemoveParenthesesSelection = _CreateButton("Xoa vung chon", 545, 245, 95, 26)
    $g_btnRemoveParenthesesText = _CreateButton("Xoa toan file", 645, 245, 95, 26)
    _EndGroup()

    ; Xu ly hinh anh
    _CreateGroup(" Xu ly hinh anh ", 35, 285, 340, 110)
    $g_btnResizeImages = _CreateButton("Resize theo trang", 50, 310, 140, 28)
    $g_btnCenterImages = _CreateButton("Can giua tat ca", 200, 310, 150, 28)
    $g_btnAutoCaption = _CreateButton("Them caption", 50, 345, 140, 28)
    $g_btnRemoveImages = _CreateButton("Xoa tat ca anh", 200, 345, 150, 28)
    _EndGroup()

    ; Xu ly bang
    _CreateGroup(" Xu ly bang ", 395, 285, 350, 110)
    $g_btnAutoFitTable = _CreateButton("AutoFit noi dung", 410, 310, 140, 28)
    $g_btnAutoFitWindow = _CreateButton("AutoFit cua so", 560, 310, 160, 28)
    $g_btnTableCaption = _CreateButton("Them caption", 410, 345, 140, 28)
    $g_btnTableBorder = _CreateButton("Them vien bang", 560, 345, 160, 28)
    _EndGroup()

    ; Thong ke & Kiem tra
    _CreateGroup(" Thong ke & Kiem tra ", 35, 405, 710, 60)
    $g_btnWordCount = _CreateButton("Dem tu chi tiet", 50, 428, 130, 28)
    $g_btnCheckSpelling = _CreateButton("Kiem tra chinh ta", 190, 428, 130, 28)
    $g_btnCheckFormat = _CreateButton("Kiem tra dinh dang", 330, 428, 130, 28)
    $g_btnShowStats = _CreateButton("Thong ke tong hop", 470, 428, 130, 28)
    $g_btnExportStats = _CreateButton("Xuat bao cao", 610, 428, 110, 28)
    _EndGroup()
EndFunc
