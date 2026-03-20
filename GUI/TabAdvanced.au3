; ============================================
; TABADVANCED.AU3 - Tab 6: Nang cao
; ============================================

#include-once

Func _CreateTabAdvanced()
    GUICtrlCreateTabItem("Nang cao")

    ; Xu ly Heading nang cao
    _CreateGroup(" Xu ly Heading nang cao ", 35, 160, 710, 90)
    $g_btnAutoHeading = _CreateButton("Tu dong gan Heading", 50, 185, 150, 30)
    $g_btnResetHeading = _CreateButton("Reset tat ca Heading", 210, 185, 140, 30)
    $g_btnHeadingToTOC = _CreateButton("Heading -> Muc luc", 360, 185, 140, 30)
    $g_btnListHeadings = _CreateButton("Liet ke Heading", 510, 185, 120, 30)
    _CreateLabel("Tu dong nhan dien va gan Heading dua tren dinh dang", 50, 220, 500, 20)
    _EndGroup()

    ; Xu ly van ban nang cao
    _CreateGroup(" Xu ly van ban nang cao ", 35, 260, 710, 100)
    $g_btnRemoveAllFormat = _CreateButton("Xoa tat ca dinh dang", 50, 285, 140, 28)
    $g_btnConvertCase = _CreateButton("Chuyen doi chu", 200, 285, 110, 28)
    $g_btnRemoveHyperlinks = _CreateButton("Xoa Hyperlinks", 320, 285, 110, 28)
    $g_btnRemoveComments = _CreateButton("Xoa Comments", 440, 285, 110, 28)
    $g_btnAcceptChanges = _CreateButton("Chap nhan thay doi", 560, 285, 130, 28)
    $g_btnNumberToText = _CreateButton("Numbering -> Text", 50, 320, 140, 28, 0x9B59B6)
    $g_btnBulletToText = _CreateButton("Bullet -> Text", 200, 320, 110, 28)
    $g_btnNumberToTextSel = _CreateButton("So vung chon -> Text", 320, 320, 140, 28)
    _CreateLabel("Chuyen danh so tu dong thanh text giu nguyen so", 480, 325, 250, 20)
    _EndGroup()

    ; Xuat file
    _CreateGroup(" Xuat file ", 35, 370, 710, 80)
    $g_btnExportPDF = _CreateButton("Xuat PDF", 50, 395, 120, 30, 0xE74C3C)
    $g_btnExportHTML = _CreateButton("Xuat HTML", 180, 395, 120, 30)
    $g_btnExportTXT = _CreateButton("Xuat TXT", 310, 395, 120, 30)
    $g_btnExportRTF = _CreateButton("Xuat RTF", 440, 395, 120, 30)
    $g_btnPrintPreview = _CreateButton("Xem truoc in", 570, 395, 120, 30)
    _EndGroup()

    ; Tien ich khac
    _CreateGroup(" Tien ich khac ", 35, 460, 710, 90)
    $g_btnCompareDoc = _CreateButton("So sanh 2 file", 50, 485, 120, 30)
    $g_btnMergeDoc = _CreateButton("Gop file", 180, 485, 100, 30)
    $g_btnSplitDoc = _CreateButton("Tach file", 290, 485, 100, 30)
    $g_btnProtectDoc = _CreateButton("Bao ve file", 400, 485, 100, 30)
    $g_btnDocProperties = _CreateButton("Thuoc tinh file", 510, 485, 110, 30)
    $g_btnCleanDoc = _CreateButton("Don dep file", 630, 485, 100, 30)
    _CreateLabel("Cac tien ich bo sung cho quan ly tai lieu", 50, 520, 400, 20)
    _EndGroup()
EndFunc
