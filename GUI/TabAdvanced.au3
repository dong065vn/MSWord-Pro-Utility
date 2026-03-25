; ============================================
; TABADVANCED.AU3 - Tab 6: Nang cao
; ============================================

#include-once

Func _CreateTabAdvanced()
    GUICtrlCreateTabItem("Nang cao")

    ; Xu ly Heading nang cao
    _CreateGroup(" Xu ly Heading nang cao ", 35, 160, 710, 110)
    $g_btnAutoHeading = _CreateButton("Tu dong gan Heading", 50, 185, 150, 30)
    $g_btnResetHeading = _CreateButton("Reset tat ca Heading", 210, 185, 140, 30)
    $g_btnHeadingToTOC = _CreateButton("Heading -> Muc luc", 360, 185, 140, 30)
    $g_btnListHeadings = _CreateButton("Liet ke Heading", 510, 185, 120, 30)
    $g_btnScanThesisHeadings = _CreateButton("Quet de muc + chon style", 50, 220, 190, 28, 0x16A085)
    GUICtrlSetTip(-1, "Quet chuong/de muc theo mau do an va cho chon style co san truoc khi ap dung")
    _CreateLabel("Tu dong nhan dien va gan Heading dua tren dinh dang", 50, 250, 500, 20)
    _CreateLabel("Quet chuong/de muc theo so thu tu va ap dung style nguoi dung chon", 250, 224, 430, 20)
    _EndGroup()

    ; Xu ly van ban nang cao
    _CreateGroup(" Xu ly van ban nang cao ", 35, 280, 710, 100)
    $g_btnRemoveAllFormat = _CreateButton("Xoa tat ca dinh dang", 50, 305, 140, 28)
    $g_btnConvertCase = _CreateButton("Chuyen doi chu", 200, 305, 110, 28)
    $g_btnRemoveHyperlinks = _CreateButton("Xoa Hyperlinks", 320, 305, 110, 28)
    $g_btnRemoveComments = _CreateButton("Xoa Comments", 440, 305, 110, 28)
    $g_btnAcceptChanges = _CreateButton("Chap nhan thay doi", 560, 305, 130, 28)
    $g_btnNumberToText = _CreateButton("Numbering -> Text", 50, 340, 140, 28, 0x9B59B6)
    $g_btnBulletToText = _CreateButton("Bullet -> Text", 200, 340, 110, 28)
    $g_btnNumberToTextSel = _CreateButton("So vung chon -> Text", 320, 340, 140, 28)
    _CreateLabel("Chuyen danh so tu dong thanh text giu nguyen so", 480, 345, 250, 20)
    _EndGroup()

    ; Xuat file
    _CreateGroup(" Xuat file ", 35, 390, 710, 80)
    $g_btnExportPDF = _CreateButton("Xuat PDF", 50, 415, 120, 30, 0xE74C3C)
    $g_btnExportHTML = _CreateButton("Xuat HTML", 180, 415, 120, 30)
    $g_btnExportTXT = _CreateButton("Xuat TXT", 310, 415, 120, 30)
    $g_btnExportRTF = _CreateButton("Xuat RTF", 440, 415, 120, 30)
    $g_btnPrintPreview = _CreateButton("Xem truoc in", 570, 415, 120, 30)
    _EndGroup()

    ; Tien ich khac
    _CreateGroup(" Tien ich khac ", 35, 480, 710, 90)
    $g_btnCompareDoc = _CreateButton("So sanh 2 file", 50, 505, 120, 30)
    $g_btnMergeDoc = _CreateButton("Gop file", 180, 505, 100, 30)
    $g_btnSplitDoc = _CreateButton("Tach file", 290, 505, 100, 30)
    $g_btnProtectDoc = _CreateButton("Bao ve file", 400, 505, 100, 30)
    $g_btnDocProperties = _CreateButton("Thuoc tinh file", 510, 505, 110, 30)
    $g_btnCleanDoc = _CreateButton("Don dep file", 630, 505, 100, 30)
    _CreateLabel("Cac tien ich bo sung cho quan ly tai lieu", 50, 540, 400, 20)
    _EndGroup()
EndFunc
