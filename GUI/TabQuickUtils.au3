; ============================================
; TABQUICKUTILS.AU3 - Tab Tien ich Nhanh
; ============================================

#include-once

Func _CreateTabQuickUtils()
    GUICtrlCreateTabItem("Quick Utils")
    
    Local $iY = 155
    Local $iColW = 230
    Local $iGap = 10
    
    ; === COT 1: PASTE & SELECT & NAVIGATE ===
    _CreateGroup(" Dan thong minh ", 30, $iY, $iColW, 75)
    $g_btnPastePlain = _CreateButton("Dan Text", 40, $iY + 20, 70, 26)
    GUICtrlSetTip(-1, "Dan khong dinh dang (chi text)")
    $g_btnPasteKeep = _CreateButton("Giu nguon", 115, $iY + 20, 70, 26)
    GUICtrlSetTip(-1, "Dan giu dinh dang nguon")
    $g_btnPasteMerge = _CreateButton("Hop nhat", 190, $iY + 20, 60, 26)
    GUICtrlSetTip(-1, "Dan va hop nhat dinh dang")
    _EndGroup()
    
    _CreateGroup(" Chon nhanh ", 30, $iY + 80, $iColW, 75)
    $g_btnSelectPara = _CreateButton("Doan van", 40, $iY + 100, 70, 26)
    $g_btnSelectSentence = _CreateButton("Cau", 115, $iY + 100, 50, 26)
    $g_btnSelectFromStart = _CreateButton("Tu dau", 170, $iY + 100, 40, 26)
    $g_btnSelectToEnd = _CreateButton("Cuoi", 215, $iY + 100, 35, 26)
    _EndGroup()
    
    _CreateGroup(" Di chuyen nhanh ", 30, $iY + 160, $iColW, 105)
    $g_btnGoToPage = _CreateButton("Den trang...", 40, $iY + 180, 100, 26)
    $g_btnGoToNextHeading = _CreateButton("Heading >", 145, $iY + 180, 105, 26)
    $g_btnGoToPrevHeading = _CreateButton("< Heading", 40, $iY + 212, 100, 26)
    $g_btnGoToNextTable = _CreateButton("Bang >", 145, $iY + 212, 50, 26)
    $g_btnGoToNextImage = _CreateButton("Hinh >", 200, $iY + 212, 50, 26)
    _EndGroup()
    
    ; === COT 2: INSERT & FORMAT ===
    _CreateGroup(" Chen nhanh ", 30 + $iColW + $iGap, $iY, $iColW, 105)
    $g_btnInsertPageBreak = _CreateButton("Ngat trang", 40 + $iColW + $iGap, $iY + 20, 70, 26)
    $g_btnInsertSectionBreak = _CreateButton("Ngat section", 115 + $iColW + $iGap, $iY + 20, 75, 26)
    $g_btnInsertDate = _CreateButton("Ngay", 195 + $iColW + $iGap, $iY + 20, 55, 26)
    $g_btnInsertTime = _CreateButton("Gio", 40 + $iColW + $iGap, $iY + 52, 55, 26)
    $g_btnInsertHLine = _CreateButton("Duong ke", 100 + $iColW + $iGap, $iY + 52, 70, 26)
    $g_btnInsertSpecialChar = _CreateButton("Ky tu...", 175 + $iColW + $iGap, $iY + 52, 75, 26)
    GUICtrlSetTip(-1, "Chen ky tu dac biet")
    _EndGroup()
    
    _CreateGroup(" Dinh dang Font ", 30 + $iColW + $iGap, $iY + 110, $iColW, 75)
    $g_btnFontIncrease = _CreateButton("A+", 40 + $iColW + $iGap, $iY + 130, 35, 26)
    GUICtrlSetTip(-1, "Tang kich thuoc font")
    $g_btnFontDecrease = _CreateButton("A-", 80 + $iColW + $iGap, $iY + 130, 35, 26)
    GUICtrlSetTip(-1, "Giam kich thuoc font")
    $g_btnToggleBold = _CreateButton("B", 120 + $iColW + $iGap, $iY + 130, 30, 26)
    GUICtrlSetFont(-1, 10, 700)
    $g_btnToggleItalic = _CreateButton("I", 155 + $iColW + $iGap, $iY + 130, 30, 26)
    GUICtrlSetFont(-1, 10, 400, 2)
    $g_btnToggleUnderline = _CreateButton("U", 190 + $iColW + $iGap, $iY + 130, 30, 26)
    $g_btnToggleSub = _CreateButton("x2", 225 + $iColW + $iGap, $iY + 130, 25, 26)
    GUICtrlSetTip(-1, "Subscript")
    $g_btnToggleSuper = _CreateButton("x2", 225 + $iColW + $iGap, $iY + 158, 25, 26)
    GUICtrlSetTip(-1, "Superscript")
    _EndGroup()
    
    _CreateGroup(" Thut le doan van ", 30 + $iColW + $iGap, $iY + 190, $iColW, 75)
    $g_btnIncreaseIndent = _CreateButton("Tang >>", 40 + $iColW + $iGap, $iY + 210, 70, 26)
    $g_btnDecreaseIndent = _CreateButton("<< Giam", 115 + $iColW + $iGap, $iY + 210, 70, 26)
    $g_btnSetFirstIndent = _CreateButton("Thut 1.27cm", 190 + $iColW + $iGap, $iY + 210, 60, 26)
    $g_btnRemoveFirstIndent = _CreateButton("Xoa thut dau dong", 40 + $iColW + $iGap, $iY + 238, 210, 22)
    _EndGroup()
    
    ; === COT 3: BOOKMARK & INFO ===
    _CreateGroup(" Bookmark ", 30 + ($iColW + $iGap) * 2, $iY, $iColW, 75)
    $g_btnAddBookmark = _CreateButton("Them", 40 + ($iColW + $iGap) * 2, $iY + 20, 70, 26)
    $g_btnGoToBookmark = _CreateButton("Nhay den", 115 + ($iColW + $iGap) * 2, $iY + 20, 70, 26)
    $g_btnDeleteBookmark = _CreateButton("Xoa", 190 + ($iColW + $iGap) * 2, $iY + 20, 60, 26)
    _EndGroup()
    
    _CreateGroup(" Thong tin & Don dep ", 30 + ($iColW + $iGap) * 2, $iY + 80, $iColW, 105)
    $g_btnShowDocInfo = _CreateButton("Thong tin chi tiet", 40 + ($iColW + $iGap) * 2, $iY + 100, 210, 28)
    GUICtrlSetBkColor(-1, 0x3498DB)
    GUICtrlSetFont(-1, 9, 600)
    $g_btnRemoveHighlightSel = _CreateButton("Xoa Highlight", 40 + ($iColW + $iGap) * 2, $iY + 132, 100, 26)
    $g_btnRemoveCommentsSel = _CreateButton("Xoa Comment", 145 + ($iColW + $iGap) * 2, $iY + 132, 105, 26)
    _EndGroup()
    
    _CreateGroup(" Chuyen doi Field ", 30 + ($iColW + $iGap) * 2, $iY + 190, $iColW, 75)
    $g_btnUnlinkFields = _CreateButton("Unlink All Fields", 40 + ($iColW + $iGap) * 2, $iY + 210, 210, 28)
    GUICtrlSetBkColor(-1, 0xE74C3C)
    GUICtrlSetTip(-1, "Chuyen tat ca Field thanh text (khong the hoan tac)")
    _EndGroup()
    
    ; === SUA DE MUC ===
    _CreateGroup(" Sua de muc so ", 30, $iY + 270, ($iColW + $iGap) * 3, 65)
    _CreateLabel("Tien to:", 40, $iY + 293, 45, 20)
    $g_inputHeadingPrefixFix = _CreateInput("", 88, $iY + 289, 120, 24)
    GUICtrlSetTip(-1, "Nhap 1. hoac 2. de chi sua cac de muc bat dau bang tien to do. De trong de sua toan bo.")
    _CreateLabel("Ngat sau so:", 220, $iY + 293, 70, 20)
    $g_inputHeadingSeparatorFix = _CreateInput(" ", 295, $iY + 289, 80, 24)
    GUICtrlSetTip(-1, "Mac dinh la 1 dau cach. Vi du 2.4. Tieu de se thanh 2.4 Tieu de")
    $g_btnFixHeadingNumberDots = _CreateButton("Sua 2.4. -> 2.4", 390, $iY + 286, 145, 28, 0x16A085)
    GUICtrlSetTip(-1, "Xoa dau cham cuoi cung trong de muc dang 1. / 2.4. / 3.2.1.")
    GUICtrlCreateLabel("Vi du: tien to 2. se sua toan bo 2.1., 2.2., 2.4. thanh 2.1, 2.2, 2.4", _
        545, $iY + 289, 180, 30)
    _EndGroup()

    ; === HUONG DAN ===
    _CreateGroup(" Huong dan su dung ", 30, $iY + 340, ($iColW + $iGap) * 3, 55)
    GUICtrlCreateLabel("Tab nay chua cac tien ich nhanh: Dan thong minh | Chon nhanh (doan van, cau) | Di chuyen (trang, heading, bang, hinh) | Chen nhanh | Dinh dang Font | Bookmark | Sua de muc so", _
        40, $iY + 360, ($iColW + $iGap) * 3 - 20, 30)
    GUICtrlSetFont(-1, 8, 400)
    GUICtrlCreateGroup("", -99, -99, 1, 1)
EndFunc
