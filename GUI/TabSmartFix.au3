; ============================================
; TABSMARTFIX.AU3 - Tab Sua loi Thong minh
; ============================================

#include-once

Func _CreateTabSmartFix()
    GUICtrlCreateTabItem("Smart Fix")
    
    Local $iY = 155
    Local $iColW = 230
    Local $iGap = 10
    
    ; === COT 1: PHAN TICH & SUA TU DONG ===
    _CreateGroup(" Phan tich & Sua tu dong ", 30, $iY, $iColW, 160)
    
    GUICtrlCreateLabel("Smart Fix tu dong phat hien va sua:", 40, $iY + 18, 210, 16)
    GUICtrlSetFont(-1, 8, 600)
    
    GUICtrlCreateLabel("- Manual Line Break (Shift+Enter)", 45, $iY + 36, 200, 14)
    GUICtrlCreateLabel("- Khoang trang thua, Dong trong thua", 45, $iY + 50, 200, 14)
    GUICtrlCreateLabel("- Tab thua, Smart Quotes", 45, $iY + 64, 200, 14)
    GUICtrlCreateLabel("- Bang troi noi, Hinh qua lon", 45, $iY + 78, 200, 14)
    
    $g_btnSmartAnalyze = _CreateButton("PHAN TICH", 40, $iY + 98, 100, 30)
    GUICtrlSetBkColor(-1, 0x3498DB)
    GUICtrlSetFont(-1, 9, 600)
    GUICtrlSetTip(-1, "Phan tich tai lieu va bao cao cac van de")
    
    $g_btnSmartFixAll = _CreateButton("SMART FIX", 145, $iY + 98, 105, 30)
    GUICtrlSetBkColor(-1, 0x27AE60)
    GUICtrlSetFont(-1, 9, 700)
    GUICtrlSetTip(-1, "Tu dong sua tat ca cac loi tim thay")
    _EndGroup()
    
    ; === COT 1: SUA LOI CU THE ===
    _CreateGroup(" Sua loi cu the ", 30, $iY + 165, $iColW, 100)
    
    $g_btnFixHyphenation = _CreateButton("Fix Hyphenation", 40, $iY + 185, 105, 26)
    GUICtrlSetTip(-1, "Sua tu bi ngat dong (vi-du -> vidu)")
    
    $g_btnFixNonBreaking = _CreateButton("Fix Non-breaking", 150, $iY + 185, 100, 26)
    GUICtrlSetTip(-1, "Chuyen Non-breaking Space thanh Space thuong")
    
    $g_btnFixDashes = _CreateButton("Fix Dashes (Em/En Dash)", 40, $iY + 217, 210, 26)
    GUICtrlSetTip(-1, "Chuyen Em Dash va En Dash thanh gach ngang thuong")
    _EndGroup()
    
    ; === COT 2: XU LY HANG LOAT ===
    _CreateGroup(" Xu ly hang loat ", 30 + $iColW + $iGap, $iY, $iColW, 105)
    
    GUICtrlCreateLabel("Xu ly nhieu file Word cung luc:", 45 + $iColW + $iGap, $iY + 18, 200, 16)
    GUICtrlSetFont(-1, 8, 600)
    
    GUICtrlCreateLabel("- Chon thu muc chua file .docx", 50 + $iColW + $iGap, $iY + 36, 200, 14)
    GUICtrlCreateLabel("- Smart Fix tu dong cho tung file", 50 + $iColW + $iGap, $iY + 50, 200, 14)
    
    $g_btnBatchProcess = _CreateButton("XU LY HANG LOAT...", 45 + $iColW + $iGap, $iY + 70, 200, 28)
    GUICtrlSetBkColor(-1, 0x9B59B6)
    GUICtrlSetFont(-1, 9, 600)
    GUICtrlSetTip(-1, "Chon thu muc va xu ly tat ca file Word trong do")
    _EndGroup()
    
    ; === COT 2: PRESET CHUAN LUAN VAN VN ===
    _CreateGroup(" Chuan Luan van Viet Nam ", 30 + $iColW + $iGap, $iY + 110, $iColW, 105)
    
    GUICtrlCreateLabel("Times New Roman 13pt, gian dong 1.5", 45 + $iColW + $iGap, $iY + 128, 200, 14)
    GUICtrlCreateLabel("Le: 3.5 - 2 - 2.5 - 2.5 cm", 45 + $iColW + $iGap, $iY + 142, 200, 14)
    GUICtrlCreateLabel("Thut dong dau tien: 1.27cm", 45 + $iColW + $iGap, $iY + 156, 200, 14)
    
    $g_btnFixThesisVN = _CreateButton("AP DUNG CHUAN VN", 45 + $iColW + $iGap, $iY + 175, 200, 30)
    GUICtrlSetBkColor(-1, 0xE74C3C)
    GUICtrlSetFont(-1, 9, 700)
    _EndGroup()
    
    ; === COT 2: PRESET CHUAN APA ===
    _CreateGroup(" Chuan APA (Quoc te) ", 30 + $iColW + $iGap, $iY + 220, $iColW, 85)
    
    GUICtrlCreateLabel("Times New Roman 12pt, gian dong 2.0", 45 + $iColW + $iGap, $iY + 238, 200, 14)
    GUICtrlCreateLabel("Le: 1 inch (2.54cm) tat ca cac canh", 45 + $iColW + $iGap, $iY + 252, 200, 14)
    
    $g_btnFixAPA = _CreateButton("AP DUNG CHUAN APA", 45 + $iColW + $iGap, $iY + 270, 200, 28)
    GUICtrlSetBkColor(-1, 0xF39C12)
    GUICtrlSetFont(-1, 9, 600)
    _EndGroup()
    
    ; === COT 3: HUONG DAN ===
    _CreateGroup(" Huong dan Smart Fix ", 30 + ($iColW + $iGap) * 2, $iY, $iColW, 160)
    
    GUICtrlCreateLabel("CACH SU DUNG:", 45 + ($iColW + $iGap) * 2, $iY + 18, 200, 16)
    GUICtrlSetFont(-1, 8, 700)
    GUICtrlSetColor(-1, 0x2C3E50)
    
    GUICtrlCreateLabel("1. Nhan 'PHAN TICH' de kiem tra", 50 + ($iColW + $iGap) * 2, $iY + 38, 200, 14)
    GUICtrlCreateLabel("   tai lieu va xem bao cao loi", 50 + ($iColW + $iGap) * 2, $iY + 52, 200, 14)
    
    GUICtrlCreateLabel("2. Nhan 'SMART FIX' de tu dong", 50 + ($iColW + $iGap) * 2, $iY + 72, 200, 14)
    GUICtrlCreateLabel("   sua tat ca cac loi tim thay", 50 + ($iColW + $iGap) * 2, $iY + 86, 200, 14)
    
    GUICtrlCreateLabel("3. Dung 'XU LY HANG LOAT' de", 50 + ($iColW + $iGap) * 2, $iY + 106, 200, 14)
    GUICtrlCreateLabel("   xu ly nhieu file cung luc", 50 + ($iColW + $iGap) * 2, $iY + 120, 200, 14)
    
    GUICtrlCreateLabel("LUU Y: Nen BACKUP truoc khi sua!", 45 + ($iColW + $iGap) * 2, $iY + 140, 200, 14)
    GUICtrlSetFont(-1, 8, 600)
    GUICtrlSetColor(-1, 0xE74C3C)
    _EndGroup()
    
    ; === COT 3: THONG TIN THEM ===
    _CreateGroup(" Thong tin them ", 30 + ($iColW + $iGap) * 2, $iY + 165, $iColW, 100)
    
    GUICtrlCreateLabel("CHUAN LUAN VAN VN:", 45 + ($iColW + $iGap) * 2, $iY + 183, 200, 14)
    GUICtrlSetFont(-1, 8, 600)
    GUICtrlCreateLabel("Theo quy dinh cua Bo GD&DT", 50 + ($iColW + $iGap) * 2, $iY + 197, 200, 14)
    
    GUICtrlCreateLabel("CHUAN APA:", 45 + ($iColW + $iGap) * 2, $iY + 217, 200, 14)
    GUICtrlSetFont(-1, 8, 600)
    GUICtrlCreateLabel("American Psychological Association", 50 + ($iColW + $iGap) * 2, $iY + 231, 200, 14)
    GUICtrlCreateLabel("(Pho bien trong nghien cuu khoa hoc)", 50 + ($iColW + $iGap) * 2, $iY + 245, 200, 14)
    _EndGroup()
EndFunc
