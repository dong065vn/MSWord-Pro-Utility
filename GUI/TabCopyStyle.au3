; ============================================
; TABCOPYSTYLE.AU3 - Tab 5: Copy Style
; ============================================

#include-once

Func _CreateTabCopyStyle()
    GUICtrlCreateTabItem("Copy Style")

    ; Huong dan
    _CreateGroup(" Huong dan Copy Style ", 35, 160, 710, 50)
    _CreateLabel("Copy dinh dang (Style, Font, Margin...) tu file nguon sang file dich.", 50, 180, 680, 20)
    _EndGroup()

    ; File nguon
    _CreateGroup(" File nguon (Copy tu) ", 35, 215, 710, 55)
    $g_cboSourceDoc = GUICtrlCreateCombo("", 50, 235, 560, 25, $CBS_DROPDOWNLIST)
    $g_btnRefreshSource = _CreateButton("Lam moi", 620, 233, 100, 28)
    _EndGroup()

    ; File dich
    _CreateGroup(" File dich (Copy sang) ", 35, 275, 710, 55)
    $g_cboTargetDoc = GUICtrlCreateCombo("", 50, 295, 560, 25, $CBS_DROPDOWNLIST)
    $g_btnRefreshTarget = _CreateButton("Lam moi", 620, 293, 100, 28)
    _EndGroup()

    ; Tuy chon copy
    _CreateGroup(" Tuy chon copy bo sung ", 35, 335, 710, 70)
    $g_chkCopyPageSetup = _CreateCheckbox(" Copy Page Setup (Le trang)", 50, 355, 200, 22, True)
    $g_chkCopyHeaderFooter = _CreateCheckbox(" Copy Header/Footer", 260, 355, 160, 22, True)
    $g_chkCopyTheme = _CreateCheckbox(" Copy Theme", 430, 355, 120, 22)
    $g_chkCopyNumbering = _CreateCheckbox(" Copy Numbering", 560, 355, 140, 22)
    $g_chkCopyStyles = _CreateCheckbox(" Copy Styles (khi dung nut Copy theo tuy chon)", 50, 378, 300, 22, True)
    _EndGroup()

    ; Buttons chinh - hang 1
    $g_btnCopyAllStyles = _CreateButton("COPY TAT CA STYLE", 35, 415, 160, 40, 0x27AE60, 10, 700)
    $g_btnSelectStyles = _CreateButton("CHON STYLE DE COPY", 205, 415, 160, 40, 0x3498DB, 10, 700)
    $g_btnCopySelectedStyles = _CreateButton("Copy theo tuy chon", 375, 415, 120, 40, -1, 9, 600)
    $g_btnPreviewStyles = _CreateButton("Xem truoc", 505, 415, 90, 40)
    
    ; Buttons hotkey - hang 2 (NEW LAYOUT)
    $g_btnApplyHotkeys = _CreateButton("Luu vao Normal.dotm", 35, 465, 140, 35, 0x27AE60, 9, 600)
    $g_btnBackupHotkeys = _CreateButton("Sao luu Hotkeys", 185, 465, 140, 35, 0x3498DB, 9, 600)
    $g_btnOpenModifyStyle = _CreateButton("Mo Modify Style", 335, 465, 140, 35, 0xF39C12, 9, 600)

    ; Huong dan su dung
    _CreateGroup(" Huong dan su dung ", 35, 510, 710, 100)
    _CreateLabel("- COPY TAT CA: Copy toan bo styles tu file nguon sang file dich", 50, 530, 650, 18)
    _CreateLabel("- CHON STYLE: Mo cua so de chon tung style can copy", 50, 548, 650, 18)
    _CreateLabel("- Copy theo tuy chon: Chi copy cac muc da tick o tren", 50, 566, 650, 18)
    Local $lbl = _CreateLabel("- Luu vao Normal.dotm: Luu phim tat vao template toan cuc | Sao luu: Backup/Restore hotkeys", 50, 584, 650, 18)
    GUICtrlSetColor($lbl, 0x27AE60)
    _EndGroup()
EndFunc
