; ============================================
; MAINGUI.AU3 - Tao GUI Chinh
; ============================================

#include-once

Func _CreateMainGUI()
    $g_hGUI = GUICreate($APP_TITLE & " v" & $APP_VERSION, 780, 750, -1, -1)
    GUISetIcon(_GetAppIconSource(), -1, $g_hGUI)
    GUISetBkColor(0xF0F4F8)

    ; Header
    GUICtrlCreateLabel("PDF TO WORD FIXER PRO v" & $APP_VERSION, 20, 8, 740, 30, $SS_CENTER)
    GUICtrlSetFont(-1, 18, 800, 0, "Segoe UI")
    GUICtrlSetColor(-1, 0x2C3E50)

    ; === FRAME KET NOI WORD ===
    _CreateGroup(" Ket noi Word ", 20, 42, 740, 55)
    $g_btnConnect = _CreateButton("Tu dong", 35, 60, 75, 26)
    $g_cboWordDocs = GUICtrlCreateCombo("", 118, 60, 380, 26, $CBS_DROPDOWNLIST)
    $g_btnRefresh = _CreateButton("Lam moi", 508, 60, 75, 26)
    $g_btnConnectManual = _CreateButton("Ket noi", 593, 60, 75, 26)
    $g_btnDisconnect = _CreateButton("Ngat", 678, 60, 65, 26)
    _EndGroup()

    ; Status label
    $g_lblStatus = GUICtrlCreateLabel("Nhan 'Lam moi' de xem danh sach file Word", 20, 102, 740, 18)
    GUICtrlSetColor($g_lblStatus, 0xE74C3C)
    GUICtrlSetFont(-1, 9, 600, 0, "Segoe UI")

    ; === TAB CHINH ===
    $g_hTab = GUICtrlCreateTab(20, 125, 740, 555)

    _CreateTabPDFFix()
    _CreateTabFormat()
    _CreateTabTools()
    _CreateTabTOC()
    _CreateTabCopyStyle()
    _CreateTabAdvanced()
    _CreateTabQuickUtils()
    _CreateTabSmartFix()
    _CreateTabAIFormat()

    GUICtrlCreateTabItem("")

    ; === FOOTER ===
    $g_btnHelp = _CreateButton("Tro giup", 20, 695, 90, 35)
    $g_btnBackup = _CreateButton("Backup", 120, 695, 90, 35)
    $g_btnSaveDoc = _CreateButton("Luu file", 220, 695, 90, 35)
    
    $g_lblProgress = GUICtrlCreateLabel("", 330, 705, 430, 20)
    GUICtrlSetColor($g_lblProgress, 0x27AE60)
    GUICtrlSetFont(-1, 9, 600, 0, "Segoe UI")

    GUISetState(@SW_SHOW)
EndFunc
