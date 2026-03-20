; ============================================
; TABTOC.AU3 - Tab 4: Muc luc
; ============================================

#include-once

Func _CreateTabTOC()
    GUICtrlCreateTabItem("Muc luc")

    ; Muc luc tu dong
    _CreateGroup(" Muc luc tu dong ", 35, 160, 710, 80)
    $g_btnCreateTOC = _CreateButton("Tao muc luc", 50, 185, 130, 35, 0x27AE60)
    $g_btnUpdateTOC = _CreateButton("Cap nhat", 195, 185, 100, 35)
    $g_btnDeleteTOC = _CreateButton("Xoa muc luc", 310, 185, 100, 35)
    _CreateLabel("So cap:", 440, 195, 50, 20)
    $g_cboTOCLevels = _CreateCombo(495, 192, 50, 25, "1|2|3|4|5", "3")
    $g_chkTOCHyperlink = _CreateCheckbox(" Hyperlink", 560, 193, 90, 22, True)
    _EndGroup()

    ; Fix loi TOC
    _CreateGroup(" Fix loi TOC (Muc luc bi mat tab leader) ", 35, 250, 710, 90)
    $g_btnFixTOCStyles = _CreateButton("Fix TOC Styles", 50, 273, 120, 28, 0xE74C3C, 10, 600)
    $g_btnFixTOC1 = _CreateButton("Fix TOC 1", 180, 273, 80, 28)
    $g_btnFixTOC2 = _CreateButton("Fix TOC 2", 270, 273, 80, 28)
    $g_btnFixTOC3 = _CreateButton("Fix TOC 3", 360, 273, 80, 28)
    $g_btnPreviewTOC = _CreateButton("Xem truoc", 450, 273, 90, 28)
    _CreateLabel("Tab Leader:", 555, 277, 70, 20)
    $g_cboTabLeader = _CreateCombo(625, 273, 100, 25, _
        "...... (Dots)|------ (Dashes)|_____ (Line)|(Khong)", "...... (Dots)")
    _CreateLabel("Fix loi: So trang khong can phai, thieu dau cham dan.", 50, 310, 650, 18)
    _EndGroup()

    ; Tai lieu tham khao
    _CreateGroup(" Tai lieu tham khao (TLTK) ", 35, 350, 710, 130)
    _CreateLabel("Dinh dang:", 50, 373, 65, 20)
    $g_cboCitationStyle = _CreateCombo(120, 370, 120, 25, "APA 7th|IEEE|Harvard|Chicago", "APA 7th")
    $g_btnFormatReferences = _CreateButton("Dinh dang TLTK", 255, 368, 120, 28)
    _CreateLabel("Tac gia:", 400, 373, 50, 20)
    $g_inputAuthor = _CreateInput("", 455, 370, 130, 24)
    _CreateLabel("Nam:", 600, 373, 35, 20)
    $g_inputYear = _CreateInput("", 640, 370, 80, 24)
    _CreateLabel("Tieu de:", 50, 403, 50, 20)
    $g_inputTitle = _CreateInput("", 105, 400, 290, 24)
    _CreateLabel("Nguon:", 410, 403, 50, 20)
    $g_inputSource = _CreateInput("", 465, 400, 255, 24)
    _CreateLabel("URL:", 50, 433, 35, 20)
    $g_inputURL = _CreateInput("", 90, 430, 305, 24)
    $g_btnAddReference = _CreateButton("Them TLTK", 410, 428, 90, 28)
    $g_btnInsertCitation = _CreateButton("Chen trich dan", 510, 428, 95, 28)
    $g_btnSortRef = _CreateButton("A-Z", 615, 428, 45, 28)
    $g_btnClearRef = _CreateButton("Xoa", 670, 428, 50, 28)
    _EndGroup()
EndFunc
