; ============================================
; TABAIFORMAT.AU3 - Tab 9: AI Format
; Dinh dang noi dung tu ChatGPT/Gemini/Claude/Copilot
; ============================================

#include-once

Func _CreateTabAIFormat()
    GUICtrlCreateTabItem("AI Format")

    ; === Group 1: Nguon AI ===
    _CreateGroup(" Nguon AI ", 35, 160, 710, 55)
    $g_radChatGPT = GUICtrlCreateRadio("ChatGPT", 50, 178, 80, 20)
    GUICtrlSetState(-1, $GUI_CHECKED)
    $g_radGemini = GUICtrlCreateRadio("Gemini", 140, 178, 70, 20)
    $g_radClaude = GUICtrlCreateRadio("Claude", 220, 178, 70, 20)
    $g_radCopilot = GUICtrlCreateRadio("Copilot", 300, 178, 80, 20)
    $g_radAutoDetect = GUICtrlCreateRadio("Tu dong", 390, 178, 80, 20)
    $g_chkAIScopeSelection = GUICtrlCreateCheckbox("Vung chon", 520, 178, 80, 20)
    GUICtrlSetState(-1, $GUI_CHECKED)
    $g_chkAIScopeAll = GUICtrlCreateCheckbox("Toan bo", 610, 178, 80, 20)
    _EndGroup()

    ; === Group 2: Xu ly Markdown ===
    _CreateGroup(" Xu ly Markdown ", 35, 220, 710, 100)
    $g_btnAIHeadings = _CreateButton("## -> Heading", 50, 242, 120, 28, 0x3498DB)
    $g_btnAIBold = _CreateButton("**Bold**", 178, 242, 90, 28)
    $g_btnAIItalic = _CreateButton("*Italic*", 276, 242, 90, 28)
    $g_btnAICodeBlock = _CreateButton("``` Code ```", 374, 242, 110, 28, 0x9B59B6)
    $g_btnAIInlineCode = _CreateButton("`Inline`", 492, 242, 90, 28)
    $g_btnAILinks = _CreateButton("[Link](url)", 590, 242, 100, 28)
    $g_btnAIBullets = _CreateButton("- Bullet list", 50, 278, 120, 28)
    $g_btnAINumberList = _CreateButton("1. Number list", 178, 278, 120, 28)
    $g_btnAITables = _CreateButton("| Bang |", 306, 278, 100, 28)
    $g_btnAICleanAllMD = _CreateButton("XOA TAT CA MARKDOWN", 460, 278, 230, 28, 0xE74C3C)
    _EndGroup()

    ; === Group 3: Chuan hoa do an ===
    _CreateGroup(" Chuan hoa do an dai hoc ", 35, 325, 710, 100)
    $g_btnAIFont = _CreateButton("Font TNR 13pt", 50, 347, 120, 28)
    $g_btnAISpacing = _CreateButton("Line spacing 1.5", 178, 347, 130, 28)
    $g_btnAIMargins = _CreateButton("Margins chuan", 316, 347, 120, 28)
    $g_btnAIIndent = _CreateButton("Thut dau dong", 444, 347, 120, 28)
    $g_btnAIFixSpaces = _CreateButton("Fix khoang trang", 572, 347, 120, 28)
    $g_btnAIFixPunctuation = _CreateButton("Fix dau cau VN", 50, 383, 120, 28)
    $g_btnAIRemoveEmptyLines = _CreateButton("Xoa dong trong", 178, 383, 120, 28)
    $g_btnAIFixAllThesis = _CreateButton("FIX TAT CA - Chuan do an", 370, 383, 320, 28, 0x27AE60)
    _EndGroup()

    ; === Group 4: Xu ly nang cao & Lam dep ===
    _CreateGroup(" Xu ly nang cao & Lam dep ", 35, 430, 710, 130)
    $g_btnAILaTeX = _CreateButton("LaTeX -> Equation", 50, 452, 130, 28)
    $g_btnAIRemoveEmoji = _CreateButton("Xoa Emoji", 188, 452, 90, 28)
    $g_btnAIFixEncoding = _CreateButton("Fix encoding VN", 286, 452, 120, 28)
    $g_btnAIPreview = _CreateButton("Preview", 414, 452, 80, 28)
    $g_btnAICleanup = _CreateButton("Don dep rac", 502, 452, 100, 28, 0xE67E22)
    $g_btnAIBeautify = _CreateButton("Lam dep", 610, 452, 80, 28, 0x8E44AD)
    $g_btnAINormalizeMath = _CreateButton("Chuan cong thuc", 50, 488, 130, 28, 0x16A085)
    $g_btnAIBeautifySettings = _CreateButton("Thiet dat lam dep", 190, 488, 140, 28, 0x34495E)
    $g_btnAISettings = _CreateButton("Cai dat format", 340, 488, 120, 28, 0x34495E)
    _CreateLabel("Moi nut phuc vu 1 muc dich rieng. Chuan cong thuc tach biet, FIX TAT CA chi xu ly van ban Word.", 480, 489, 240, 32)
    _EndGroup()
EndFunc
