; ============================================
; CONFIG.AU3 - Constants & Global Variables
; Version 6.0 - Synced from v5.0
; ============================================

#include-once
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <StaticConstants.au3>
#include <ComboConstants.au3>
#include <TabConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <Array.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>

; === VERSION ===
Global Const $VERSION = "6.1.8"
Global Const $APP_VERSION = "6.1.8"
Global Const $APP_TITLE = "PDF to Word Fixer Pro"

; === WORD CONSTANTS ===
Global Const $CM_TO_POINTS = 28.35
Global Const $WD_REPLACE_ALL = 2
Global Const $WD_COLLAPSE_END = 0
Global Const $WD_COLLAPSE_START = 1
Global Const $WD_ALIGN_CENTER = 1
Global Const $WD_ALIGN_LEFT = 0
Global Const $WD_ALIGN_RIGHT = 2
Global Const $WD_ALIGN_JUSTIFY = 3
Global Const $WD_LINE_SPACE_MULTIPLE = 5
Global Const $WD_TAB_ALIGN_RIGHT = 2
Global Const $WD_TAB_LEADER_DOTS = 2
Global Const $WD_TAB_LEADER_DASHES = 3
Global Const $WD_TAB_LEADER_LINES = 4
Global Const $WD_EXPORT_PDF = 17
Global Const $WD_FORMAT_HTML = 8
Global Const $WD_FORMAT_TEXT = 2
Global Const $WD_FORMAT_RTF = 6

; === ORGANIZER COPY CONSTANTS ===
Global Const $wdOrganizerObjectStyles = 0
Global Const $wdFormatOriginalFormatting = 16

; === STYLE TYPE CONSTANTS ===
Global Const $wdStyleTypeParagraph = 1
Global Const $wdStyleTypeCharacter = 2

; === GLOBAL STATE ===
Global $g_oWord = 0
Global $g_oDoc = 0
Global $g_hGUI = 0
Global $g_aWordDocs[1] = [0]
Global $g_bProcessing = False

; === GUI CONTROLS - Connection ===
Global $g_btnConnect, $g_cboWordDocs, $g_btnRefresh
Global $g_btnConnectManual, $g_btnDisconnect
Global $g_lblStatus, $g_hTab, $g_editPreview, $g_lblProgress

; === Tab 1: PDF Fix ===
Global $g_chkLineBreaks, $g_chkExtraSpaces, $g_chkHyphenation, $g_chkSpecialChars
Global $g_chkParagraphs, $g_chkTabs, $g_chkVietnamese, $g_chkPageNumbers, $g_chkFixQuotes, $g_chkRemoveFakeNumbering
Global $g_chkFixLineSpacing, $g_chkResetSpacing, $g_chkRemoveEmptyLines, $g_chkFixSpacingBefore
Global $g_btnFixSelected, $g_btnFixAll, $g_btnQuickFix, $g_btnUndoFix, $g_btnCleanUp, $g_btnFixLayout, $g_btnPDFFixHelp

; === Tab 2: Format ===
Global $g_cboFont, $g_cboFontSize, $g_cboLineSpacing, $g_cboAlignment
Global $g_inputLeftMargin, $g_inputRightMargin, $g_inputTopMargin, $g_inputBottomMargin
Global $g_chkAutoFirstLine, $g_inputCaptionPrefix, $g_inputChapterNum
Global $g_btnApplyFormat, $g_btnApplySelection, $g_btnPresetVN, $g_btnPresetUS, $g_btnCheckThesis
Global $g_btnFormatH1, $g_btnFormatH2, $g_btnFormatH3
Global $g_btnFormatCaption, $g_btnFormatNormal, $g_btnClearFormat
Global $g_btnRemoveHighlight, $g_btnUnifyFont, $g_btnFixAllSpacing
Global $g_btnAddPageNum, $g_btnAddHeader, $g_btnRemovePageNum
Global $g_btnAutoNumImg, $g_btnAutoNumTbl, $g_btnAutoNumEq, $g_btnRemoveNumEq

; === Tab 3: Tools ===
Global $g_inputFind, $g_inputReplace, $g_chkMatchCase, $g_chkWholeWord
Global $g_btnFindNext, $g_btnFindReplace
Global $g_btnPreviewParentheses, $g_btnRemoveParenthesesSelection, $g_btnRemoveParenthesesText
Global $g_btnResizeImages, $g_btnCenterImages, $g_btnAutoCaption, $g_btnRemoveImages
Global $g_btnAutoFitTable, $g_btnAutoFitWindow, $g_btnTableCaption, $g_btnTableBorder
Global $g_btnWordCount, $g_btnCheckSpelling, $g_btnCheckFormat
Global $g_btnShowStats, $g_btnExportStats

; === Tab 4: TOC ===
Global $g_btnCreateTOC, $g_btnUpdateTOC, $g_btnDeleteTOC
Global $g_cboTOCLevels, $g_chkTOCHyperlink
Global $g_btnFixTOCStyles, $g_btnFixTOC1, $g_btnFixTOC2, $g_btnFixTOC3
Global $g_btnPreviewTOC, $g_cboTabLeader
Global $g_cboCitationStyle, $g_btnFormatReferences
Global $g_inputAuthor, $g_inputYear, $g_inputTitle, $g_inputSource, $g_inputURL
Global $g_btnAddReference, $g_btnInsertCitation, $g_btnSortRef, $g_btnClearRef

; === Tab 5: Copy Style ===
Global $g_cboSourceDoc, $g_cboTargetDoc, $g_btnRefreshSource, $g_btnRefreshTarget
Global $g_chkCopyStyles, $g_chkCopyPageSetup, $g_chkCopyHeaderFooter
Global $g_chkCopyTheme, $g_chkCopyNumbering
Global $g_btnCopyAllStyles, $g_btnSelectStyles, $g_btnCopySelectedStyles
Global $g_btnPreviewStyles, $g_btnApplyHotkeys, $g_btnBackupHotkeys
Global $g_btnOpenModifyStyle, $g_btnRefreshHotkeys

; === Style Selector Popup ===
Global $g_hStyleSelector = 0
Global $g_listStyles

; === Important Styles for Quick Gallery ===
Global $g_aImportantStyles[15] = ["Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4", _
    "Title", "Subtitle", "Caption", "TOC 1", "TOC 2", "TOC 3", "Quote", "List Paragraph", "No Spacing", "Strong"]

; === Tab 6: Advanced ===
Global $g_btnAutoHeading, $g_btnResetHeading, $g_btnHeadingToTOC, $g_btnListHeadings
Global $g_btnScanThesisHeadings
Global $g_btnRemoveAllFormat, $g_btnConvertCase, $g_btnRemoveHyperlinks
Global $g_btnRemoveComments, $g_btnAcceptChanges
Global $g_btnNumberToText, $g_btnBulletToText, $g_btnNumberToTextSel
Global $g_btnExportPDF, $g_btnExportHTML, $g_btnExportTXT, $g_btnExportRTF, $g_btnPrintPreview
Global $g_btnCompareDoc, $g_btnMergeDoc, $g_btnSplitDoc
Global $g_btnProtectDoc, $g_btnDocProperties, $g_btnCleanDoc

; === Footer ===
Global $g_btnHelp, $g_btnBackup, $g_btnSaveDoc

; === Tab 7: Quick Utils (NEW) ===
Global $g_btnPastePlain, $g_btnPasteKeep, $g_btnPasteMerge
Global $g_btnSelectPara, $g_btnSelectSentence, $g_btnSelectFromStart, $g_btnSelectToEnd
Global $g_btnGoToPage, $g_btnGoToNextHeading, $g_btnGoToPrevHeading, $g_btnGoToNextTable, $g_btnGoToNextImage
Global $g_btnInsertPageBreak, $g_btnInsertSectionBreak, $g_btnInsertDate, $g_btnInsertTime, $g_btnInsertHLine
Global $g_btnFontIncrease, $g_btnFontDecrease, $g_btnToggleBold, $g_btnToggleItalic, $g_btnToggleUnderline
Global $g_btnToggleSub, $g_btnToggleSuper
Global $g_btnIncreaseIndent, $g_btnDecreaseIndent, $g_btnRemoveFirstIndent, $g_btnSetFirstIndent
Global $g_btnInsertSpecialChar, $g_btnAddBookmark, $g_btnGoToBookmark, $g_btnDeleteBookmark
Global $g_btnShowDocInfo, $g_btnRemoveHighlightSel, $g_btnRemoveCommentsSel, $g_btnUnlinkFields
Global $g_btnRemoveCitationsSel, $g_btnRemoveCitationsDoc
Global $g_btnPreviewCitations, $g_cboCitationMode
Global $g_inputCitationFilter
Global $g_inputHeadingPrefixFix, $g_inputHeadingSeparatorFix, $g_btnFixHeadingNumberDots

; === Tab 8: Smart Fix (NEW) ===
Global $g_btnSmartAnalyze, $g_btnSmartFixAll
Global $g_btnFixHyphenation, $g_btnFixNonBreaking, $g_btnFixDashes
Global $g_btnBatchProcess, $g_btnFixThesisVN, $g_btnFixAPA

; === Tab 9: AI Format (NEW) ===
Global $g_radChatGPT, $g_radGemini, $g_radClaude, $g_radCopilot, $g_radAutoDetect
Global $g_chkAIScopeSelection, $g_chkAIScopeAll
Global $g_btnAIHeadings, $g_btnAIBold, $g_btnAIItalic, $g_btnAICodeBlock
Global $g_btnAIInlineCode, $g_btnAIBullets, $g_btnAINumberList, $g_btnAITables
Global $g_btnAILinks, $g_btnAICleanAllMD
Global $g_btnAIFont, $g_btnAISpacing, $g_btnAIMargins, $g_btnAIIndent
Global $g_btnAIFixSpaces, $g_btnAIFixPunctuation, $g_btnAIRemoveEmptyLines
Global $g_btnAIFixAllThesis
Global $g_btnAILaTeX, $g_btnAIRemoveEmoji, $g_btnAIFixEncoding, $g_btnAIPreview
Global $g_btnAINormalizeMath, $g_btnAISettings, $g_btnAICleanup, $g_btnAIBeautify, $g_btnAIBeautifySettings

; === COM ERROR HANDLER ===
Global $g_bMuteComErrors = False
Global $g_oMyError = ObjEvent("AutoIt.Error", "_ComErrorHandler")

Func _ComErrorHandler($oError)
    If $g_bMuteComErrors Then Return SetError(1, 0, 0)

    Local $sDescription = $oError.description
    Local $iNumber = $oError.number

    ; Word COM voi AutoIt thinh thoang van bao warning du thao tac property/setter thanh cong.
    If $iNumber = 0x80020009 And StringInStr($sDescription, "'CustomizationContext' is not a by reference property.") Then
        Return SetError(1, 0, 0)
    EndIf

    ; Word co the tam thoi tu choi lenh khi dang cap nhat UI/list/table. Retry logic o caller se xu ly.
    If $iNumber = 0x80010001 Then
        Return SetError(1, 0, 0)
    EndIf

    ConsoleWrite("COM Error: " & $oError.description & " [0x" & Hex($oError.number) & "]" & @CRLF)
    Return SetError(1, 0, 0)
EndFunc







