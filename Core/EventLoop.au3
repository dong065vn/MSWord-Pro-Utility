; ============================================
; EVENTLOOP.AU3 - Main Loop & Event Handling
; ============================================

#include-once

; Khoi tao ung dung
Func _InitApplication()
    _CreateMainGUI()
    _RefreshWordDocsList()
    _MainLoop()
EndFunc

; Main event loop
Func _MainLoop()
    Local $iMsg
    
    While 1
        $iMsg = GUIGetMsg()
        If $iMsg = 0 Then ContinueLoop
        If $g_bProcessing And $iMsg <> $GUI_EVENT_CLOSE Then ContinueLoop

        Switch $iMsg
            Case $GUI_EVENT_CLOSE
                ExitLoop
                
            ; === CONNECTION ===
            Case $g_btnConnect
                _SafeExecute("_ConnectToWord")
            Case $g_btnRefresh
                _SafeExecute("_RefreshWordDocsList")
            Case $g_btnConnectManual
                _SafeExecute("_ConnectManual")
            Case $g_btnDisconnect
                _DisconnectWord()
                
            ; === TAB 1: PDF FIX ===
            Case $g_btnFixSelected
                _SafeExecute("_FixSelectedText")
            Case $g_btnFixAll
                _SafeExecute("_FixAllDocument")
            Case $g_btnQuickFix
                _SafeExecute("_QuickFixAll")
            Case $g_btnCleanUp
                _SafeExecute("_CleanUpDocument")
            Case $g_btnFixLayout
                _SafeExecute("_FixLayoutProblems")
            Case $g_btnPDFFixHelp
                _ShowPDFFixHelp()
            Case $g_btnUndoFix
                _UndoAction()
                
            ; === TAB 2: FORMAT ===
            Case $g_btnApplyFormat
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _ApplyFormat(False)
                    $g_bProcessing = False
                EndIf
            Case $g_btnApplySelection
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _ApplyFormat(True)
                    $g_bProcessing = False
                EndIf
            Case $g_btnPresetVN
                _LoadPresetVN()
            Case $g_btnPresetUS
                _LoadPresetUS()
            Case $g_btnCheckThesis
                _SafeExecute("_CheckThesisFormat")
            Case $g_btnFormatH1
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _FormatHeading(1)
                    $g_bProcessing = False
                EndIf
            Case $g_btnFormatH2
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _FormatHeading(2)
                    $g_bProcessing = False
                EndIf
            Case $g_btnFormatH3
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _FormatHeading(3)
                    $g_bProcessing = False
                EndIf
            Case $g_btnFormatCaption
                _SafeExecute("_FormatCaption")
            Case $g_btnFormatNormal
                _SafeExecute("_FormatNormal")
            Case $g_btnClearFormat
                _SafeExecute("_ClearFormat")
            Case $g_btnRemoveHighlight
                _SafeExecute("_RemoveHighlight")
            Case $g_btnUnifyFont
                _SafeExecute("_UnifyFont")
            Case $g_btnFixAllSpacing
                _SafeExecute("_FixAllSpacing")
            Case $g_btnAddPageNum
                _SafeExecute("_AddPageNumbers")
            Case $g_btnAddHeader
                _SafeExecute("_AddHeader")
            Case $g_btnRemovePageNum
                _SafeExecute("_RemovePageNumbers")
            Case $g_btnAutoNumImg
                _SafeExecute("_AutoNumberImages")
            Case $g_btnAutoNumTbl
                _SafeExecute("_AutoNumberTables")
            Case $g_btnAutoNumEq
                _SafeExecute("_NumberEquations")
            Case $g_btnRemoveNumEq
                _SafeExecute("_RemoveEquationNumbers")
                
            ; === TAB 3: TOOLS ===
            Case $g_btnFindReplace
                _SafeExecute("_DoFindReplace")
            Case $g_btnFindNext
                _DoFindNext()
            Case $g_btnResizeImages
                _SafeExecute("_ResizeImages")
            Case $g_btnCenterImages
                _SafeExecute("_CenterImages")
            Case $g_btnAutoCaption
                _SafeExecute("_AutoCaptionImg")
            Case $g_btnRemoveImages
                _SafeExecute("_RemoveAllImages")
            Case $g_btnAutoFitTable
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _AutoFitTables(1)
                    $g_bProcessing = False
                EndIf
            Case $g_btnAutoFitWindow
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _AutoFitTables(2)
                    $g_bProcessing = False
                EndIf
            Case $g_btnTableCaption
                _SafeExecute("_AutoCaptionTbl")
            Case $g_btnTableBorder
                _SafeExecute("_AddTableBorders")
            Case $g_btnWordCount
                _ShowWordCount()
            Case $g_btnCheckSpelling
                _SafeExecute("_CheckSpelling")
            Case $g_btnCheckFormat
                _SafeExecute("_CheckFormat")
            Case $g_btnShowStats
                _ShowDetailedStats()
            Case $g_btnExportStats
                _SafeExecute("_ExportStats")
                
            ; === TAB 4: TOC ===
            Case $g_btnCreateTOC
                _SafeExecute("_CreateTOC")
            Case $g_btnUpdateTOC
                _SafeExecute("_UpdateTOC")
            Case $g_btnDeleteTOC
                _SafeExecute("_DeleteTOC")
            Case $g_btnFixTOCStyles
                _SafeExecute("_FixAllTOCStyles")
            Case $g_btnFixTOC1
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _FixTOCStyle(1)
                    $g_bProcessing = False
                EndIf
            Case $g_btnFixTOC2
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _FixTOCStyle(2)
                    $g_bProcessing = False
                EndIf
            Case $g_btnFixTOC3
                If Not $g_bProcessing Then
                    $g_bProcessing = True
                    _FixTOCStyle(3)
                    $g_bProcessing = False
                EndIf
            Case $g_btnPreviewTOC
                _PreviewTOCStyles()
            Case $g_btnAddReference
                _SafeExecute("_AddReference")
            Case $g_btnInsertCitation
                _SafeExecute("_InsertCitation")
            Case $g_btnSortRef
                _SafeExecute("_SortReferences")
            Case $g_btnClearRef
                _ClearRefForm()
            Case $g_btnFormatReferences
                _SafeExecute("_FormatReferences")
                
            ; === TAB 5: COPY STYLE ===
            Case $g_btnRefreshSource, $g_btnRefreshTarget
                _RefreshStyleDocsList()
            Case $g_btnCopyAllStyles
                _SafeExecute("_CopyAllStyles")
            Case $g_btnSelectStyles
                _ShowStyleSelector()
            Case $g_btnCopySelectedStyles
                _SafeExecute("_CopySelectedStyles")
            Case $g_btnPreviewStyles
                _PreviewSourceStyles()
            Case $g_btnApplyHotkeys
                _SafeExecute("_ApplyHotkeysToCurrentDoc")
            Case $g_btnBackupHotkeys
                _ShowBackupHotkeysDialog()
            Case $g_btnRefreshHotkeys
                _SafeExecute("_RefreshHotkeysFromWord")
            Case $g_btnOpenModifyStyle
                _SafeExecute("_OpenModifyStyleDialog")
                
            ; === TAB 6: ADVANCED ===
            Case $g_btnAutoHeading
                _SafeExecute("_AutoDetectHeading")
            Case $g_btnResetHeading
                _SafeExecute("_ResetAllHeadings")
            Case $g_btnHeadingToTOC
                _SafeExecute("_HeadingToTOC")
            Case $g_btnListHeadings
                _ListAllHeadings()
            Case $g_btnScanThesisHeadings
                _SafeExecute("_ScanAndApplyThesisHeadingStyles")
            Case $g_btnRemoveAllFormat
                _SafeExecute("_RemoveAllFormatting")
            Case $g_btnConvertCase
                _SafeExecute("_ConvertTextCase")
            Case $g_btnRemoveHyperlinks
                _SafeExecute("_RemoveAllHyperlinks")
            Case $g_btnRemoveComments
                _SafeExecute("_RemoveAllComments")
            Case $g_btnAcceptChanges
                _SafeExecute("_AcceptAllChanges")
            Case $g_btnNumberToText
                _SafeExecute("_ConvertNumberingToText")
            Case $g_btnBulletToText
                _SafeExecute("_ConvertBulletToText")
            Case $g_btnNumberToTextSel
                _SafeExecute("_ConvertNumberingToTextSelection")
            Case $g_btnExportPDF
                _SafeExecute("_ExportToPDF")
            Case $g_btnExportHTML
                _SafeExecute("_ExportToHTML")
            Case $g_btnExportTXT
                _SafeExecute("_ExportToTXT")
            Case $g_btnExportRTF
                _SafeExecute("_ExportToRTF")
            Case $g_btnPrintPreview
                _ShowPrintPreview()
            Case $g_btnCompareDoc
                _SafeExecute("_CompareDocuments")
            Case $g_btnMergeDoc
                _SafeExecute("_MergeDocuments")
            Case $g_btnSplitDoc
                _SafeExecute("_SplitDocument")
            Case $g_btnProtectDoc
                _SafeExecute("_ProtectDocument")
            Case $g_btnDocProperties
                _ShowDocProperties()
            Case $g_btnCleanDoc
                _SafeExecute("_CleanDocument")
                
            ; === TAB 9: AI FORMAT ===
            Case $g_btnAIHeadings
                _SafeExecute("_AI_ConvertHeadings")
            Case $g_btnAIBold
                _SafeExecute("_AI_ConvertBold")
            Case $g_btnAIItalic
                _SafeExecute("_AI_ConvertItalic")
            Case $g_btnAICodeBlock
                _SafeExecute("_AI_ConvertCodeBlocks")
            Case $g_btnAIInlineCode
                _SafeExecute("_AI_ConvertInlineCode")
            Case $g_btnAIBullets
                _SafeExecute("_AI_ConvertBullets")
            Case $g_btnAINumberList
                _SafeExecute("_AI_ConvertNumberedLists")
            Case $g_btnAITables
                _SafeExecute("_AI_ConvertTables")
            Case $g_btnAILinks
                _SafeExecute("_AI_ConvertLinks")
            Case $g_btnAICleanAllMD
                _SafeExecute("_AI_CleanAllMarkdown")
            Case $g_btnAIFont
                _SafeExecute("_AI_ApplyThesisFont")
            Case $g_btnAISpacing
                _SafeExecute("_AI_ApplyThesisSpacing")
            Case $g_btnAIMargins
                _SafeExecute("_AI_ApplyThesisMargins")
            Case $g_btnAIIndent
                _SafeExecute("_AI_ApplyFirstLineIndent")
            Case $g_btnAIFixSpaces
                _SafeExecute("_AI_FixExtraSpaces")
            Case $g_btnAIFixPunctuation
                _SafeExecute("_AI_FixVietnamesePunctuation")
            Case $g_btnAIRemoveEmptyLines
                _SafeExecute("_AI_RemoveEmptyLines")
            Case $g_btnAIFixAllThesis
                _SafeExecute("_AI_FixAllThesis")
            Case $g_btnAILaTeX
                _SafeExecute("_AI_ConvertLaTeX")
            Case $g_btnAINormalizeMath
                _SafeExecute("_AI_NormalizeAllMath")
            Case $g_btnAIRemoveEmoji
                _SafeExecute("_AI_RemoveEmoji")
            Case $g_btnAIFixEncoding
                _SafeExecute("_AI_FixEncoding")
            Case $g_btnAIPreview
                _AI_PreviewChanges()
            Case $g_btnAISettings
                _AI_ShowSettings()
            Case $g_btnAICleanup
                _SafeExecute("_AI_CleanupVisual")
            Case $g_btnAIBeautify
                _SafeExecute("_AI_BeautifyDocument")
            Case $g_btnAIBeautifySettings
                _AI_ShowBeautifySettings()
                
            ; === FOOTER ===
            Case $g_btnHelp
                _ShowHelp()
            Case $g_btnBackup
                _BackupDocument()
            Case $g_btnSaveDoc
                _SaveDocument()
                
            ; === TAB 7: QUICK UTILS ===
            Case $g_btnPastePlain
                _SafeExecute("_PasteAsPlainText")
            Case $g_btnPasteKeep
                _SafeExecute("_PasteKeepSourceFormat")
            Case $g_btnPasteMerge
                _SafeExecute("_PasteMergeFormat")
            Case $g_btnSelectPara
                _SafeExecute("_SelectCurrentParagraph")
            Case $g_btnSelectSentence
                _SafeExecute("_SelectCurrentSentence")
            Case $g_btnSelectFromStart
                _SafeExecute("_SelectFromStart")
            Case $g_btnSelectToEnd
                _SafeExecute("_SelectToEnd")
            Case $g_btnGoToPage
                _SafeExecute("_GoToPage")
            Case $g_btnGoToNextHeading
                _SafeExecute("_GoToNextHeading")
            Case $g_btnGoToPrevHeading
                _SafeExecute("_GoToPrevHeading")
            Case $g_btnGoToNextTable
                _SafeExecute("_GoToNextTable")
            Case $g_btnGoToNextImage
                _SafeExecute("_GoToNextImage")
            Case $g_btnInsertPageBreak
                _SafeExecute("_InsertPageBreak")
            Case $g_btnInsertSectionBreak
                _SafeExecute("_InsertSectionBreak")
            Case $g_btnInsertDate
                _SafeExecute("_InsertCurrentDate")
            Case $g_btnInsertTime
                _SafeExecute("_InsertCurrentTime")
            Case $g_btnInsertHLine
                _SafeExecute("_InsertHorizontalLine")
            Case $g_btnFontIncrease
                _SafeExecute("_IncreaseFontSize")
            Case $g_btnFontDecrease
                _SafeExecute("_DecreaseFontSize")
            Case $g_btnToggleBold
                _SafeExecute("_ToggleBold")
            Case $g_btnToggleItalic
                _SafeExecute("_ToggleItalic")
            Case $g_btnToggleUnderline
                _SafeExecute("_ToggleUnderline")
            Case $g_btnToggleSub
                _SafeExecute("_ToggleSubscript")
            Case $g_btnToggleSuper
                _SafeExecute("_ToggleSuperscript")
            Case $g_btnIncreaseIndent
                _SafeExecute("_IncreaseIndent")
            Case $g_btnDecreaseIndent
                _SafeExecute("_DecreaseIndent")
            Case $g_btnRemoveFirstIndent
                _SafeExecute("_RemoveFirstLineIndent")
            Case $g_btnSetFirstIndent
                _SafeExecute("_SetFirstLineIndent")
            Case $g_btnInsertSpecialChar
                _SafeExecute("_InsertSpecialChar")
            Case $g_btnAddBookmark
                _SafeExecute("_AddBookmark")
            Case $g_btnGoToBookmark
                _SafeExecute("_GoToBookmark")
            Case $g_btnDeleteBookmark
                _SafeExecute("_DeleteBookmark")
            Case $g_btnShowDocInfo
                _SafeExecute("_ShowDetailedDocInfo")
            Case $g_btnRemoveHighlightSel
                _SafeExecute("_RemoveHighlightSelection")
            Case $g_btnRemoveCommentsSel
                _SafeExecute("_RemoveCommentsSelection")
            Case $g_btnUnlinkFields
                _SafeExecute("_UnlinkAllFields")
            Case $g_btnFixHeadingNumberDots
                _SafeExecute("_FixHeadingNumberDots")
                
            ; === TAB 8: SMART FIX ===
            Case $g_btnSmartAnalyze
                _SafeExecute("_SmartAnalyzeDocument")
            Case $g_btnSmartFixAll
                _SafeExecute("_SmartFixAll")
            Case $g_btnFixHyphenation
                _SafeExecute("_FixHyphenation")
            Case $g_btnFixNonBreaking
                _SafeExecute("_FixNonBreakingSpaces")
            Case $g_btnFixDashes
                _SafeExecute("_FixDashes")
            Case $g_btnBatchProcess
                _SafeExecute("_BatchProcessFiles")
            Case $g_btnFixThesisVN
                _SafeExecute("_FixForThesisVN")
            Case $g_btnFixAPA
                _SafeExecute("_FixForAPA")
        EndSwitch
    WEnd
    
    ; Cleanup
    $g_oWord = 0
    $g_oDoc = 0
    GUIDelete()
EndFunc
