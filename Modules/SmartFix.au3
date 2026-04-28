; ============================================
; SMARTFIX.AU3 - Module Sua loi Thong minh
; Phat hien va sua loi tu dong
; Version: 6.1
; ============================================

#include-once

; ============================================
; 1. SMART ANALYSIS - Phan tich thong minh
; ============================================

; Phan tich van ban va de xuat sua loi
Func _SmartAnalyzeDocument()
    If Not _CheckConnection() Then Return
    
    _Process_Start("Smart Analyze", 7, "toan bo tai lieu")
    _Process_Step("Dang phan tich tai lieu", 0)
    
    Local $sReport = "=== BAO CAO PHAN TICH TAI LIEU ===" & @CRLF & @CRLF
    Local $aIssues[1][3] ; [Issue, Count, Severity]
    Local $iIssueCount = 0
    
    ; 1. Kiem tra Manual Line Breaks (Shift+Enter)
    Local $iLineBreaks = _CountPattern("^l")
    _Process_Step("Kiem tra Manual Line Break", 1)
    If $iLineBreaks > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Manual Line Break (Shift+Enter)"
        $aIssues[$iIssueCount][1] = $iLineBreaks
        $aIssues[$iIssueCount][2] = "Cao"
        $iIssueCount += 1
    EndIf
    
    ; 2. Kiem tra nhieu khoang trang lien tiep
    Local $iDoubleSpaces = _CountPattern("  ")
    _Process_Step("Kiem tra khoang trang", 2)
    If $iDoubleSpaces > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Nhieu khoang trang lien tiep"
        $aIssues[$iIssueCount][1] = $iDoubleSpaces
        $aIssues[$iIssueCount][2] = "Trung binh"
        $iIssueCount += 1
    EndIf
    
    ; 3. Kiem tra dong trong thua
    Local $iEmptyLines = _CountPattern("^p^p^p")
    _Process_Step("Kiem tra dong trong", 3)
    If $iEmptyLines > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Dong trong thua (3+ Enter)"
        $aIssues[$iIssueCount][1] = $iEmptyLines
        $aIssues[$iIssueCount][2] = "Trung binh"
        $iIssueCount += 1
    EndIf
    
    ; 4. Kiem tra Tab thua
    Local $iDoubleTabs = _CountPattern("^t^t")
    _Process_Step("Kiem tra tab", 4)
    If $iDoubleTabs > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Tab thua (2+ Tab)"
        $aIssues[$iIssueCount][1] = $iDoubleTabs
        $aIssues[$iIssueCount][2] = "Thap"
        $iIssueCount += 1
    EndIf
    
    ; 5. Kiem tra dau ngoac kep cong
    Local $iSmartQuotes = _CountSmartQuotes()
    _Process_Step("Kiem tra smart quotes", 5)
    If $iSmartQuotes > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Dau ngoac kep cong (Smart Quotes)"
        $aIssues[$iIssueCount][1] = $iSmartQuotes
        $aIssues[$iIssueCount][2] = "Thap"
        $iIssueCount += 1
    EndIf
    
    ; 6. Kiem tra bang troi noi
    Local $iFloatingTables = _CountFloatingTables()
    _Process_Step("Kiem tra bang", 6)
    If $iFloatingTables > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Bang troi noi (WrapAroundText)"
        $aIssues[$iIssueCount][1] = $iFloatingTables
        $aIssues[$iIssueCount][2] = "Cao"
        $iIssueCount += 1
    EndIf
    
    ; 7. Kiem tra hinh qua lon
    Local $iOversizedImages = _CountOversizedImages()
    _Process_Step("Kiem tra hinh", 7)
    If $iOversizedImages > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Hinh qua lon (tran le)"
        $aIssues[$iIssueCount][1] = $iOversizedImages
        $aIssues[$iIssueCount][2] = "Cao"
        $iIssueCount += 1
    EndIf
    
    ; Tao bao cao
    If $iIssueCount = 0 Then
        $sReport &= "KHONG TIM THAY VAN DE NAO!" & @CRLF
        $sReport &= "Tai lieu cua ban da sach!" & @CRLF
    Else
        $sReport &= "TIM THAY " & $iIssueCount & " LOAI VAN DE:" & @CRLF & @CRLF
        
        For $i = 0 To $iIssueCount - 1
            $sReport &= ($i + 1) & ". " & $aIssues[$i][0] & @CRLF
            $sReport &= "   So luong: " & $aIssues[$i][1] & @CRLF
            $sReport &= "   Muc do: " & $aIssues[$i][2] & @CRLF & @CRLF
        Next
        
        $sReport &= "DE XUAT: Nhan 'SMART FIX' de sua tat ca tu dong."
    EndIf
    
    _LogPreview($sReport)
    _Process_Done("Phan tich xong: " & $iIssueCount & " loai van de")
    
    Return $iIssueCount
EndFunc

; Dem so lan xuat hien cua pattern
Func _CountPattern($sPattern)
    If Not IsObj($g_oDoc) Then Return 0
    
    Local $oFind = $g_oDoc.Content.Find
    $oFind.ClearFormatting()
    $oFind.Text = $sPattern
    $oFind.Forward = True
    $oFind.Wrap = 0 ; wdFindStop
    $oFind.MatchWildcards = False
    
    Local $iCount = 0
    While $oFind.Execute()
        $iCount += 1
        If $iCount > 9999 Then ExitLoop ; Gioi han de tranh treo
    WEnd
    
    Return $iCount
EndFunc

; Dem dau ngoac kep cong
Func _CountSmartQuotes()
    If Not IsObj($g_oDoc) Then Return 0
    
    Local $sText = $g_oDoc.Content.Text
    Local $iCount = 0
    
    ; Dem cac loai smart quotes
    $iCount += StringLen($sText) - StringLen(StringReplace($sText, ChrW(8220), "")) ; "
    $iCount += StringLen($sText) - StringLen(StringReplace($sText, ChrW(8221), "")) ; "
    $iCount += StringLen($sText) - StringLen(StringReplace($sText, ChrW(8216), "")) ; '
    $iCount += StringLen($sText) - StringLen(StringReplace($sText, ChrW(8217), "")) ; '
    
    Return $iCount
EndFunc

; Dem bang troi noi
Func _CountFloatingTables()
    If Not IsObj($g_oDoc) Then Return 0
    
    Local $oTables = $g_oDoc.Tables
    Local $iCount = 0
    
    For $i = 1 To $oTables.Count
        Local $oTbl = $oTables.Item($i)
        If IsObj($oTbl) And $oTbl.Rows.WrapAroundText = True Then
            $iCount += 1
        EndIf
    Next
    
    Return $iCount
EndFunc

; Dem hinh qua lon
Func _CountOversizedImages()
    If Not IsObj($g_oDoc) Then Return 0
    
    Local $oShapes = $g_oDoc.InlineShapes
    Local $fMaxW = $g_oDoc.PageSetup.PageWidth - $g_oDoc.PageSetup.LeftMargin - $g_oDoc.PageSetup.RightMargin
    Local $iCount = 0
    
    For $i = 1 To $oShapes.Count
        Local $oS = $oShapes.Item($i)
        If IsObj($oS) And $oS.Width > $fMaxW Then
            $iCount += 1
        EndIf
    Next
    
    Return $iCount
EndFunc

; ============================================
; 2. SMART FIX - Sua loi tu dong
; ============================================

; Sua tat ca loi tu dong
Func _SmartFixAll()
    If Not _CheckConnection() Then Return
    
    ; Phan tich truoc
    Local $iIssues = _SmartAnalyzeDocument()
    
    If $iIssues = 0 Then
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong co van de nao can sua!")
        Return
    EndIf
    
    If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Smart Fix", _
        "Tim thay " & $iIssues & " loai van de." & @CRLF & @CRLF & _
        "Ban co muon sua tat ca tu dong?" & @CRLF & _
        "(Nen backup truoc khi sua)") <> $IDYES Then Return
    
    Local $aBatch = _Perf_BeginWordBatch("Smart Fix")
    _Process_Start("Smart Fix", 7, "toan bo tai lieu")
    _Process_Step("Dang sua loi tu dong", 0)
    Local $sLog = "=== KET QUA SMART FIX ===" & @CRLF & @CRLF
    
    ; 1. Fix Manual Line Breaks
    _UpdateProgress("Dang sua Manual Line Breaks...")
    Local $oFind = $g_oDoc.Content.Find
    _DoSmartReplace($oFind, "^l", " ")
    $sLog &= "[OK] Da sua Manual Line Breaks" & @CRLF
    _Process_Step("Manual Line Breaks", 1, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    
    ; 2. Fix nhieu khoang trang
    _UpdateProgress("Dang sua khoang trang thua...")
    For $i = 1 To 5
        _DoSmartReplace($oFind, "  ", " ")
    Next
    _DoSmartReplace($oFind, "^p ", "^p")
    _DoSmartReplace($oFind, " ^p", "^p")
    $sLog &= "[OK] Da sua khoang trang thua" & @CRLF
    _Process_Step("Khoang trang thua", 2, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    
    ; 3. Fix dong trong thua
    _UpdateProgress("Dang sua dong trong thua...")
    For $i = 1 To 10
        _DoSmartReplace($oFind, "^p^p^p", "^p^p")
    Next
    $sLog &= "[OK] Da sua dong trong thua" & @CRLF
    _Process_Step("Dong trong thua", 3, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    
    ; 4. Fix Tab thua
    _UpdateProgress("Dang sua Tab thua...")
    For $i = 1 To 3
        _DoSmartReplace($oFind, "^t^t", "^t")
    Next
    $sLog &= "[OK] Da sua Tab thua" & @CRLF
    _Process_Step("Tab thua", 4, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    
    ; 5. Fix Smart Quotes
    _UpdateProgress("Dang sua dau ngoac kep...")
    _ReplaceSmartQuotes($oFind)
    $sLog &= "[OK] Da sua dau ngoac kep" & @CRLF
    _Process_Step("Smart Quotes", 5, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    
    ; 6. Fix bang troi noi
    _UpdateProgress("Dang sua bang troi noi...")
    Local $oTables = $g_oDoc.Tables
    Local $iFixedTables = 0
    Local $iTableCount = _Perf_CollectionCount($oTables)
    For $i = 1 To $iTableCount
        Local $oTbl = $oTables.Item($i)
        If IsObj($oTbl) And $oTbl.Rows.WrapAroundText = True Then
            If Not $g_bDomDryRun Then
                $oTbl.Rows.WrapAroundText = False
                $oTbl.Rows.Alignment = 1 ; Center
            EndIf
            $iFixedTables += 1
            _Process_AddChanged()
        EndIf
    Next
    $sLog &= "[OK] Da sua " & $iFixedTables & " bang troi noi" & @CRLF
    _Process_Step("Bang troi noi", 6, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    
    ; 7. Fix hinh qua lon
    _UpdateProgress("Dang sua hinh qua lon...")
    Local $oShapes = $g_oDoc.InlineShapes
    Local $fMaxW = $g_oDoc.PageSetup.PageWidth - $g_oDoc.PageSetup.LeftMargin - $g_oDoc.PageSetup.RightMargin
    Local $iFixedImages = 0
    Local $iShapeCount = _Perf_CollectionCount($oShapes)
    For $i = 1 To $iShapeCount
        Local $oS = $oShapes.Item($i)
        If IsObj($oS) And $oS.Width > $fMaxW Then
            Local $fRatio = $fMaxW / $oS.Width
            If Not $g_bDomDryRun Then
                $oS.Width = $fMaxW
                $oS.Height = $oS.Height * $fRatio
            EndIf
            $iFixedImages += 1
            _Process_AddChanged()
        EndIf
    Next
    $sLog &= "[OK] Da sua " & $iFixedImages & " hinh qua lon" & @CRLF
    _Process_Step("Hinh qua lon", 7, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    
    $sLog &= @CRLF & "=== HOAN TAT SMART FIX ===" & @CRLF
    
    _LogPreview($sLog)
    _Perf_EndWordBatch($aBatch)
    _Process_Done("Smart Fix hoan tat: " & _Process_BuildSummary())
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da sua tat ca loi tu dong!")
EndFunc

; Helper: Thuc hien replace
Func _DoSmartReplace($oFind, $sFind, $sReplace)
    If Not IsObj($oFind) Then Return
    If $g_bDomDryRun Then
        _Process_AddSkipped()
        Return
    EndIf
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    $oFind.Text = $sFind
    $oFind.Replacement.Text = $sReplace
    $oFind.Forward = True
    $oFind.Wrap = 1 ; wdFindContinue
    $oFind.MatchWildcards = False
    Local $vResult = $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2) ; wdReplaceAll
    If @error Then
        _Process_AddError()
    ElseIf $vResult Then
        _Process_AddChanged()
    EndIf
EndFunc

; ============================================
; 3. FIX SPECIFIC ISSUES - Sua loi cu the
; ============================================

; Sua loi tu bi ngat dong (hyphenation)
Func _FixHyphenation()
    If Not _CheckConnection() Then Return
    
    _UpdateProgress("Dang sua tu bi ngat dong...")
    
    Local $oFind = $g_oDoc.Content.Find
    
    ; Pattern: dau gach noi + xuong dong
    _DoSmartReplace($oFind, "- ^p", "")
    _DoSmartReplace($oFind, "-^p", "")
    _DoSmartReplace($oFind, "- ^l", "")
    _DoSmartReplace($oFind, "-^l", "")
    
    _UpdateProgress("Da sua tu bi ngat dong!")
    _LogPreview("Da sua cac tu bi ngat dong (hyphenation)")
EndFunc

; Sua loi Non-breaking Space
Func _FixNonBreakingSpaces()
    If Not _CheckConnection() Then Return
    
    _UpdateProgress("Dang sua Non-breaking Space...")
    
    Local $oFind = $g_oDoc.Content.Find
    _DoSmartReplace($oFind, ChrW(160), " ")
    
    _UpdateProgress("Da sua Non-breaking Space!")
    _LogPreview("Da chuyen Non-breaking Space thanh Space thuong")
EndFunc

; Sua loi Em Dash va En Dash
Func _FixDashes()
    If Not _CheckConnection() Then Return
    
    _UpdateProgress("Dang sua Em Dash va En Dash...")
    
    Local $oFind = $g_oDoc.Content.Find
    _DoSmartReplace($oFind, ChrW(8212), "-") ; Em Dash
    _DoSmartReplace($oFind, ChrW(8211), "-") ; En Dash
    
    _UpdateProgress("Da sua Em Dash va En Dash!")
    _LogPreview("Da chuyen Em Dash va En Dash thanh gach ngang thuong")
EndFunc

; ============================================
; 4. BATCH OPERATIONS - Xu ly hang loat
; ============================================

; Xu ly hang loat nhieu file
Func _BatchProcessFiles()
    If Not IsObj($g_oWord) Then
        $g_oWord = ObjGet("", "Word.Application")
        If Not IsObj($g_oWord) Then
            MsgBox($MB_ICONWARNING, "Loi", "Vui long mo Word truoc!")
            Return
        EndIf
    EndIf
    
    ; Chon thu muc chua file
    Local $sFolder = FileSelectFolder("Chon thu muc chua file Word", @DesktopDir)
    If @error Or $sFolder = "" Then Return
    
    ; Tim tat ca file Word
    Local $aFiles = _FileListToArray($sFolder, "*.docx", 1)
    If @error Or Not IsArray($aFiles) Or $aFiles[0] = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong tim thay file Word nao!")
        Return
    EndIf
    
    If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Xu ly hang loat", _
        "Tim thay " & $aFiles[0] & " file Word." & @CRLF & @CRLF & _
        "Ban co muon xu ly tat ca?" & @CRLF & _
        "(Moi file se duoc Smart Fix va luu lai)") <> $IDYES Then Return
    
    _Process_Start("Batch Smart Fix", $aFiles[0], $sFolder)
    _Process_Step("Dang xu ly hang loat", 0)
    Local $sLog = "=== KET QUA XU LY HANG LOAT ===" & @CRLF & @CRLF
    Local $iSuccess = 0
    Local $iFailed = 0
    
    For $i = 1 To $aFiles[0]
        Local $sFilePath = $sFolder & "\" & $aFiles[$i]
        _Process_Step("Dang xu ly: " & $aFiles[$i], $i, $iSuccess, $iFailed, $g_iProcessErrors)
        
        ; Mo file
        Local $oDoc = $g_oWord.Documents.Open($sFilePath)
        If Not IsObj($oDoc) Then
            $sLog &= "[FAIL] " & $aFiles[$i] & " - Khong mo duoc" & @CRLF
            $iFailed += 1
            _Process_AddError()
            ContinueLoop
        EndIf
        
        ; Luu reference cu
        Local $oDocOld = $g_oDoc
        $g_oDoc = $oDoc
        
        ; Thuc hien Smart Fix (khong hien dialog)
        _BatchSmartFix()
        
        ; Luu va dong
        $oDoc.Save()
        $oDoc.Close()
        
        ; Khoi phuc reference
        $g_oDoc = $oDocOld
        
        $sLog &= "[OK] " & $aFiles[$i] & @CRLF
        $iSuccess += 1
        _Process_Step("Da xu ly: " & $aFiles[$i], $i, $iSuccess, $iFailed, $g_iProcessErrors)
    Next
    
    $sLog &= @CRLF & "TONG KET: " & $iSuccess & " thanh cong, " & $iFailed & " that bai"
    
    _LogPreview($sLog)
    _Process_Done("Xu ly hang loat: " & $iSuccess & " thanh cong, " & $iFailed & " that bai")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", _
        "Da xu ly " & $iSuccess & "/" & $aFiles[0] & " file thanh cong!")
EndFunc

; Smart Fix cho batch (khong hien dialog)
Func _BatchSmartFix()
    If Not IsObj($g_oDoc) Then Return
    
    Local $oFind = $g_oDoc.Content.Find
    
    ; Fix Manual Line Breaks
    _DoSmartReplace($oFind, "^l", " ")

    ; Fix hyphenation
    _DoSmartReplace($oFind, "- ^p", "")
    _DoSmartReplace($oFind, "-^p", "")
    _DoSmartReplace($oFind, "- ^l", "")
    _DoSmartReplace($oFind, "-^l", "")
    
    ; Fix khoang trang
    For $i = 1 To 5
        _DoSmartReplace($oFind, "  ", " ")
    Next
    _DoSmartReplace($oFind, "^p ", "^p")
    _DoSmartReplace($oFind, " ^p", "^p")
    
    ; Fix dong trong
    For $i = 1 To 10
        _DoSmartReplace($oFind, "^p^p^p", "^p^p")
    Next
    
    ; Fix Tab
    For $i = 1 To 3
        _DoSmartReplace($oFind, "^t^t", "^t")
    Next
    
    ; Fix non-breaking spaces va dashes
    _DoSmartReplace($oFind, ChrW(160), " ")
    _DoSmartReplace($oFind, ChrW(8212), "-")
    _DoSmartReplace($oFind, ChrW(8211), "-")

    ; Fix Smart Quotes
    _ReplaceSmartQuotes($oFind)
EndFunc

Func _ReplaceSmartQuotes($oFind)
    If Not IsObj($oFind) Then Return

    Local $bPrevAutoType = False, $bPrevAutoReplace = False
    Local $bHasOptions = IsObj($g_oWord) And IsObj($g_oWord.Options)

    If $bHasOptions Then
        $bPrevAutoType = $g_oWord.Options.AutoFormatAsYouTypeReplaceQuotes
        $bPrevAutoReplace = $g_oWord.Options.AutoFormatReplaceQuotes
        $g_oWord.Options.AutoFormatAsYouTypeReplaceQuotes = False
        $g_oWord.Options.AutoFormatReplaceQuotes = False
    EndIf

    _DoSmartReplace($oFind, ChrW(8220), '"')
    _DoSmartReplace($oFind, ChrW(8221), '"')
    _DoSmartReplace($oFind, ChrW(8216), "'")
    _DoSmartReplace($oFind, ChrW(8217), "'")

    If $bHasOptions Then
        $g_oWord.Options.AutoFormatAsYouTypeReplaceQuotes = $bPrevAutoType
        $g_oWord.Options.AutoFormatReplaceQuotes = $bPrevAutoReplace
    EndIf
EndFunc

; ============================================
; 4B. THUC CHIEN PRO - Safe analyze / preview / fix pipeline
; ============================================

Global Const $SMARTPRO_OPT_DRYRUN = 0
Global Const $SMARTPRO_OPT_SCOPE_SELECTION = 1
Global Const $SMARTPRO_OPT_SKIP_TABLES = 2
Global Const $SMARTPRO_OPT_SKIP_EQUATIONS = 3
Global Const $SMARTPRO_OPT_SKIP_LINKS = 4
Global Const $SMARTPRO_OPT_SKIP_FIELDS = 5
Global Const $SMARTPRO_OPT_FIX_LAYOUT = 6
Global Const $SMARTPRO_OPT_FAKE_NUMBERING = 7

Func _SmartAnalyzePro()
    If Not _CheckConnection() Then Return
    Local $aOptions = _SmartPro_ReadOptions(True)
    _SmartPro_Run($aOptions, False)
EndFunc

Func _SmartPreviewPro()
    If Not _CheckConnection() Then Return
    Local $aOptions = _SmartPro_ReadOptions(True)
    _SmartPro_Run($aOptions, False)
EndFunc

Func _SmartFixPro()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Fix Pro", _
        "Chay pipeline Thuc chien Pro voi cac option bao ve dang tick?" & @CRLF & @CRLF & _
        "Nen backup tai lieu truoc khi sua.") <> $IDYES Then Return

    Local $aOptions = _SmartPro_ReadOptions(False)
    _SmartPro_Run($aOptions, True)
EndFunc

Func _SmartPro_ReadOptions($bDryRun)
    Local $aOptions[8]
    $aOptions[$SMARTPRO_OPT_DRYRUN] = $bDryRun
    $aOptions[$SMARTPRO_OPT_SCOPE_SELECTION] = _SmartPro_ReadCheck($g_chkProScopeSelection, True)
    $aOptions[$SMARTPRO_OPT_SKIP_TABLES] = _SmartPro_ReadCheck($g_chkProSkipTables, True)
    $aOptions[$SMARTPRO_OPT_SKIP_EQUATIONS] = _SmartPro_ReadCheck($g_chkProSkipEquations, True)
    $aOptions[$SMARTPRO_OPT_SKIP_LINKS] = _SmartPro_ReadCheck($g_chkProSkipLinks, True)
    $aOptions[$SMARTPRO_OPT_SKIP_FIELDS] = _SmartPro_ReadCheck($g_chkProSkipFields, True)
    $aOptions[$SMARTPRO_OPT_FIX_LAYOUT] = _SmartPro_ReadCheck($g_chkProFixLayout, True)
    $aOptions[$SMARTPRO_OPT_FAKE_NUMBERING] = _SmartPro_ReadCheck($g_chkProFakeNumbering, False)
    Return $aOptions
EndFunc

Func _SmartPro_ReadCheck($iControl, $bDefault)
    If $iControl <= 0 Then Return $bDefault
    Return (GUICtrlRead($iControl) = $GUI_CHECKED)
EndFunc

Func _SmartPro_Run(ByRef $aOptions, $bApply)
    Local $bOldDryRun = $g_bDomDryRun
    Local $bOldPreferSelection = $g_bDomPreferSelection
    Local $bOldSkipTables = $g_bDomSkipTables
    Local $bOldSkipEquations = $g_bDomSkipEquations
    Local $bOldSkipLinks = $g_bDomSkipLinks
    Local $bOldSkipFields = $g_bDomSkipFields

    $g_bDomDryRun = (Not $bApply) Or $aOptions[$SMARTPRO_OPT_DRYRUN]
    $g_bDomPreferSelection = $aOptions[$SMARTPRO_OPT_SCOPE_SELECTION]
    $g_bDomSkipTables = $aOptions[$SMARTPRO_OPT_SKIP_TABLES]
    $g_bDomSkipEquations = $aOptions[$SMARTPRO_OPT_SKIP_EQUATIONS]
    $g_bDomSkipLinks = $aOptions[$SMARTPRO_OPT_SKIP_LINKS]
    $g_bDomSkipFields = $aOptions[$SMARTPRO_OPT_SKIP_FIELDS]

    Local $oRange = _SmartPro_GetTargetRange($aOptions[$SMARTPRO_OPT_SCOPE_SELECTION])
    If Not IsObj($oRange) Then
        MsgBox($MB_ICONWARNING, "Fix Pro", "Khong lay duoc pham vi xu ly.")
        _SmartPro_RestoreSafety($bOldDryRun, $bOldPreferSelection, $bOldSkipTables, $bOldSkipEquations, $bOldSkipLinks, $bOldSkipFields)
        Return
    EndIf

    Local $sScope = _Dom_GetScopeLabel($oRange)
    Local $aStats = _SmartPro_AnalyzeRange($oRange, $aOptions)
    Local $sReport = _SmartPro_BuildReport($aStats, $aOptions, $sScope, $bApply)

    If Not $bApply Then
        _Process_Start(($aOptions[$SMARTPRO_OPT_DRYRUN] ? "Preview Pro" : "Analyze Pro"), 1, $sScope)
        _Process_Step("Da phan tich pipeline Pro", 1)
        _LogPreview($sReport)
        _Process_Done("Preview: candidates=" & _SmartPro_TotalIssues($aStats))
        _SmartPro_RestoreSafety($bOldDryRun, $bOldPreferSelection, $bOldSkipTables, $bOldSkipEquations, $bOldSkipLinks, $bOldSkipFields)
        Return
    EndIf

    Local $aBatch = _Perf_BeginWordBatch("Fix Pro")
    _Process_Start("Fix Pro", 7, $sScope)
    _Process_Step("Dang xu ly text", 1)

    Local $iTextChanged = _SmartPro_ApplyTextFixesToParagraphs($oRange, $aOptions)
    _Process_Step("Text fixes", 2, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)

    Local $iParaChanged = _SmartPro_CollapseEmptyParagraphs($oRange)
    _Process_Step("Dong trong", 3, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)

    Local $iTableChanged = 0
    Local $iImageChanged = 0
    If $aOptions[$SMARTPRO_OPT_FIX_LAYOUT] Then
        $iTableChanged = _SmartPro_FixTables($oRange)
        _Process_Step("Table layout", 4, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
        $iImageChanged = _SmartPro_FixImages($oRange)
        _Process_Step("Image layout", 5, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    Else
        _Process_Step("Bo qua layout", 5, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    EndIf

    $sReport &= @CRLF & "=== AP DUNG ===" & @CRLF
    $sReport &= "Text paragraphs changed: " & $iTextChanged & @CRLF
    $sReport &= "Paragraph breaks changed: " & $iParaChanged & @CRLF
    $sReport &= "Tables fixed: " & $iTableChanged & @CRLF
    $sReport &= "Images fixed: " & $iImageChanged & @CRLF
    $sReport &= "Summary: " & _Process_BuildSummary() & @CRLF

    _Process_Step("Hoan tat pipeline", 7, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)
    _LogPreview($sReport)
    _Perf_EndWordBatch($aBatch)
    _Process_Done("Fix Pro: " & _Process_BuildSummary())
    MsgBox($MB_ICONINFORMATION, "Fix Pro", "Da chay xong pipeline Thuc chien Pro." & @CRLF & _Process_BuildSummary())

    _SmartPro_RestoreSafety($bOldDryRun, $bOldPreferSelection, $bOldSkipTables, $bOldSkipEquations, $bOldSkipLinks, $bOldSkipFields)
EndFunc

Func _SmartPro_RestoreSafety($bDryRun, $bPreferSelection, $bSkipTables, $bSkipEquations, $bSkipLinks, $bSkipFields)
    $g_bDomDryRun = $bDryRun
    $g_bDomPreferSelection = $bPreferSelection
    $g_bDomSkipTables = $bSkipTables
    $g_bDomSkipEquations = $bSkipEquations
    $g_bDomSkipLinks = $bSkipLinks
    $g_bDomSkipFields = $bSkipFields
EndFunc

Func _SmartPro_GetTargetRange($bPreferSelection)
    If $bPreferSelection Then
        Local $oSelectionRange = _Dom_GetSelectionRange()
        If IsObj($oSelectionRange) Then Return $oSelectionRange
    EndIf
    Return _Dom_GetScopeRange("document")
EndFunc

Func _SmartPro_AnalyzeRange($oRange, ByRef $aOptions)
    Local $aStats[10]
    If Not IsObj($oRange) Then Return $aStats

    Local $sText = $oRange.Text
    $aStats[0] = _SmartPro_CountString($sText, Chr(11))
    $aStats[1] = _SmartPro_CountString($sText, "  ")
    $aStats[2] = _SmartPro_CountString($sText, @CR & @CR & @CR)
    $aStats[3] = _SmartPro_CountString($sText, @TAB & @TAB)
    $aStats[4] = _SmartPro_CountSmartQuotesInText($sText)
    $aStats[5] = _SmartPro_CountString($sText, ChrW(160))
    $aStats[6] = _SmartPro_CountDashesInText($sText)
    $aStats[7] = _SmartPro_CountFakeNumbering($oRange, $aOptions)
    $aStats[8] = _SmartPro_CountFloatingTablesInRange($oRange)
    $aStats[9] = _SmartPro_CountOversizedImagesInRange($oRange)
    Return $aStats
EndFunc

Func _SmartPro_BuildReport(ByRef $aStats, ByRef $aOptions, $sScope, $bApply)
    Local $sReport = "=== THUC CHIEN PRO ===" & @CRLF
    $sReport &= "Mode: " & ($bApply ? "FIX" : "PREVIEW / DRY-RUN") & @CRLF
    $sReport &= "Scope: " & $sScope & @CRLF
    $sReport &= "Skip: tables=" & $aOptions[$SMARTPRO_OPT_SKIP_TABLES] & _
        ", equations=" & $aOptions[$SMARTPRO_OPT_SKIP_EQUATIONS] & _
        ", links=" & $aOptions[$SMARTPRO_OPT_SKIP_LINKS] & _
        ", fields=" & $aOptions[$SMARTPRO_OPT_SKIP_FIELDS] & @CRLF & @CRLF

    $sReport &= "Manual line breaks: " & $aStats[0] & @CRLF
    $sReport &= "Double spaces: " & $aStats[1] & @CRLF
    $sReport &= "Extra empty lines: " & $aStats[2] & @CRLF
    $sReport &= "Double tabs: " & $aStats[3] & @CRLF
    $sReport &= "Smart quotes: " & $aStats[4] & @CRLF
    $sReport &= "Non-breaking spaces: " & $aStats[5] & @CRLF
    $sReport &= "En/Em dashes: " & $aStats[6] & @CRLF
    $sReport &= "Fake numbering candidates: " & $aStats[7] & @CRLF
    $sReport &= "Floating tables: " & $aStats[8] & @CRLF
    $sReport &= "Oversized images: " & $aStats[9] & @CRLF
    $sReport &= "Total candidates: " & _SmartPro_TotalIssues($aStats) & @CRLF
    Return $sReport
EndFunc

Func _SmartPro_TotalIssues(ByRef $aStats)
    Local $iTotal = 0
    For $i = 0 To UBound($aStats) - 1
        $iTotal += $aStats[$i]
    Next
    Return $iTotal
EndFunc

Func _SmartPro_CountString($sText, $sNeedle)
    If $sNeedle = "" Then Return 0
    Local $iCount = 0
    Local $iPos = 1
    While 1
        $iPos = StringInStr($sText, $sNeedle, 0, 1, $iPos)
        If $iPos = 0 Then ExitLoop
        $iCount += 1
        $iPos += StringLen($sNeedle)
    WEnd
    Return $iCount
EndFunc

Func _SmartPro_CountSmartQuotesInText($sText)
    Return _SmartPro_CountString($sText, ChrW(8220)) + _
        _SmartPro_CountString($sText, ChrW(8221)) + _
        _SmartPro_CountString($sText, ChrW(8216)) + _
        _SmartPro_CountString($sText, ChrW(8217))
EndFunc

Func _SmartPro_CountDashesInText($sText)
    Return _SmartPro_CountString($sText, ChrW(8211)) + _SmartPro_CountString($sText, ChrW(8212))
EndFunc

Func _SmartPro_CountFakeNumbering($oRange, ByRef $aOptions)
    If Not $aOptions[$SMARTPRO_OPT_FAKE_NUMBERING] Then Return 0
    If Not IsObj($oRange) Then Return 0
    Local $oParas = $oRange.Paragraphs
    If Not IsObj($oParas) Then Return 0

    Local $iCount = 0
    Local $iParaCount = _Perf_CollectionCount($oParas)
    For $i = 1 To $iParaCount
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        If _Dom_ShouldSkipRange($oPara.Range) Then ContinueLoop
        If $oPara.Range.ListFormat.ListType <> 0 Then ContinueLoop
        Local $sOriginal = $oPara.Range.Text
        If _StripUnnecessaryLeadingNumbering($sOriginal) <> $sOriginal Then $iCount += 1
    Next
    Return $iCount
EndFunc

Func _SmartPro_ApplyTextFixesToParagraphs($oRange, ByRef $aOptions)
    If Not IsObj($oRange) Then Return 0
    Local $oParas = $oRange.Paragraphs
    If Not IsObj($oParas) Then Return 0

    Local $iChanged = 0
    Local $iParaTotal = _Perf_CollectionCount($oParas)
    For $i = 1 To $iParaTotal
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        If _Perf_ShouldRenderStep($i, $iParaTotal, 50) Then _Process_Step("Text paragraph " & $i & "/" & $iParaTotal, 1, $g_iProcessChanged, $g_iProcessSkipped, $g_iProcessErrors)

        Local $oParaRange = $oPara.Range
        If _Dom_ShouldSkipRange($oParaRange) Then
            _Process_AddSkipped()
            ContinueLoop
        EndIf

        Local $sOriginal = $oParaRange.Text
        Local $sTail = _SmartPro_GetRangeTail($sOriginal)
        Local $sCore = StringLeft($sOriginal, StringLen($sOriginal) - StringLen($sTail))
        Local $sFixed = _SmartPro_FixText($sCore, $aOptions) & $sTail

        If $sFixed = $sOriginal Then ContinueLoop
        If _Dom_SetRangeText($oParaRange, $sFixed) Then
            $iChanged += 1
            _Process_AddChanged()
        EndIf
    Next
    Return $iChanged
EndFunc

Func _SmartPro_GetRangeTail($sText)
    Local $sTail = ""
    While StringLen($sText) > 0
        Local $sLast = StringRight($sText, 1)
        If $sLast <> @CR And $sLast <> Chr(7) Then ExitLoop
        $sTail = $sLast & $sTail
        $sText = StringTrimRight($sText, 1)
    WEnd
    Return $sTail
EndFunc

Func _SmartPro_FixText($sText, ByRef $aOptions)
    $sText = StringReplace($sText, Chr(11), " ")
    $sText = StringRegExpReplace($sText, "-\s*[\r\n]+\s*", "")
    $sText = StringReplace($sText, "- " & Chr(11), "")
    $sText = StringReplace($sText, "-" & Chr(11), "")
    $sText = StringReplace($sText, ChrW(160), " ")
    $sText = StringReplace($sText, ChrW(8220), '"')
    $sText = StringReplace($sText, ChrW(8221), '"')
    $sText = StringReplace($sText, ChrW(8216), "'")
    $sText = StringReplace($sText, ChrW(8217), "'")
    $sText = StringReplace($sText, ChrW(8211), "-")
    $sText = StringReplace($sText, ChrW(8212), "-")

    For $i = 1 To 8
        If Not StringInStr($sText, "  ") Then ExitLoop
        $sText = StringReplace($sText, "  ", " ")
    Next
    For $i = 1 To 4
        If Not StringInStr($sText, @TAB & @TAB) Then ExitLoop
        $sText = StringReplace($sText, @TAB & @TAB, @TAB)
    Next

    If $aOptions[$SMARTPRO_OPT_FAKE_NUMBERING] Then
        $sText = _StripUnnecessaryLeadingNumbering($sText)
    EndIf

    Return $sText
EndFunc

Func _SmartPro_CollapseEmptyParagraphs($oRange)
    If Not IsObj($oRange) Then Return 0
    If $g_bDomSkipTables Or $g_bDomSkipEquations Or $g_bDomSkipLinks Or $g_bDomSkipFields Then
        _Process_AddSkipped()
        Return 0
    EndIf

    Local $iChanged = 0
    For $i = 1 To 8
        Local $iResult = _Dom_SafeFindReplace($oRange, "^p^p^p", "^p^p")
        If $iResult Then $iChanged += 1
    Next
    Return $iChanged
EndFunc

Func _SmartPro_CountFloatingTablesInRange($oRange)
    If Not IsObj($oRange) Then Return 0
    Local $oTables = $oRange.Tables
    If Not IsObj($oTables) Then Return 0
    Local $iCount = 0
    Local $iTableTotal = _Perf_CollectionCount($oTables)
    For $i = 1 To $iTableTotal
        Local $oTbl = $oTables.Item($i)
        If IsObj($oTbl) And $oTbl.Rows.WrapAroundText = True Then $iCount += 1
    Next
    Return $iCount
EndFunc

Func _SmartPro_FixTables($oRange)
    If Not IsObj($oRange) Then Return 0
    If $g_bDomSkipTables Then
        _Process_AddSkipped()
        Return 0
    EndIf

    Local $oTables = $oRange.Tables
    If Not IsObj($oTables) Then Return 0
    Local $iFixed = 0
    Local $iTableCount = _Perf_CollectionCount($oTables)
    For $i = 1 To $iTableCount
        Local $oTbl = $oTables.Item($i)
        If Not IsObj($oTbl) Then ContinueLoop
        If $oTbl.Rows.WrapAroundText = True Then
            $oTbl.Rows.WrapAroundText = False
            $oTbl.Rows.Alignment = 1
            If @error Then
                _Process_AddError()
            Else
                _Process_AddChanged()
                $iFixed += 1
            EndIf
        EndIf
    Next
    Return $iFixed
EndFunc

Func _SmartPro_CountOversizedImagesInRange($oRange)
    If Not IsObj($oRange) Then Return 0
    Local $oShapes = $oRange.InlineShapes
    If Not IsObj($oShapes) Then Return 0
    Local $fMaxW = _SmartPro_GetTextWidth()
    Local $iCount = 0
    Local $iShapeTotal = _Perf_CollectionCount($oShapes)
    For $i = 1 To $iShapeTotal
        Local $oShape = $oShapes.Item($i)
        If Not IsObj($oShape) Then ContinueLoop
        If $g_bDomSkipEquations And _Dom_IsMathInlineShape($oShape) Then ContinueLoop
        If $oShape.Width > $fMaxW Then $iCount += 1
    Next
    Return $iCount
EndFunc

Func _SmartPro_FixImages($oRange)
    If Not IsObj($oRange) Then Return 0
    Local $oShapes = $oRange.InlineShapes
    If Not IsObj($oShapes) Then Return 0
    Local $fMaxW = _SmartPro_GetTextWidth()
    Local $iFixed = 0

    Local $iShapeCount = _Perf_CollectionCount($oShapes)
    For $i = 1 To $iShapeCount
        Local $oShape = $oShapes.Item($i)
        If Not IsObj($oShape) Then ContinueLoop
        If $g_bDomSkipEquations And _Dom_IsMathInlineShape($oShape) Then
            _Process_AddSkipped()
            ContinueLoop
        EndIf
        If $oShape.Width <= $fMaxW Then ContinueLoop

        Local $fRatio = $fMaxW / $oShape.Width
        $oShape.Width = $fMaxW
        $oShape.Height = $oShape.Height * $fRatio
        If @error Then
            _Process_AddError()
        Else
            _Process_AddChanged()
            $iFixed += 1
        EndIf
    Next
    Return $iFixed
EndFunc

Func _SmartPro_GetTextWidth()
    If Not IsObj($g_oDoc) Then Return 468
    Local $fWidth = $g_oDoc.PageSetup.PageWidth - $g_oDoc.PageSetup.LeftMargin - $g_oDoc.PageSetup.RightMargin
    If @error Or $fWidth <= 0 Then Return 468
    Return $fWidth
EndFunc

; ============================================
; 5. PRESET FIXES - Sua theo mau
; ============================================

; Sua theo chuan luan van VN
Func _FixForThesisVN()
    If Not _CheckConnection() Then Return
    
    If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Chuan luan van VN", _
        "Se thuc hien:" & @CRLF & _
        "1. Smart Fix (sua loi co ban)" & @CRLF & _
        "2. Thong nhat font Times New Roman 13pt" & @CRLF & _
        "3. Gian dong 1.5" & @CRLF & _
        "4. Le: 3.5 - 2 - 2.5 - 2.5 cm" & @CRLF & _
        "5. Thut dong dau tien 1.27cm" & @CRLF & @CRLF & _
        "Tiep tuc?") <> $IDYES Then Return
    
    _UpdateProgress("Dang ap dung chuan luan van VN...")
    
    ; 1. Smart Fix
    _BatchSmartFix()
    
    ; 2. Thong nhat font
    $g_oDoc.Content.Font.Name = "Times New Roman"
    $g_oDoc.Content.Font.Size = 13
    
    ; 3. Gian dong 1.5
    $g_oDoc.Content.ParagraphFormat.LineSpacingRule = $WD_LINE_SPACE_MULTIPLE
    $g_oDoc.Content.ParagraphFormat.LineSpacing = 18 ; 12 * 1.5
    
    ; 4. Le trang
    $g_oDoc.PageSetup.LeftMargin = 3.5 * $CM_TO_POINTS
    $g_oDoc.PageSetup.RightMargin = 2 * $CM_TO_POINTS
    $g_oDoc.PageSetup.TopMargin = 2.5 * $CM_TO_POINTS
    $g_oDoc.PageSetup.BottomMargin = 2.5 * $CM_TO_POINTS
    
    ; 5. Thut dong dau tien
    $g_oDoc.Content.ParagraphFormat.FirstLineIndent = 1.27 * $CM_TO_POINTS
    
    _UpdateProgress("Da ap dung chuan luan van VN!")
    _LogPreview("Da ap dung chuan luan van VN:" & @CRLF & _
        "- Font: Times New Roman 13pt" & @CRLF & _
        "- Gian dong: 1.5" & @CRLF & _
        "- Le: 3.5 - 2 - 2.5 - 2.5 cm" & @CRLF & _
        "- Thut dong dau: 1.27cm")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da ap dung chuan luan van VN!")
EndFunc

; Sua theo chuan APA
Func _FixForAPA()
    If Not _CheckConnection() Then Return
    
    If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Chuan APA", _
        "Se thuc hien:" & @CRLF & _
        "1. Smart Fix (sua loi co ban)" & @CRLF & _
        "2. Thong nhat font Times New Roman 12pt" & @CRLF & _
        "3. Gian dong 2.0" & @CRLF & _
        "4. Le: 1 inch (2.54cm) tat ca cac canh" & @CRLF & _
        "5. Thut dong dau tien 0.5 inch" & @CRLF & @CRLF & _
        "Tiep tuc?") <> $IDYES Then Return
    
    _UpdateProgress("Dang ap dung chuan APA...")
    
    ; 1. Smart Fix
    _BatchSmartFix()
    
    ; 2. Thong nhat font
    $g_oDoc.Content.Font.Name = "Times New Roman"
    $g_oDoc.Content.Font.Size = 12
    
    ; 3. Gian dong 2.0
    $g_oDoc.Content.ParagraphFormat.LineSpacingRule = $WD_LINE_SPACE_MULTIPLE
    $g_oDoc.Content.ParagraphFormat.LineSpacing = 24 ; 12 * 2.0
    
    ; 4. Le trang 1 inch
    $g_oDoc.PageSetup.LeftMargin = 2.54 * $CM_TO_POINTS
    $g_oDoc.PageSetup.RightMargin = 2.54 * $CM_TO_POINTS
    $g_oDoc.PageSetup.TopMargin = 2.54 * $CM_TO_POINTS
    $g_oDoc.PageSetup.BottomMargin = 2.54 * $CM_TO_POINTS
    
    ; 5. Thut dong dau tien 0.5 inch
    $g_oDoc.Content.ParagraphFormat.FirstLineIndent = 1.27 * $CM_TO_POINTS
    
    _UpdateProgress("Da ap dung chuan APA!")
    _LogPreview("Da ap dung chuan APA:" & @CRLF & _
        "- Font: Times New Roman 12pt" & @CRLF & _
        "- Gian dong: 2.0" & @CRLF & _
        "- Le: 1 inch (2.54cm)" & @CRLF & _
        "- Thut dong dau: 0.5 inch")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da ap dung chuan APA!")
EndFunc

