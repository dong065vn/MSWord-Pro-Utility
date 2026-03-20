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
    
    _UpdateProgress("Dang phan tich tai lieu...")
    
    Local $sReport = "=== BAO CAO PHAN TICH TAI LIEU ===" & @CRLF & @CRLF
    Local $aIssues[1][3] ; [Issue, Count, Severity]
    Local $iIssueCount = 0
    
    ; 1. Kiem tra Manual Line Breaks (Shift+Enter)
    Local $iLineBreaks = _CountPattern("^l")
    If $iLineBreaks > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Manual Line Break (Shift+Enter)"
        $aIssues[$iIssueCount][1] = $iLineBreaks
        $aIssues[$iIssueCount][2] = "Cao"
        $iIssueCount += 1
    EndIf
    
    ; 2. Kiem tra nhieu khoang trang lien tiep
    Local $iDoubleSpaces = _CountPattern("  ")
    If $iDoubleSpaces > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Nhieu khoang trang lien tiep"
        $aIssues[$iIssueCount][1] = $iDoubleSpaces
        $aIssues[$iIssueCount][2] = "Trung binh"
        $iIssueCount += 1
    EndIf
    
    ; 3. Kiem tra dong trong thua
    Local $iEmptyLines = _CountPattern("^p^p^p")
    If $iEmptyLines > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Dong trong thua (3+ Enter)"
        $aIssues[$iIssueCount][1] = $iEmptyLines
        $aIssues[$iIssueCount][2] = "Trung binh"
        $iIssueCount += 1
    EndIf
    
    ; 4. Kiem tra Tab thua
    Local $iDoubleTabs = _CountPattern("^t^t")
    If $iDoubleTabs > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Tab thua (2+ Tab)"
        $aIssues[$iIssueCount][1] = $iDoubleTabs
        $aIssues[$iIssueCount][2] = "Thap"
        $iIssueCount += 1
    EndIf
    
    ; 5. Kiem tra dau ngoac kep cong
    Local $iSmartQuotes = _CountSmartQuotes()
    If $iSmartQuotes > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Dau ngoac kep cong (Smart Quotes)"
        $aIssues[$iIssueCount][1] = $iSmartQuotes
        $aIssues[$iIssueCount][2] = "Thap"
        $iIssueCount += 1
    EndIf
    
    ; 6. Kiem tra bang troi noi
    Local $iFloatingTables = _CountFloatingTables()
    If $iFloatingTables > 0 Then
        ReDim $aIssues[$iIssueCount + 1][3]
        $aIssues[$iIssueCount][0] = "Bang troi noi (WrapAroundText)"
        $aIssues[$iIssueCount][1] = $iFloatingTables
        $aIssues[$iIssueCount][2] = "Cao"
        $iIssueCount += 1
    EndIf
    
    ; 7. Kiem tra hinh qua lon
    Local $iOversizedImages = _CountOversizedImages()
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
    _UpdateProgress("Phan tich xong: " & $iIssueCount & " loai van de")
    
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
    
    _UpdateProgress("Dang sua loi tu dong...")
    Local $sLog = "=== KET QUA SMART FIX ===" & @CRLF & @CRLF
    
    ; 1. Fix Manual Line Breaks
    _UpdateProgress("Dang sua Manual Line Breaks...")
    Local $oFind = $g_oDoc.Content.Find
    _DoSmartReplace($oFind, "^l", " ")
    $sLog &= "[OK] Da sua Manual Line Breaks" & @CRLF
    
    ; 2. Fix nhieu khoang trang
    _UpdateProgress("Dang sua khoang trang thua...")
    For $i = 1 To 5
        _DoSmartReplace($oFind, "  ", " ")
    Next
    _DoSmartReplace($oFind, "^p ", "^p")
    _DoSmartReplace($oFind, " ^p", "^p")
    $sLog &= "[OK] Da sua khoang trang thua" & @CRLF
    
    ; 3. Fix dong trong thua
    _UpdateProgress("Dang sua dong trong thua...")
    For $i = 1 To 10
        _DoSmartReplace($oFind, "^p^p^p", "^p^p")
    Next
    $sLog &= "[OK] Da sua dong trong thua" & @CRLF
    
    ; 4. Fix Tab thua
    _UpdateProgress("Dang sua Tab thua...")
    For $i = 1 To 3
        _DoSmartReplace($oFind, "^t^t", "^t")
    Next
    $sLog &= "[OK] Da sua Tab thua" & @CRLF
    
    ; 5. Fix Smart Quotes
    _UpdateProgress("Dang sua dau ngoac kep...")
    _ReplaceSmartQuotes($oFind)
    $sLog &= "[OK] Da sua dau ngoac kep" & @CRLF
    
    ; 6. Fix bang troi noi
    _UpdateProgress("Dang sua bang troi noi...")
    Local $oTables = $g_oDoc.Tables
    Local $iFixedTables = 0
    For $i = 1 To $oTables.Count
        Local $oTbl = $oTables.Item($i)
        If IsObj($oTbl) And $oTbl.Rows.WrapAroundText = True Then
            $oTbl.Rows.WrapAroundText = False
            $oTbl.Rows.Alignment = 1 ; Center
            $iFixedTables += 1
        EndIf
    Next
    $sLog &= "[OK] Da sua " & $iFixedTables & " bang troi noi" & @CRLF
    
    ; 7. Fix hinh qua lon
    _UpdateProgress("Dang sua hinh qua lon...")
    Local $oShapes = $g_oDoc.InlineShapes
    Local $fMaxW = $g_oDoc.PageSetup.PageWidth - $g_oDoc.PageSetup.LeftMargin - $g_oDoc.PageSetup.RightMargin
    Local $iFixedImages = 0
    For $i = 1 To $oShapes.Count
        Local $oS = $oShapes.Item($i)
        If IsObj($oS) And $oS.Width > $fMaxW Then
            Local $fRatio = $fMaxW / $oS.Width
            $oS.Width = $fMaxW
            $oS.Height = $oS.Height * $fRatio
            $iFixedImages += 1
        EndIf
    Next
    $sLog &= "[OK] Da sua " & $iFixedImages & " hinh qua lon" & @CRLF
    
    $sLog &= @CRLF & "=== HOAN TAT SMART FIX ===" & @CRLF
    
    _LogPreview($sLog)
    _UpdateProgress("Smart Fix hoan tat!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", "Da sua tat ca loi tu dong!")
EndFunc

; Helper: Thuc hien replace
Func _DoSmartReplace($oFind, $sFind, $sReplace)
    If Not IsObj($oFind) Then Return
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    $oFind.Text = $sFind
    $oFind.Replacement.Text = $sReplace
    $oFind.Forward = True
    $oFind.Wrap = 1 ; wdFindContinue
    $oFind.MatchWildcards = False
    $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2) ; wdReplaceAll
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
    
    _UpdateProgress("Dang xu ly hang loat...")
    Local $sLog = "=== KET QUA XU LY HANG LOAT ===" & @CRLF & @CRLF
    Local $iSuccess = 0
    Local $iFailed = 0
    
    For $i = 1 To $aFiles[0]
        Local $sFilePath = $sFolder & "\" & $aFiles[$i]
        _UpdateProgress("Dang xu ly " & $i & "/" & $aFiles[0] & ": " & $aFiles[$i])
        
        ; Mo file
        Local $oDoc = $g_oWord.Documents.Open($sFilePath)
        If Not IsObj($oDoc) Then
            $sLog &= "[FAIL] " & $aFiles[$i] & " - Khong mo duoc" & @CRLF
            $iFailed += 1
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
    Next
    
    $sLog &= @CRLF & "TONG KET: " & $iSuccess & " thanh cong, " & $iFailed & " that bai"
    
    _LogPreview($sLog)
    _UpdateProgress("Xu ly hang loat hoan tat!")
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

