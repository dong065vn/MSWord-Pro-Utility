; ============================================
; WORDCONNECTION.AU3 - Ket noi Word an toan
; Uu tien ObjGet, xu ly UAC issues
; ============================================

#include-once

; Lay Word Application an toan
Func _GetWordApp($bCreateIfNotExist = False)
    Local $oWord = ObjGet("", "Word.Application")
    
    If @error Or Not IsObj($oWord) Then
        Local $aProcs = ProcessList("WINWORD.EXE")
        If $aProcs[0][0] > 0 Then
            Sleep(300)
            $oWord = ObjGet("", "Word.Application")
            If @error Or Not IsObj($oWord) Then
                Return SetError(2, 0, 0) ; UAC issue
            EndIf
        ElseIf $bCreateIfNotExist Then
            $oWord = ObjCreate("Word.Application")
            If IsObj($oWord) Then $oWord.Visible = True
        Else
            Return SetError(1, 0, 0) ; Word not running
        EndIf
    EndIf
    
    Return $oWord
EndFunc

; Refresh danh sach file Word
Func _RefreshWordDocsList()
    _UpdateProgress("Dang quet file Word...")
    GUICtrlSetData($g_cboWordDocs, "")
    ReDim $g_aWordDocs[1]
    $g_aWordDocs[0] = 0

    Local $oWordTemp = _GetWordApp(False)
    Local $iErr = @error
    
    If $iErr = 2 Then
        _ShowUACWarning()
        _UpdateProgress("")
        Return
    EndIf

    If Not IsObj($oWordTemp) Then
        If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Thong bao", _
            "Khong tim thay Word!" & @CRLF & "Mo Word moi?") = $IDYES Then
            $oWordTemp = _GetWordApp(True)
            If IsObj($oWordTemp) Then
                $g_oWord = $oWordTemp
                _SetStatus("Da mo Word moi. Mo file va nhan Lam moi.", 0x3498DB)
            EndIf
        Else
            _SetStatus("Mo Word truoc roi nhan Lam moi!", 0xE74C3C)
        EndIf
        _UpdateProgress("")
        Return
    EndIf

    $g_oWord = $oWordTemp
    $g_oWord.Visible = True

    Local $oDocs = $g_oWord.Documents
    If Not IsObj($oDocs) Or $oDocs.Count = 0 Then
        _SetStatus("Word mo nhung chua co file!", 0xF39C12)
        _UpdateProgress("")
        Return
    EndIf

    Local $iCount = $oDocs.Count
    ReDim $g_aWordDocs[$iCount + 1]
    $g_aWordDocs[0] = $iCount
    Local $sDocList = ""

    For $i = 1 To $iCount
        Local $oDocTemp = $oDocs.Item($i)
        If IsObj($oDocTemp) Then
            $g_aWordDocs[$i] = $oDocTemp.FullName
            $sDocList &= ($sDocList <> "" ? "|" : "") & $i & ". " & $oDocTemp.Name
        EndIf
    Next

    GUICtrlSetData($g_cboWordDocs, $sDocList)
    If $sDocList <> "" Then GUICtrlSendMsg($g_cboWordDocs, 0x14E, 0, 0)
    _SetStatus("Tim thay " & $iCount & " file. Chon va nhan 'Ket noi'", 0x27AE60)
    _UpdateProgress("San sang")
EndFunc

; Ket noi thu cong
Func _ConnectManual()
    Local $sSelected = GUICtrlRead($g_cboWordDocs)
    If $sSelected = "" Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon file Word!")
        Return False
    EndIf

    Local $aTemp = StringSplit($sSelected, ".", 2)
    If UBound($aTemp) < 1 Then Return False
    
    Local $iIndex = Int($aTemp[0])
    If $iIndex < 1 Or $iIndex > $g_aWordDocs[0] Then Return False

    If Not IsObj($g_oWord) Then
        $g_oWord = ObjGet("", "Word.Application")
        If Not IsObj($g_oWord) Then Return False
    EndIf

    $g_oDoc = $g_oWord.Documents.Item($iIndex)
    If Not IsObj($g_oDoc) Then Return False

    $g_oDoc.Activate()
    _SetStatus("DA KET NOI: " & $g_oDoc.Name, 0x27AE60)
    _UpdateProgress("San sang!")
    Return True
EndFunc

; Ket noi tu dong
Func _ConnectToWord()
    $g_oWord = ObjGet("", "Word.Application")
    
    If @error Or Not IsObj($g_oWord) Then
        Local $aProcs = ProcessList("WINWORD.EXE")
        If $aProcs[0][0] > 0 Then
            Sleep(500)
            $g_oWord = ObjGet("", "Word.Application")
            If @error Or Not IsObj($g_oWord) Then
                _ShowUACWarning()
                Return False
            EndIf
        Else
            If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Thong bao", _
                "Khong tim thay Word!" & @CRLF & "Mo Word moi?") = $IDYES Then
                $g_oWord = ObjCreate("Word.Application")
                If IsObj($g_oWord) Then
                    $g_oWord.Visible = True
                    _SetStatus("Da mo Word moi. Mo file va nhan 'Lam moi'", 0x3498DB)
                EndIf
            EndIf
            Return False
        EndIf
    EndIf
    
    $g_oWord.Visible = True
    
    If $g_oWord.Documents.Count = 0 Then
        _SetStatus("Word mo nhung chua co file!", 0xF39C12)
        Return False
    EndIf
    
    $g_oDoc = $g_oWord.ActiveDocument
    If Not IsObj($g_oDoc) Then
        _SetStatus("Khong lay duoc ActiveDocument!", 0xE74C3C)
        Return False
    EndIf
    
    _SetStatus("DA KET NOI: " & $g_oDoc.Name, 0x27AE60)
    _UpdateProgress("San sang!")
    Return True
EndFunc

; Ngat ket noi
Func _DisconnectWord()
    $g_oDoc = 0
    _SetStatus("Da ngat ket noi", 0xF39C12)
    _UpdateProgress("")
EndFunc

; Kiem tra ket noi
Func _CheckConnection()
    If Not IsObj($g_oWord) Then
        If MsgBox($MB_YESNO, "Chua ket noi", "Ket noi Word ngay?") = $IDYES Then
            Return _ConnectToWord()
        EndIf
        Return False
    EndIf
    
    If Not IsObj($g_oDoc) Then
        If $g_oWord.Documents.Count > 0 Then
            $g_oDoc = $g_oWord.ActiveDocument
            If IsObj($g_oDoc) Then
                _SetStatus("Tu dong ket noi: " & $g_oDoc.Name, 0x27AE60)
                Return True
            EndIf
        EndIf
        If MsgBox($MB_YESNO, "Chua ket noi", "Ket noi Word ngay?") = $IDYES Then
            Return _ConnectToWord()
        EndIf
        Return False
    EndIf
    
    ; Check document still open
    Local $sName = $g_oDoc.Name
    If @error Then
        $g_oDoc = 0
        _SetStatus("File da dong! Ket noi lai.", 0xE74C3C)
        Return False
    EndIf
    
    Return True
EndFunc

; Hien thi canh bao UAC
Func _ShowUACWarning()
    MsgBox($MB_ICONWARNING, "Loi quyen truy cap", _
        "Word dang chay nhung khong ket noi duoc!" & @CRLF & @CRLF & _
        "NGUYEN NHAN:" & @CRLF & _
        "1. Tool chay 'Run as Admin' -> Chay binh thuong" & @CRLF & _
        "2. Word dang hien dialog -> Dong dialog" & @CRLF & _
        "3. Word bi treo -> Khoi dong lai Word")
    _SetStatus("Loi UAC! Xem huong dan.", 0xE74C3C)
EndFunc
