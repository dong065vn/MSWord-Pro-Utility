; ============================================
; STYLEHOTKEY.AU3 - Module quan ly phim tat cho Style
; Chuc nang: Gan, luu, khoi phuc, ap dung hotkey cho Word Styles
; Dua tren: tailieuhuongdan.txt - Su dung Word COM API
; Version: 1.0
; ============================================

#include-once
#include "..\Shared\Helpers.au3"

; ============================================
; CORE HOTKEY FUNCTIONS
; ============================================

; === FUNCTION: Luu hotkey vao Normal.dotm (theo tailieuhuongdan.txt) ===
; Tham so:
;   $sStyleName - Ten style can gan hotkey
;   $sHotkey - Chuoi hotkey (VD: "Ctrl+1", "Alt+H")
; Tra ve: True neu thanh cong, False neu that bai
; QUAN TRONG: Ham nay luu hotkey vao Normal.dotm, khong phai document hien tai
Func _SaveHotkeysToNormalDotm($sStyleName, $sHotkey)
    ; Lay Word Application object
    Local $oWord = ObjGet("", "Word.Application")
    If @error Or Not IsObj($oWord) Then
        ConsoleWrite("! Loi: Khong the ket noi Word Application" & @CRLF)
        Return False
    EndIf
    
    ConsoleWrite(@CRLF & "=== BAT DAU LUU HOTKEY VAO NORMAL.DOTM ===" & @CRLF)
    ConsoleWrite("  Style: " & $sStyleName & @CRLF)
    ConsoleWrite("  Hotkey: " & $sHotkey & @CRLF)
    
    ; [QUAN TRONG 1] Bat che do chinh sua Normal.dotm
    ; Neu thieu dong nay, no chi luu vao file docx hien tai
    _SetCustomizationContextSafe($oWord, $oWord.NormalTemplate)
    ConsoleWrite("  [1] Da set CustomizationContext = Normal.dotm" & @CRLF)
    
    ; [QUAN TRONG 2] Kiem tra xem Style da co trong Normal.dotm chua?
    ; Nhieu khi style "0LV1" chi co o file hien tai, trong Normal chua co -> Gan se loi.
    Local $bStyleExists = False
    Local $oStyle = 0
    Local $bPrevMute = _SetComErrorsMuted(True)
    For $i = 1 To $oWord.NormalTemplate.Styles.Count
        Local $oTempStyle = $oWord.NormalTemplate.Styles.Item($i)
        If IsObj($oTempStyle) Then
            If _StyleMatchesName($oTempStyle, $sStyleName) Then
                $oStyle = $oTempStyle
                $bStyleExists = True
                ConsoleWrite("  [2] Style da ton tai trong Normal.dotm" & @CRLF)
                ExitLoop
            EndIf
        EndIf
    Next
    
    ; Neu Style chua co trong Normal, ta phai copy tu file hien tai sang
    If Not $bStyleExists Then
        ConsoleWrite("  [2] Style chua co trong Normal.dotm, dang copy..." & @CRLF)
        
        ; Lay ActiveDocument
        Local $oActiveDoc = $oWord.ActiveDocument
        If Not IsObj($oActiveDoc) Then
            ConsoleWrite("! Loi: Khong co document nao dang mo" & @CRLF)
            Return False
        EndIf
        
        ; Kiem tra style co trong ActiveDocument khong
        Local $bStyleInDoc = False
        $bPrevMute = _SetComErrorsMuted(True)
        For $i = 1 To $oActiveDoc.Styles.Count
            Local $oDocStyle = $oActiveDoc.Styles.Item($i)
            If IsObj($oDocStyle) Then
                If _StyleMatchesName($oDocStyle, $sStyleName) Then
                    $bStyleInDoc = True
                    ConsoleWrite("  [2a] Tim thay style '" & $sStyleName & "' trong document hien tai" & @CRLF)
                    ExitLoop
                EndIf
            EndIf
        Next
        _SetComErrorsMuted($bPrevMute)
        
        If Not $bStyleInDoc Then
            ConsoleWrite("! LOI QUAN TRONG: Style '" & $sStyleName & "' KHONG TON TAI trong document hien tai!" & @CRLF)
            ConsoleWrite("! => Khong the copy sang Normal.dotm" & @CRLF)
            ConsoleWrite("! => GIAI PHAP: Mo file co style nay, hoac xoa hotkey nay khoi file INI" & @CRLF)
            Return False
        EndIf
        
        ; Kiem tra xem document da duoc luu chua (can co FullName)
        Local $sDocPath = _EnsureDocumentHasPath($oActiveDoc)
        If $sDocPath = "" Or Not FileExists($sDocPath) Then
            ConsoleWrite("! LOI: Khong the tao duong dan hop le cho document truoc khi copy style." & @CRLF)
            Return False
        EndIf
        
        ; Copy style tu ActiveDocument sang NormalTemplate
        ; OrganizerCopy(Source, Dest, Name, ObjectType=wdOrganizerObjectStyles)
        ; wdOrganizerObjectStyles = 0
        Local Const $wdOrganizerObjectStyles = 0
        
        ConsoleWrite("  [2c] Dang copy style tu: " & $sDocPath & @CRLF)
        ConsoleWrite("  [2c] Sang: " & $oWord.NormalTemplate.FullName & @CRLF)
        
        ; Copy style - Thu 3 lan neu that bai
        Local $bCopySuccess = False
        For $iRetry = 1 To 3
            If $iRetry > 1 Then
                ConsoleWrite("  [2c] Thu lai lan " & $iRetry & "..." & @CRLF)
                Sleep(500) ; Doi 0.5s truoc khi thu lai
            EndIf
            
            ; Copy style
            Local $oError = ObjEvent("AutoIt.Error")
            $oWord.OrganizerCopy($sDocPath, $oWord.NormalTemplate.FullName, $sStyleName, $wdOrganizerObjectStyles)
            
            If Not @error Then
                $bCopySuccess = True
                ConsoleWrite("  [2c] Da copy style sang Normal.dotm thanh cong" & @CRLF)
                ExitLoop
            Else
                ConsoleWrite("  [2c] Lan " & $iRetry & " that bai (COM Error: " & @error & ")" & @CRLF)
            EndIf
        Next
        
        If Not $bCopySuccess Then
            ConsoleWrite("! LOI: Khong the copy style sang Normal.dotm sau 3 lan thu!" & @CRLF)
            ConsoleWrite("! => Co the style '" & $sStyleName & "' khong hop le hoac bi loi" & @CRLF)
            Return False
        EndIf
    EndIf
    
    ; Parse hotkey string to Word key codes
    Local $aKeys = _ParseHotkeyToWordKeys($sHotkey)
    If UBound($aKeys) < 4 Then
        ConsoleWrite("! Loi: Parse hotkey that bai: " & $sHotkey & @CRLF)
        Return False
    EndIf
    
    ; Kiem tra neu tat ca keys deu la wdNoKey thi loi
    If $aKeys[0] = 255 And $aKeys[1] = 255 And $aKeys[2] = 255 And $aKeys[3] = 255 Then
        ConsoleWrite("! Loi: Khong co key nao duoc parse tu: " & $sHotkey & @CRLF)
        Return False
    EndIf
    
    ; Build key code
    Local $iKeyCode = _BuildWordKeyCode($oWord, $aKeys)
    If @error Then
        ConsoleWrite("! Loi: BuildKeyCode that bai cho hotkey: " & $sHotkey & @CRLF)
        Return False
    EndIf
    
    ConsoleWrite("  [3] BuildKeyCode thanh cong: " & $iKeyCode & @CRLF)
    
    ; Xoa hotkey cu neu co (trong Normal.dotm)
    Local $oKeyBindings = $oWord.KeyBindings
    If IsObj($oKeyBindings) Then
        Local $iDeleted = 0
        For $i = $oKeyBindings.Count To 1 Step -1
            Local $oBinding = $oKeyBindings.Item($i)
            If IsObj($oBinding) Then
                ; Xoa neu trung key code hoac trung style name
                If $oBinding.KeyCode = $iKeyCode Or _
                   ($oBinding.KeyCategory = 5 And ($oBinding.Command = $sStyleName)) Then
                    $oBinding.Clear()
                    $iDeleted += 1
                EndIf
            EndIf
        Next
        If $iDeleted > 0 Then
            ConsoleWrite("  [3] Xoa " & $iDeleted & " hotkey cu" & @CRLF)
        EndIf
    EndIf
    
    ; [QUAN TRONG 3] Gan phim tat (KeyBinding)
    ; Tham so 1: 5 = wdKeyCategoryStyle (Gan cho Style)
    ; Tham so 2: Ten Style
    ; Tham so 3: Ma phim (Da tinh toan qua BuildKeyCode)
    Local Const $wdKeyCategoryStyle = 5
    
    $oWord.KeyBindings.Add($wdKeyCategoryStyle, $sStyleName, $iKeyCode)
    
    If @error Then
        ConsoleWrite("! Loi: KeyBindings.Add that bai" & @CRLF)
        Return False
    EndIf
    
    ConsoleWrite("  [4] Da gan hotkey thanh cong" & @CRLF)
    
    ; [QUAN TRONG 4] Luu de Normal.dotm
    ; Neu khong Save, tat Word di la mat het
    $oWord.NormalTemplate.Save()
    
    If @error Then
        ConsoleWrite("! Loi: Khong the luu Normal.dotm" & @CRLF)
        Return False
    EndIf
    
    ConsoleWrite("  [5] Da luu Normal.dotm thanh cong" & @CRLF)
    
    ; Verify: Kiem tra lai xem hotkey da duoc gan chua
    Local $oKeysBound = $oWord.KeysBoundTo($wdKeyCategoryStyle, $sStyleName)
    If IsObj($oKeysBound) And $oKeysBound.Count > 0 Then
        ConsoleWrite("  [6] VERIFY THANH CONG: Tim thay " & $oKeysBound.Count & " hotkey(s) cho style" & @CRLF)
        ConsoleWrite("=== HOAN TAT LUU HOTKEY VAO NORMAL.DOTM ===" & @CRLF & @CRLF)
        Return True
    Else
        ConsoleWrite("  ! CANH BAO: Khong tim thay hotkey sau khi luu" & @CRLF)
        ConsoleWrite("=== LUU HOTKEY THAT BAI ===" & @CRLF & @CRLF)
        Return False
    EndIf
EndFunc

; === FUNCTION: Ap dung hotkey cho style trong Word (DEPRECATED - Su dung _SaveHotkeysToNormalDotm) ===
; Tham so:
;   $oDoc - Document object
;   $sStyleName - Ten style can gan hotkey
;   $sHotkey - Chuoi hotkey (VD: "Ctrl+1", "Alt+H")
; Tra ve: True neu thanh cong, False neu that bai
; NOTE: Ham nay da bi thay the boi _SaveHotkeysToNormalDotm()
Func _ApplyStyleHotkeyViaWord($oDoc, $sStyleName, $sHotkey)
    If Not IsObj($oDoc) Or $sStyleName = "" Or $sHotkey = "" Then Return False

    Local $oWord = $oDoc.Application
    If Not IsObj($oWord) Then Return False

    ; Kiem tra style co ton tai (tim theo ca NameLocal va Name)
    Local $oStyle = 0
    Local $sStyleNameToUse = ""
    For $i = 1 To $oDoc.Styles.Count
        Local $oTempStyle = $oDoc.Styles.Item($i)
        If IsObj($oTempStyle) Then
            If _GetStyleNameSafe($oTempStyle, True) = $sStyleName Then
                $oStyle = $oTempStyle
                $sStyleNameToUse = _GetStyleNameSafe($oTempStyle, True)
                ExitLoop
            ElseIf _GetStyleNameSafe($oTempStyle, False) = $sStyleName Then
                $oStyle = $oTempStyle
                $sStyleNameToUse = _GetStyleNameSafe($oTempStyle, False)
                ExitLoop
            EndIf
        EndIf
    Next
    
    If Not IsObj($oStyle) Then
        ConsoleWrite("! Loi: Khong tim thay style '" & $sStyleName & "' trong file dich." & @CRLF)
        Return False
    EndIf
    
    ConsoleWrite("  Tim thay style: '" & $sStyleNameToUse & "'" & @CRLF)

    ; Parse hotkey string to Word key codes
    Local $aKeys = _ParseHotkeyToWordKeys($sHotkey)
    If UBound($aKeys) < 4 Then
        ConsoleWrite("! Loi: Parse hotkey that bai: " & $sHotkey & @CRLF)
        Return False
    EndIf
    
    ; Kiem tra neu tat ca keys deu la wdNoKey thi loi
    If $aKeys[0] = 255 And $aKeys[1] = 255 And $aKeys[2] = 255 And $aKeys[3] = 255 Then
        ConsoleWrite("! Loi: Khong co key nao duoc parse tu: " & $sHotkey & @CRLF)
        Return False
    EndIf

    ; Build key code
    Local $iKeyCode = _BuildWordKeyCode($oWord, $aKeys)
    If @error Then
        ConsoleWrite("! Loi: BuildKeyCode that bai cho hotkey: " & $sHotkey & @CRLF)
        ConsoleWrite("  Keys: [" & $aKeys[0] & ", " & $aKeys[1] & ", " & $aKeys[2] & ", " & $aKeys[3] & "]" & @CRLF)
        Return False
    EndIf
    
    ConsoleWrite("  BuildKeyCode thanh cong: " & $iKeyCode & @CRLF)

    ; Xoa hotkey cu neu co (trong Normal.dotm)
    Local $oKeyBindings = $oWord.KeyBindings
    If IsObj($oKeyBindings) Then
        Local $iDeleted = 0
        For $i = $oKeyBindings.Count To 1 Step -1
            Local $oBinding = $oKeyBindings.Item($i)
            If IsObj($oBinding) Then
                ; Xoa neu trung key code hoac trung style name
                If $oBinding.KeyCode = $iKeyCode Or _
                   ($oBinding.KeyCategory = 5 And ($oBinding.Command = $sStyleNameToUse Or $oBinding.Command = $sStyleName)) Then
                    $oBinding.Clear()
                    $iDeleted += 1
                EndIf
            EndIf
        Next
        If $iDeleted > 0 Then
            ConsoleWrite("  Xoa " & $iDeleted & " hotkey cu" & @CRLF)
        EndIf
    EndIf

    ; Them hotkey moi - wdKeyCategoryStyle = 5
    Local Const $wdKeyCategoryStyle = 5
    
    ; QUAN TRONG: Set CustomizationContext = Normal.dotm (template toan cuc)
    ; De hotkey hoat dong voi MOI document co style nay
    _SetCustomizationContextSafe($oWord, $oWord.NormalTemplate)
    ConsoleWrite("  CustomizationContext = Normal.dotm (template toan cuc)" & @CRLF)
    
    ; Gan hotkey vao Normal.dotm
    Local $bSuccess = False
    $oWord.KeyBindings.Add($wdKeyCategoryStyle, $sStyleNameToUse, $iKeyCode)
    
    If Not @error Then
        $bSuccess = True
        ConsoleWrite("+ Thanh cong: Gan hotkey " & $sHotkey & " cho style '" & $sStyleNameToUse & "' vao Normal.dotm" & @CRLF)
        
        ; Verify: Kiem tra lai xem hotkey da duoc gan chua
        Local $oKeysBound = $oWord.KeysBoundTo($wdKeyCategoryStyle, $sStyleNameToUse)
        If IsObj($oKeysBound) And $oKeysBound.Count > 0 Then
            ConsoleWrite("  Verify: Tim thay " & $oKeysBound.Count & " hotkey(s) cho style nay" & @CRLF)
            For $k = 1 To $oKeysBound.Count
                Local $oKey = $oKeysBound.Item($k)
                If IsObj($oKey) Then
                    ConsoleWrite("    - KeyString: " & $oKey.KeyString & @CRLF)
                    ConsoleWrite("    - Context: " & $oKey.Context & @CRLF)
                EndIf
            Next
        Else
            ConsoleWrite("  ! Canh bao: Khong tim thay hotkey sau khi gan" & @CRLF)
        EndIf
    Else
        ConsoleWrite("! Loi: KeyBindings.Add that bai" & @CRLF)
        $bSuccess = False
    EndIf
    
    ; Luu Normal.dotm de persist hotkey
    If $bSuccess Then
        $oWord.NormalTemplate.Save()
        ConsoleWrite("  Da luu Normal.dotm de persist hotkey" & @CRLF)
        
        ; Verify lan 2 sau khi Save
        Local $oKeysBound2 = $oWord.KeysBoundTo($wdKeyCategoryStyle, $sStyleNameToUse)
        If IsObj($oKeysBound2) And $oKeysBound2.Count > 0 Then
            ConsoleWrite("  Verify sau Save: Hotkey da duoc luu vao Normal.dotm thanh cong!" & @CRLF)
            ConsoleWrite("  => Hotkey se hoat dong voi MOI document co style '" & $sStyleNameToUse & "'" & @CRLF)
        Else
            ConsoleWrite("  ! Loi: Hotkey bi mat sau khi Save Normal.dotm!" & @CRLF)
            $bSuccess = False
        EndIf
    EndIf

    Return $bSuccess
EndFunc

; === FUNCTION: Tao KeyCode hop le cho Word tu mang key da parse ===
Func _BuildWordKeyCode($oWord, ByRef $aKeys)
    If Not IsObj($oWord) Then Return SetError(1, 0, 0)

    Local $aUsed[1] = [0]
    For $i = 0 To UBound($aKeys) - 1
        If $aKeys[$i] <> 255 Then
            ReDim $aUsed[$aUsed[0] + 2]
            $aUsed[0] += 1
            $aUsed[$aUsed[0]] = $aKeys[$i]
        EndIf
    Next

    If $aUsed[0] = 0 Then Return SetError(2, 0, 0)

    Switch $aUsed[0]
        Case 1
            Return $oWord.BuildKeyCode($aUsed[1])
        Case 2
            Return $oWord.BuildKeyCode($aUsed[1], $aUsed[2])
        Case 3
            Return $oWord.BuildKeyCode($aUsed[1], $aUsed[2], $aUsed[3])
        Case Else
            Return $oWord.BuildKeyCode($aUsed[1], $aUsed[2], $aUsed[3], $aUsed[4])
    EndSwitch
EndFunc

Func _SetCustomizationContextSafe($oWord, $oContext)
    If Not IsObj($oWord) Or Not IsObj($oContext) Then Return False

    Local $bPrevMute = False
    If IsDeclared("g_bMuteComErrors") Then
        $bPrevMute = $g_bMuteComErrors
        $g_bMuteComErrors = True
    EndIf

    $oWord.CustomizationContext = $oContext
    Local $bOk = (Not @error)

    If IsDeclared("g_bMuteComErrors") Then
        $g_bMuteComErrors = $bPrevMute
    EndIf

    Return $bOk
EndFunc

Func _GetStyleNameSafe($oStyle, $bPreferLocal = True)
    If Not IsObj($oStyle) Then Return ""

    Local $bPrevMute = False
    If IsDeclared("g_bMuteComErrors") Then
        $bPrevMute = $g_bMuteComErrors
        $g_bMuteComErrors = True
    EndIf

    Local $sLocalName = ""
    Local $sName = ""

    $sLocalName = $oStyle.NameLocal
    $sName = $oStyle.Name

    If IsDeclared("g_bMuteComErrors") Then
        $g_bMuteComErrors = $bPrevMute
    EndIf

    If $bPreferLocal Then
        If $sLocalName <> "" Then Return $sLocalName
        Return $sName
    EndIf

    If $sName <> "" Then Return $sName
    Return $sLocalName
EndFunc

Func _StyleMatchesName($oStyle, $sStyleName)
    If Not IsObj($oStyle) Then Return False

    Local $sLocalName = _GetStyleNameSafe($oStyle, True)
    Local $sName = _GetStyleNameSafe($oStyle, False)

    Return ($sLocalName = $sStyleName Or $sName = $sStyleName)
EndFunc

Func _SetComErrorsMuted($bMuted)
    Local $bPrevMute = False

    If IsDeclared("g_bMuteComErrors") Then
        $bPrevMute = $g_bMuteComErrors
        $g_bMuteComErrors = $bMuted
    EndIf

    Return $bPrevMute
EndFunc

; === FUNCTION: Dam bao document co duong dan de OrganizerCopy doc duoc ===
Func _EnsureDocumentHasPath($oDoc)
    If Not IsObj($oDoc) Then Return ""

    Local $sDocPath = ""
    If IsObj($oDoc) Then $sDocPath = $oDoc.FullName

    If $sDocPath <> "" And FileExists($sDocPath) Then
        ConsoleWrite("  [2b] Document da duoc luu: " & $sDocPath & @CRLF)
        $oDoc.Save()
        ConsoleWrite("  [2b] Da Save document de dam bao du lieu moi nhat" & @CRLF)
        Return $sDocPath
    EndIf

    ConsoleWrite("! CANH BAO: Document chua duoc luu (Save)" & @CRLF)
    ConsoleWrite("  => Dang tao file tam de thuc hien OrganizerCopy..." & @CRLF)

    Local $sTempDir = @TempDir & "\PDFToWordFixer"
    If Not FileExists($sTempDir) Then DirCreate($sTempDir)

    $sDocPath = $sTempDir & "\HotkeyTemp_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".docx"
    Local Const $wdFormatXMLDocument = 12
    $oDoc.SaveAs($sDocPath, $wdFormatXMLDocument)

    If @error Or $sDocPath = "" Or Not FileExists($sDocPath) Then
        ConsoleWrite("! LOI: Khong the SaveAs document tam tai: " & $sDocPath & @CRLF)
        Return ""
    EndIf

    ConsoleWrite("  => Da tao document tam thanh cong: " & $sDocPath & @CRLF)
    Return $sDocPath
EndFunc

; === FUNCTION: Ap dung tat ca hotkeys da luu vao Normal.dotm ===
; NOTE: Ham nay da duoc cap nhat de luu vao Normal.dotm thay vi document hien tai
Func _ApplyAllSavedHotkeys($oDoc = 0)
    ; Khong can $oDoc nua vi luu vao Normal.dotm
    ; Tham so $oDoc giu lai de backward compatibility
    
    Local $sIniFile = _GetHotkeyIniPath()
    If Not FileExists($sIniFile) Then Return 0

    Local $aHotkeys = IniReadSection($sIniFile, "Hotkeys")
    If @error Or Not IsArray($aHotkeys) Then Return 0

    Local $iApplied = 0
    For $i = 1 To $aHotkeys[0][0]
        Local $sStyleName = $aHotkeys[$i][0]
        Local $sHotkey = $aHotkeys[$i][1]
        
        ; Su dung ham moi: _SaveHotkeysToNormalDotm
        If _SaveHotkeysToNormalDotm($sStyleName, $sHotkey) Then
            $iApplied += 1
        EndIf
    Next

    Return $iApplied
EndFunc

; ============================================
; HOTKEY PARSING & VALIDATION
; ============================================

; === FUNCTION: Chuyen doi hotkey string sang mang cac Word key constants ===
Func _ParseHotkeyToWordKeys($sHotkey)
    ; Word Key Constants
    Local Const $wdKeyControl = 512
    Local Const $wdKeyAlt = 1024
    Local Const $wdKeyShift = 256
    Local Const $wdNoKey = 255

    Local $aResult[4] = [$wdNoKey, $wdNoKey, $wdNoKey, $wdNoKey]
    Local $iIndex = 0

    ConsoleWrite("  ParseHotkey: Input = '" & $sHotkey & "'" & @CRLF)

    ; Parse modifiers
    Local $sOriginal = $sHotkey
    $sHotkey = StringUpper($sHotkey)
    If StringInStr($sHotkey, "CTRL+") Then
        $aResult[$iIndex] = $wdKeyControl
        $iIndex += 1
        $sHotkey = StringReplace($sHotkey, "CTRL+", "")
        ConsoleWrite("  - Modifier: Ctrl" & @CRLF)
    EndIf
    If StringInStr($sHotkey, "ALT+") Then
        $aResult[$iIndex] = $wdKeyAlt
        $iIndex += 1
        $sHotkey = StringReplace($sHotkey, "ALT+", "")
        ConsoleWrite("  - Modifier: Alt" & @CRLF)
    EndIf
    If StringInStr($sHotkey, "SHIFT+") Then
        $aResult[$iIndex] = $wdKeyShift
        $iIndex += 1
        $sHotkey = StringReplace($sHotkey, "SHIFT+", "")
        ConsoleWrite("  - Modifier: Shift" & @CRLF)
    EndIf

    ; Parse main key
    Local $iMainKey = 0
    $sHotkey = StringStripWS($sHotkey, 3)

    ; Function keys F1-F12 (Word key codes: wdKeyF1=112 to wdKeyF12=123)
    If StringRegExp($sHotkey, "^F([1-9]|1[0-2])$") Then
        Local $iFNum = Int(StringReplace($sHotkey, "F", ""))
        $iMainKey = 111 + $iFNum ; F1=112, F2=113, ..., F12=123
        ConsoleWrite("  - Main key: " & $sHotkey & " = " & $iMainKey & @CRLF)
    ; Number keys 0-9 (ASCII codes)
    ElseIf StringRegExp($sHotkey, "^[0-9]$") Then
        $iMainKey = Asc($sHotkey)
        ConsoleWrite("  - Main key: " & $sHotkey & " = " & $iMainKey & @CRLF)
    ; Letter keys A-Z (ASCII codes)
    ElseIf StringRegExp($sHotkey, "^[A-Z]$") Then
        $iMainKey = Asc($sHotkey)
        ConsoleWrite("  - Main key: " & $sHotkey & " = " & $iMainKey & @CRLF)
    ; Special keys
    ElseIf $sHotkey = "`" Then
        $iMainKey = 0xC0
        ConsoleWrite("  - Main key: Backtick = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "-" Then
        $iMainKey = 0xBD
        ConsoleWrite("  - Main key: Minus = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "=" Then
        $iMainKey = 0xBB
        ConsoleWrite("  - Main key: Equal = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "[" Then
        $iMainKey = 0xDB
        ConsoleWrite("  - Main key: [ = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "]" Then
        $iMainKey = 0xDD
        ConsoleWrite("  - Main key: ] = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "\" Then
        $iMainKey = 0xDC
        ConsoleWrite("  - Main key: \ = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = ";" Then
        $iMainKey = 0xBA
        ConsoleWrite("  - Main key: ; = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "'" Then
        $iMainKey = 0xDE
        ConsoleWrite("  - Main key: ' = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "," Then
        $iMainKey = 0xBC
        ConsoleWrite("  - Main key: , = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "." Then
        $iMainKey = 0xBE
        ConsoleWrite("  - Main key: . = " & $iMainKey & @CRLF)
    ElseIf $sHotkey = "/" Then
        $iMainKey = 0xBF
        ConsoleWrite("  - Main key: / = " & $iMainKey & @CRLF)
    Else
        ConsoleWrite("  ! Loi: Khong nhan dien duoc main key: '" & $sHotkey & "'" & @CRLF)
        Return $aResult
    EndIf

    $aResult[$iIndex] = $iMainKey
    ConsoleWrite("  ParseHotkey: Result = [" & $aResult[0] & ", " & $aResult[1] & ", " & $aResult[2] & ", " & $aResult[3] & "]" & @CRLF)
    Return $aResult
EndFunc

; === FUNCTION: Validate dinh dang phim tat ===
Func _ValidateHotkeyFormat($sHotkey)
    ; Chap nhan cac dinh dang: Ctrl+X, Alt+X, Shift+X, Ctrl+Shift+X, Ctrl+Alt+X, etc.
    ; X co the la: 0-9, A-Z, F1-F12, hoac cac ky tu dac biet ` - = [ ] \ ; ' , . /
    
    ; Kiem tra co it nhat 1 modifier (Ctrl, Alt, Shift)
    If Not StringRegExp($sHotkey, "(?i)(Ctrl|Alt|Shift)") Then Return False
    
    ; Kiem tra dinh dang tong the
    ; Pattern: (Modifier+)+ MainKey
    Local $sPattern = "(?i)^(Ctrl\+|Alt\+|Shift\+)+(([A-Z]|[0-9]|F[1-9]|F1[0-2])|[`\-=\[\];',\./\\])$"
    Return StringRegExp($sHotkey, $sPattern)
EndFunc

; ============================================
; WORD KEYBINDING UTILITIES
; ============================================

; === FUNCTION: Convert Word KeyCode to String ===
Func _ConvertWordKeyToString($iKeyCode, $iKeyCode2 = 0)
    If $iKeyCode = 0 Then Return ""
    
    ; Word Key Constants
    Local Const $wdKeyControl = 512
    Local Const $wdKeyAlt = 1024
    Local Const $wdKeyShift = 256
    
    Local $sResult = ""
    Local $sModifiers = ""
    
    ; Kiem tra modifiers
    If BitAND($iKeyCode, $wdKeyControl) Then $sModifiers &= "Ctrl+"
    If BitAND($iKeyCode, $wdKeyAlt) Then $sModifiers &= "Alt+"
    If BitAND($iKeyCode, $wdKeyShift) Then $sModifiers &= "Shift+"
    
    ; Lay main key (remove modifiers)
    Local $iMainKey = BitAND($iKeyCode, 255)
    Local $sMainKey = ""
    
    ; Convert main key to string
    If $iMainKey >= 65 And $iMainKey <= 90 Then ; A-Z
        $sMainKey = Chr($iMainKey)
    ElseIf $iMainKey >= 48 And $iMainKey <= 57 Then ; 0-9
        $sMainKey = Chr($iMainKey)
    ElseIf $iMainKey >= 112 And $iMainKey <= 123 Then ; F1-F12
        $sMainKey = "F" & ($iMainKey - 111)
    Else
        ; Special keys
        Switch $iMainKey
            Case 192 ; VK_OEM_3
                $sMainKey = "`"
            Case 189 ; VK_OEM_MINUS
                $sMainKey = "-"
            Case 187 ; VK_OEM_PLUS
                $sMainKey = "="
            Case 219 ; VK_OEM_4
                $sMainKey = "["
            Case 221 ; VK_OEM_6
                $sMainKey = "]"
            Case 220 ; VK_OEM_5
                $sMainKey = "\"
            Case 186 ; VK_OEM_1
                $sMainKey = ";"
            Case 222 ; VK_OEM_7
                $sMainKey = "'"
            Case 188 ; VK_OEM_COMMA
                $sMainKey = ","
            Case 190 ; VK_OEM_PERIOD
                $sMainKey = "."
            Case 191 ; VK_OEM_2
                $sMainKey = "/"
            Case Else
                $sMainKey = "Key" & $iMainKey ; Fallback
        EndSwitch
    EndIf
    
    Return $sModifiers & $sMainKey
EndFunc

; === FUNCTION: Refresh Hotkey List from Word ===
Func _RefreshHotkeyListFromWord($oDoc, $hListView, ByRef $aAllStyles, $iValidCount, $sIniFile)
    If Not IsObj($oDoc) Or $hListView = 0 Then Return False
    
    Local $oWord = $oDoc.Application
    If Not IsObj($oWord) Then Return False
    
    ; Duyet qua tat ca KeyBindings trong Word
    Local $oKeyBindings = $oWord.KeyBindings
    If Not IsObj($oKeyBindings) Then Return False
    
    ; Tao mang luu hotkey tu Word
    Local $aWordHotkeys[1][2] ; [StyleName, HotkeyString]
    Local $iHotkeyCount = 0
    
    For $i = 1 To $oKeyBindings.Count
        Local $oBinding = $oKeyBindings.Item($i)
        If IsObj($oBinding) Then
            ; Kiem tra neu la style binding (wdKeyCategoryStyle = 5)
            If $oBinding.KeyCategory = 5 Then
                Local $sCommand = $oBinding.Command
                Local $sKeyString = _ConvertWordKeyToString($oBinding.KeyCode, $oBinding.KeyCode2)
                
                If $sCommand <> "" And $sKeyString <> "" Then
                    ReDim $aWordHotkeys[$iHotkeyCount + 1][2]
                    $aWordHotkeys[$iHotkeyCount][0] = $sCommand
                    $aWordHotkeys[$iHotkeyCount][1] = $sKeyString
                    $iHotkeyCount += 1
                EndIf
            EndIf
        EndIf
    Next
    
    ; Cap nhat ListView va INI file
    For $j = 0 To $iValidCount - 1
        Local $sStyleName = $aAllStyles[$j][0]
        Local $sNewHotkey = ""
        
        ; Tim hotkey cho style nay trong Word
        For $k = 0 To $iHotkeyCount - 1
            If $aWordHotkeys[$k][0] = $sStyleName Then
                $sNewHotkey = $aWordHotkeys[$k][1]
                ExitLoop
            EndIf
        Next
        
        ; Cap nhat ListView
        _GUICtrlListView_SetItemText($hListView, $j, $sNewHotkey, 3)
        $aAllStyles[$j][3] = $sNewHotkey
        
        ; Cap nhat INI file
        If $sNewHotkey <> "" Then
            IniWrite($sIniFile, "Hotkeys", $sStyleName, $sNewHotkey)
        Else
            IniDelete($sIniFile, "Hotkeys", $sStyleName)
        EndIf
    Next
    
    Return True
EndFunc

; === FUNCTION: Open Word's Modify Style Dialog ===
Func _OpenWordModifyStyleDialog($oDoc, $sStyleName)
    If Not IsObj($oDoc) Or $sStyleName = "" Then Return False
    
    Local $oWord = $oDoc.Application
    If Not IsObj($oWord) Then Return False
    
    ; Tim style trong document
    Local $oStyle = 0
    For $i = 1 To $oDoc.Styles.Count
        Local $oTempStyle = $oDoc.Styles.Item($i)
        If IsObj($oTempStyle) And _StyleMatchesName($oTempStyle, $sStyleName) Then
            $oStyle = $oTempStyle
            ExitLoop
        EndIf
    Next
    
    If Not IsObj($oStyle) Then Return False
    
    ; Activate Word window
    $oWord.Activate()
    $oDoc.Activate()
    
    ; Su dung Word's built-in command de mo Modify Style dialog
    ; wdDialogFormatStyle = 120
    Local Const $wdDialogFormatStyle = 120
    
    ; Set style name trong dialog
    Local $oDialog = $oWord.Dialogs.Item($wdDialogFormatStyle)
    If IsObj($oDialog) Then
        ; Set style name parameter
        $oDialog.Name = $sStyleName
        
        ; Show dialog (non-modal)
        $oDialog.Show()
        Return True
    EndIf
    
    Return False
EndFunc

; ============================================
; INI FILE MANAGEMENT
; ============================================

; === FUNCTION: Luu hotkey vao INI file ===
Func _SaveHotkeyToIni($sStyleName, $sHotkey, $sIniFile = "")
    If $sIniFile = "" Then $sIniFile = _GetHotkeyIniPath()
    
    If $sHotkey <> "" Then
        IniWrite($sIniFile, "Hotkeys", $sStyleName, $sHotkey)
        Return True
    Else
        IniDelete($sIniFile, "Hotkeys", $sStyleName)
        Return True
    EndIf
    
    Return False
EndFunc

; === FUNCTION: Doc hotkey tu INI file ===
Func _LoadHotkeyFromIni($sStyleName, $sIniFile = "")
    If $sIniFile = "" Then $sIniFile = _GetHotkeyIniPath()
    
    Return IniRead($sIniFile, "Hotkeys", $sStyleName, "")
EndFunc

; === FUNCTION: Doc tat ca hotkeys tu INI file ===
Func _LoadAllHotkeysFromIni($sIniFile = "")
    If $sIniFile = "" Then $sIniFile = _GetHotkeyIniPath()
    
    If Not FileExists($sIniFile) Then Return 0
    
    Local $aHotkeys = IniReadSection($sIniFile, "Hotkeys")
    If @error Or Not IsArray($aHotkeys) Then Return 0
    
    Return $aHotkeys
EndFunc


; ============================================
; BACKUP & RESTORE FUNCTIONS
; ============================================

; === FUNCTION: Backup Hotkeys Now ===
Func _BackupHotkeysNow($sIniFile = "", $sBackupDir = "")
    If $sIniFile = "" Then $sIniFile = _GetHotkeyIniPath()
    If $sBackupDir = "" Then $sBackupDir = _GetHotkeyBackupDir()
    
    If Not FileExists($sIniFile) Then
        ; Neu chua co file, tao file rong
        IniWrite($sIniFile, "Hotkeys", "_placeholder", "")
        IniDelete($sIniFile, "Hotkeys", "_placeholder")
    EndIf
    
    ; Tao thu muc backup neu chua co
    If Not FileExists($sBackupDir) Then DirCreate($sBackupDir)
    
    ; Tao ten file backup voi timestamp
    Local $sTimestamp = @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC
    Local $sBackupName = "StyleHotkeys_" & $sTimestamp & ".ini"
    Local $sBackupPath = $sBackupDir & "\" & $sBackupName
    
    ; Copy file
    If FileCopy($sIniFile, $sBackupPath, 1) Then
        Return $sBackupName
    EndIf
    
    Return ""
EndFunc

; === FUNCTION: Restore Hotkeys from Backup ===
Func _RestoreHotkeysFromBackup($sBackupPath, $sIniFile = "")
    If $sIniFile = "" Then $sIniFile = _GetHotkeyIniPath()
    
    If Not FileExists($sBackupPath) Then Return False
    
    ; Backup file hien tai truoc khi ghi de (auto-backup)
    Local $sAutoBackupDir = _GetHotkeyBackupDir()
    If Not FileExists($sAutoBackupDir) Then DirCreate($sAutoBackupDir)
    
    If FileExists($sIniFile) Then
        Local $sAutoBackup = $sAutoBackupDir & "\StyleHotkeys_BeforeRestore_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".ini"
        FileCopy($sIniFile, $sAutoBackup, 1)
    EndIf
    
    ; Ghi de file hien tai bang backup
    Return FileCopy($sBackupPath, $sIniFile, 1)
EndFunc

; === FUNCTION: Export Hotkeys to Text (for sharing) ===
Func _ExportHotkeysToText($sIniFile = "")
    If $sIniFile = "" Then $sIniFile = _GetHotkeyIniPath()
    
    If Not FileExists($sIniFile) Then
        MsgBox($MB_ICONWARNING, "Chua co hotkey", "Chua co phim tat nao duoc luu!")
        Return ""
    EndIf
    
    Local $aHotkeys = IniReadSection($sIniFile, "Hotkeys")
    If @error Or Not IsArray($aHotkeys) Or $aHotkeys[0][0] = 0 Then
        MsgBox($MB_ICONWARNING, "Trong", "File hotkeys khong co du lieu!")
        Return ""
    EndIf
    
    Local $sExport = "=== DANH SACH PHIM TAT STYLE ===" & @CRLF
    $sExport &= "Xuat ngay: " & @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN & @CRLF
    $sExport &= "Tong so: " & $aHotkeys[0][0] & " hotkeys" & @CRLF & @CRLF
    
    For $i = 1 To $aHotkeys[0][0]
        $sExport &= $aHotkeys[$i][1] & " = " & $aHotkeys[$i][0] & @CRLF
    Next
    
    ; Luu ra file text
    Local $sExportPath = @ScriptDir & "\HotkeyExport_" & @YEAR & @MON & @MDAY & ".txt"
    Local $hFile = FileOpen($sExportPath, 2 + 128) ; UTF-8
    If $hFile <> -1 Then
        FileWrite($hFile, $sExport)
        FileClose($hFile)
        Return $sExportPath
    EndIf
    
    Return ""
EndFunc

; ============================================
; KEYBOARD INPUT CAPTURE
; ============================================

; === FUNCTION: Capture Hotkey Press (Improved) ===
Func _CaptureHotkeyPress()
    Local $sResult = ""
    Local $sModifiers = ""
    Local $sKey = ""
    
    ; Kiem tra cac phim modifier
    Local $bCtrl = _IsKeyPressed(0x11)  ; VK_CONTROL
    Local $bAlt = _IsKeyPressed(0x12)   ; VK_MENU (Alt)
    Local $bShift = _IsKeyPressed(0x10) ; VK_SHIFT
    
    ; Neu khong co modifier nao, khong xu ly
    If Not $bCtrl And Not $bAlt And Not $bShift Then Return ""
    
    ; Tao chuoi modifier
    If $bCtrl Then $sModifiers &= "Ctrl+"
    If $bAlt Then $sModifiers &= "Alt+"
    If $bShift Then $sModifiers &= "Shift+"
    
    ; Kiem tra cac phim chinh (A-Z)
    For $i = 65 To 90 ; A-Z
        If _IsKeyPressed($i) Then
            $sKey = Chr($i)
            ExitLoop
        EndIf
    Next
    
    ; Kiem tra cac phim so (0-9)
    If $sKey = "" Then
        For $i = 48 To 57 ; 0-9
            If _IsKeyPressed($i) Then
                $sKey = Chr($i)
                ExitLoop
            EndIf
        Next
    EndIf
    
    ; Kiem tra cac phim F1-F12
    If $sKey = "" Then
        For $i = 112 To 123 ; F1-F12 (VK_F1 = 0x70 = 112)
            If _IsKeyPressed($i) Then
                $sKey = "F" & ($i - 111)
                ExitLoop
            EndIf
        Next
    EndIf
    
    ; Kiem tra cac phim dac biet khac
    If $sKey = "" Then
        If _IsKeyPressed(0xC0) Then
            $sKey = "`"      ; Backtick
        ElseIf _IsKeyPressed(0xBD) Then
            $sKey = "-"      ; Minus
        ElseIf _IsKeyPressed(0xBB) Then
            $sKey = "="      ; Equal
        ElseIf _IsKeyPressed(0xDB) Then
            $sKey = "["      ; Left bracket
        ElseIf _IsKeyPressed(0xDD) Then
            $sKey = "]"      ; Right bracket
        ElseIf _IsKeyPressed(0xDC) Then
            $sKey = "\"      ; Backslash
        ElseIf _IsKeyPressed(0xBA) Then
            $sKey = ";"      ; Semicolon
        ElseIf _IsKeyPressed(0xDE) Then
            $sKey = "'"      ; Quote
        ElseIf _IsKeyPressed(0xBC) Then
            $sKey = ","      ; Comma
        ElseIf _IsKeyPressed(0xBE) Then
            $sKey = "."      ; Period
        ElseIf _IsKeyPressed(0xBF) Then
            $sKey = "/"      ; Slash
        EndIf
    EndIf
    
    ; Neu co phim chinh, tra ve ket qua va doi 1 chut de tranh capture lien tuc
    If $sKey <> "" Then
        $sResult = $sModifiers & $sKey
        Sleep(200) ; Delay nho de tranh capture nhieu lan
    EndIf
    
    Return $sResult
EndFunc

; === FUNCTION: Check if key is pressed ===
Func _IsKeyPressed($iVKCode)
    Local $aRet = DllCall("user32.dll", "short", "GetAsyncKeyState", "int", $iVKCode)
    If @error Then Return False
    Return BitAND($aRet[0], 0x8000) <> 0
EndFunc

; ============================================
; UTILITY FUNCTIONS
; ============================================

; === FUNCTION: Lay thong tin file hotkey hien tai ===
Func _GetHotkeyFileInfo($sIniFile = "")
    If $sIniFile = "" Then $sIniFile = _GetHotkeyIniPath()
    
    If Not FileExists($sIniFile) Then
        Return "Chua co file hotkeys nao duoc tao."
    EndIf
    
    Local $aHotkeys = IniReadSection($sIniFile, "Hotkeys")
    Local $iCount = 0
    If IsArray($aHotkeys) Then $iCount = $aHotkeys[0][0]
    
    Local $aTime = FileGetTime($sIniFile, 0)
    Local $sTime = ""
    If IsArray($aTime) Then
        $sTime = $aTime[2] & "/" & $aTime[1] & "/" & $aTime[0] & " " & $aTime[3] & ":" & $aTime[4]
    EndIf
    
    Return "So hotkeys: " & $iCount & " | Cap nhat: " & $sTime & @CRLF & "Duong dan: " & $sIniFile
EndFunc
