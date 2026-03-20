# Hướng dẫn tích hợp Normal.dotm Backup vào Main App

## Bước 1: Thêm buttons vào GUI

### Option A: Thêm vào Tab "Advanced"

Mở file `GUI/TabAdvanced.au3`, thêm vào cuối tab:

```autoit
; === Normal.dotm Management ===
GUICtrlCreateGroup("Normal.dotm Management", 10, 450, 380, 80)

Global $g_btnScanNormalDotm = GUICtrlCreateButton("Scan Normal.dotm", 20, 470, 110, 25)
GUICtrlSetTip(-1, "Quet thong tin styles trong Normal.dotm")

Global $g_btnBackupNormalDotm = GUICtrlCreateButton("Backup Normal.dotm", 140, 470, 120, 25)
GUICtrlSetTip(-1, "Backup toan bo Normal.dotm (styles + settings)")

Global $g_btnRestoreNormalDotm = GUICtrlCreateButton("Restore Normal.dotm", 270, 470, 110, 25)
GUICtrlSetTip(-1, "Restore Normal.dotm tu backup")

Global $g_btnOpenBackupFolder = GUICtrlCreateButton("Mo Backup Folder", 20, 500, 110, 25)
GUICtrlSetTip(-1, "Mo folder chua cac backups")

GUICtrlCreateGroup("", -99, -99, 1, 1)
```

### Option B: Tạo Tab mới "Normal.dotm"

Tạo file `GUI/TabNormalDotm.au3`:

```autoit
; ============================================
; TABNORMALDOTM.AU3
; Tab quan ly Normal.dotm
; ============================================

#include-once

Func _CreateTabNormalDotm()
    ; === Info ===
    GUICtrlCreateGroup("Normal.dotm Info", 10, 40, 380, 100)
    
    GUICtrlCreateLabel("Path:", 20, 60, 50, 20)
    Global $g_lblNormalDotmPath = GUICtrlCreateLabel("", 80, 60, 300, 20)
    
    GUICtrlCreateLabel("Styles:", 20, 85, 50, 20)
    Global $g_lblNormalDotmStyles = GUICtrlCreateLabel("", 80, 85, 300, 20)
    
    Global $g_btnRefreshInfo = GUICtrlCreateButton("Refresh", 20, 110, 80, 25)
    
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; === Actions ===
    GUICtrlCreateGroup("Actions", 10, 150, 380, 120)
    
    Global $g_btnScanNormalDotm = GUICtrlCreateButton("Scan Styles", 20, 170, 110, 30)
    GUICtrlSetTip(-1, "Quet va hien thi thong tin styles")
    
    Global $g_btnBackupNormalDotm = GUICtrlCreateButton("Full Backup", 140, 170, 110, 30)
    GUICtrlSetTip(-1, "Backup toan bo Normal.dotm")
    
    Global $g_btnRestoreNormalDotm = GUICtrlCreateButton("Restore", 260, 170, 110, 30)
    GUICtrlSetTip(-1, "Restore tu backup")
    
    Global $g_btnExportTxt = GUICtrlCreateButton("Export TXT", 20, 210, 110, 30)
    Global $g_btnExportCSV = GUICtrlCreateButton("Export CSV", 140, 210, 110, 30)
    Global $g_btnOpenBackupFolder = GUICtrlCreateButton("Mo Backup Folder", 260, 210, 110, 30)
    
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; === Backups List ===
    GUICtrlCreateGroup("Backups", 10, 280, 780, 290)
    
    Global $g_listNormalDotmBackups = GUICtrlCreateListView("Backup Name|Date|Styles|Size", _
        20, 300, 760, 250, $LVS_REPORT + $LVS_SINGLESEL, $LVS_EX_FULLROWSELECT + $LVS_EX_GRIDLINES)
    _GUICtrlListView_SetColumnWidth($g_listNormalDotmBackups, 0, 300)
    _GUICtrlListView_SetColumnWidth($g_listNormalDotmBackups, 1, 150)
    _GUICtrlListView_SetColumnWidth($g_listNormalDotmBackups, 2, 100)
    _GUICtrlListView_SetColumnWidth($g_listNormalDotmBackups, 3, 100)
    
    Global $g_btnRefreshBackups = GUICtrlCreateButton("Refresh", 20, 555, 80, 25)
    Global $g_btnDeleteBackup = GUICtrlCreateButton("Delete", 110, 555, 80, 25)
    
    GUICtrlCreateGroup("", -99, -99, 1, 1)
EndFunc
```

## Bước 2: Thêm vào Config.au3

Mở `Config.au3`, thêm vào phần Global variables:

```autoit
; === Tab: Normal.dotm (NEW) ===
Global $g_btnScanNormalDotm, $g_btnBackupNormalDotm, $g_btnRestoreNormalDotm
Global $g_btnOpenBackupFolder, $g_btnExportTxt, $g_btnExportCSV
Global $g_listNormalDotmBackups, $g_lblNormalDotmPath, $g_lblNormalDotmStyles
```

## Bước 3: Thêm handlers vào EventLoop

Mở `Core/EventLoop.au3`, thêm vào event loop:

```autoit
; === Normal.dotm Management ===
Case $g_btnScanNormalDotm
    _OnScanNormalDotm()

Case $g_btnBackupNormalDotm
    _OnBackupNormalDotm()

Case $g_btnRestoreNormalDotm
    _OnRestoreNormalDotm()

Case $g_btnOpenBackupFolder
    _OnOpenBackupFolder()

Case $g_btnExportTxt
    _OnExportNormalDotmTxt()

Case $g_btnExportCSV
    _OnExportNormalDotmCSV()
```

## Bước 4: Include module vào Main.au3

Mở `Main.au3`, thêm include:

```autoit
#include "Modules\NormalDotmManager.au3"
```

## Bước 5: (Optional) Thêm vào Menu

Nếu có menu bar, thêm menu item:

```autoit
; Menu: Tools
Local $menuTools = GUICtrlCreateMenu("Tools")
Local $menuBackupNormalDotm = GUICtrlCreateMenuItem("Backup Normal.dotm", $menuTools)
Local $menuRestoreNormalDotm = GUICtrlCreateMenuItem("Restore Normal.dotm", $menuTools)
Local $menuOpenBackupFolder = GUICtrlCreateMenuItem("Open Backup Folder", $menuTools)

; Handler
Case $menuBackupNormalDotm
    _OnBackupNormalDotm()
Case $menuRestoreNormalDotm
    _OnRestoreNormalDotm()
Case $menuOpenBackupFolder
    _OnOpenBackupFolder()
```

## Bước 6: (Optional) Auto-backup khi đóng app

Trong `Main.au3`, trước khi exit:

```autoit
; Auto-backup truoc khi dong app
If IsObj($g_oWord) Then
    _QuickBackupNormalDotm() ; Silent backup, khong hoi
EndIf
```

## Bước 7: Test

1. Compile lại Main.au3
2. Chạy app
3. Kết nối Word
4. Test các chức năng:
   - Scan Normal.dotm
   - Backup
   - Restore
   - Export TXT/CSV

## Các functions có sẵn

Từ `Modules/NormalDotmManager.au3`:

```autoit
_OnBackupNormalDotm()          ; Backup với confirm
_OnRestoreNormalDotm()         ; Restore với GUI chọn backup
_OnScanNormalDotm()            ; Scan và hiển thị thống kê
_OnOpenBackupFolder()          ; Mở folder backups
_QuickBackupNormalDotm()       ; Backup nhanh, không hỏi
```

## Hotkeys (Optional)

Thêm hotkeys để backup nhanh:

```autoit
; Trong Main.au3
HotKeySet("^!b", "_HotkeyBackupNormalDotm") ; Ctrl+Alt+B

Func _HotkeyBackupNormalDotm()
    If IsObj($g_oWord) Then
        _QuickBackupNormalDotm()
        ToolTip("Normal.dotm backed up!", @DesktopWidth - 200, @DesktopHeight - 100)
        Sleep(2000)
        ToolTip("")
    EndIf
EndFunc
```

## Scheduled Backup (Optional)

Backup tự động mỗi 30 phút:

```autoit
; Trong Main.au3
AdlibRegister("_AutoBackupNormalDotm", 1800000) ; 30 phút = 1800000 ms

Func _AutoBackupNormalDotm()
    If IsObj($g_oWord) Then
        _QuickBackupNormalDotm()
        _UpdateProgress("Auto-backup Normal.dotm thanh cong")
    EndIf
EndFunc
```

## Troubleshooting

### Lỗi: "Khong tim thay module"
- Kiểm tra đường dẫn include
- Đảm bảo file `Tools/NormalDotmBackup.au3` tồn tại

### Lỗi: "Function undefined"
- Kiểm tra đã include `Modules/NormalDotmManager.au3` chưa
- Kiểm tra đã declare global variables trong `Config.au3` chưa

### Backup folder không tạo được
- Kiểm tra quyền ghi file
- Thử chạy app với quyền Administrator

## Best Practices

1. **Backup trước khi thay đổi lớn**: Luôn backup trước khi import styles từ template khác
2. **Backup định kỳ**: Setup auto-backup mỗi ngày/tuần
3. **Giữ nhiều backups**: Không xóa backups cũ, disk space rẻ
4. **Test restore**: Thỉnh thoảng test restore để đảm bảo backups hoạt động
5. **Export CSV**: Export ra CSV để audit styles định kỳ

## Support

Nếu gặp vấn đề, check:
1. Console log (ConsoleWrite)
2. File `MANIFEST.txt` trong backup folder
3. Test bằng `Test_NormalDotmBackup.au3`
4. Chạy GUI standalone `NormalDotmBackupGUI.au3`
