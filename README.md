# PDF to Word Fixer Pro v6.1.1

Ung dung AutoIt ho tro sua loi va chuan hoa tai lieu Word sau khi chuyen tu PDF, AI, hoac tai lieu bi vo dinh dang. Du an duoc to chuc theo kien truc modular de de mo rong, build, va test.

## Review nhanh

- Kien truc tach ro `Core`, `GUI`, `Modules`, `Shared`, `Tests`.
- Ket noi va thao tac voi Microsoft Word qua COM.
- Ho tro nhieu luong xu ly: sua loi PDF/Word, copy style, so sanh 2 ban Word, AI Format, quet de muc do an, va tien ich sua de muc hang loat.
- Tab `AI Format` da duoc nang cap de xu ly markdown, preview, latex, cleanup, beautify, va co nut rieng de chuan cong thuc.
- `FIX TAT CA - Chuan do an` chi con xu ly van ban Word, khong bao gom cong thuc, dung theo muc dich tach rieng tung nut.

## Tinh nang noi bat

- Sua loi sau OCR/PDF: xuong dong, khoang trang, hyphen, dash, non-breaking.
- Chuan hoa do an/bao cao: font, line spacing, margins, thut dau dong, dau cau, dong trong.
- AI Format: heading markdown, bold/italic, code block, links, bullets, numbered list, bang markdown, cleanup, preview.
- LaTeX va cong thuc: chuyen `$...$`, `$$...$$`, va chuan hoa hien thi cong thuc bang nut rieng.
- Quet de muc do an va cho phep chon style tu chinh van ban hoac van ban nguon.
- So sanh 2 file Word va liet ke chi tiet thay doi de xem truc quan.

## Cau truc thu muc

- `Main.au3`: entry point
- `Config.au3`: constants va globals
- `Core/`: ket noi Word, event loop
- `GUI/`: tabs, dialogs
- `Modules/`: logic xu ly chinh
- `Shared/`: helper dung chung
- `Tests/`: test tu dong va log
- `Releases/`: release notes va cac goi phat hanh

## Cach chay

1. Mo file Word can xu ly.
2. Chay `Main.au3` hoac dung ban `.exe`.
3. Nhan `Lam moi`, chon file, va `Ket noi`.
4. Su dung tab phu hop voi muc dich.

## Build

- Syntax check:
  - `C:\Program Files (x86)\AutoIt3\Au3Check.exe Main.au3`
- Compile:
  - `C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe /in Main.au3 /out Main_compiled.exe`

## Validation da thuc hien

- `Au3Check`: sach `0 error, 0 warning`
- Test AI LaTeX/Emoji: pass
- Da verify lai wiring `Config -> GUI -> EventLoop -> Module` cho nut `Chuan cong thuc`
- Da verify `FIX TAT CA` khong con goi xu ly cong thuc

## Luu y

- Nen backup truoc khi xu ly tai lieu.
- Tinh nang cong thuc duoc tach rieng khoi luong `FIX TAT CA`.
- Chuc nang chuan cong thuc hien tap trung vao chuan hoa hien thi cho `Word Equation` va object `MathType/Equation` dang co san.
