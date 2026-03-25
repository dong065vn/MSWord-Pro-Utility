PDF TO WORD FIXER PRO v6.0 - MODULAR ARCHITECTURE
=================================================

HUONG DAN SU DUNG:
1. Mo file Word can xu ly
2. Chay Main.au3 (hoac compile thanh .exe)
3. Nhan "Lam moi" de quet danh sach file
4. Chon file va nhan "Ket noi"
5. Su dung cac tab de xu ly

CAU TRUC THU MUC:
- Main.au3         : Entry point
- Config.au3       : Constants & Globals
- Core/            : Ket noi Word, Event loop
- Modules/         : Logic xu ly (PDFFix, Format, Tools...)
- GUI/             : Giao dien (Tabs, Dialogs)
- Shared/          : Helpers dung chung
- Resources/       : Config files

TINH NANG CHINH:
- Tab 1: Sua loi PDF (xuong dong, khoang trang, hyphen...)
- Tab 2: Dinh dang (font, le trang, heading, caption...)
- Tab 3: Cong cu (tim thay the, hinh anh, bang...)
- Tab 4: Muc luc va Tai lieu tham khao
- Tab 5: Copy Style giua cac file (OrganizerCopy)
- Tab 6: Tinh nang nang cao (export, protect...)

LUU Y:
- Nen Backup truoc khi sua
- Dung Ctrl+Z de Undo
- Khong chay "Run as Administrator" de tranh loi UAC

PHIEN BAN: 6.1.1 (Modular Architecture)
