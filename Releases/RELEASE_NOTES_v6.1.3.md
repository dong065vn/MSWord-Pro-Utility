# PDF to Word Fixer Pro v6.1.3

Release date: 2026-03-26

Highlights:
- Split `Numbering -> Text` and `Bullet -> Text` into separate detection flows so each feature can recognize both real Word lists and plain-text markers copied from PDF/OCR.
- Added plain-text numbering detection for formats such as `1.`, `1)`, `1 -`, and `1:` before converting them into stable text output.
- Added plain-text bullet detection for formats such as `-`, `+`, `*`, `•`, and checkbox markers like `[x]`.
- Improved AI numbered-list conversion so tabbed heading patterns like `1.[tab] Tieu de` followed by wrapped body text no longer pull the following paragraph into the same numbered item.
- Compiled the Windows executable with the project logo icon embedded from `app_icons\app_icon_rounded.ico`.

Validation:
- Au3Check clean for `Main.au3`
- Added regression coverage for markdown numbered-list splitting
- Added helper tests for plain-text numbering and bullet marker detection
- Extracted the icon from the freshly compiled `Main_compiled.exe` to verify the logo is embedded

Primary artifacts:
- Main_compiled.exe
- PDF_to_Word_Fixer_Pro_v6.1.3_binary.zip
- PDF_to_Word_Fixer_Pro_v6.1.3_source.zip
