# PDF to Word Fixer Pro v6.1.5

Release date: 2026-03-26

Highlights:
- Fixed the Windows Explorer executable icon so the rounded project logo now appears correctly in small and large icon views.
- Updated the release build workflow to generate a multi-size `.ico` from the rounded PNG icon set before embedding it into the compiled `.exe`.
- Synced `Resources\icon.ico` and `Resources\icon.png` with the same rounded icon source used for the executable.

Validation:
- Au3Check clean for `Main.au3`
- Explorer small icon verification passed against `icon_rounded_16x16.png`
- Explorer large icon verification passed against `icon_rounded_32x32.png`
- Release binary and source assets rebuilt for v6.1.5

Primary artifacts:
- Main_compiled.exe
- Main_compiled_v6.1.5.exe
- PDF_to_Word_Fixer_Pro_v6.1.5_binary.zip
- PDF_to_Word_Fixer_Pro_v6.1.5_source.zip
