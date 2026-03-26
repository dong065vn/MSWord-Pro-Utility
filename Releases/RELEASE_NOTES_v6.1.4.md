# PDF to Word Fixer Pro v6.1.4

Release date: 2026-03-26

Highlights:
- Added a dedicated `AutoTextNumbering` module for text-based caption renumbering so the feature runs independently from generic tools and avoids cross-module logic conflicts.
- Added a new workflow to renumber captions from a sample such as `Hinh 1.1 Anh edge computing`, with support for `Hinh`, `Bang`, `Bieu do`, and `So do`.
- Added style-source selection, style filtering, style-only scanning, and optional style re-application after renumbering captions.
- Improved Word style access helpers to reduce COM noise during style enumeration and caption-style matching.
- Rebuilt the Windows executable with the logo icon embedded from `app_icons\app_icon_rounded.ico` and synchronized release logo assets from `app_icons`.

Validation:
- Au3Check clean for `Main.au3`
- Helper tests for caption sample parsing and group matching all pass
- Word COM integration test for auto text numbering passes end-to-end
- Runtime verification confirms correct renumbering order, style filtering, and style application

Primary artifacts:
- Main_compiled.exe
- Main_compiled_v6.1.4.exe
- PDF_to_Word_Fixer_Pro_v6.1.4_binary.zip
- PDF_to_Word_Fixer_Pro_v6.1.4_source.zip
