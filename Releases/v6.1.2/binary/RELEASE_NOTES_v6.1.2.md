# PDF to Word Fixer Pro v6.1.2

Release date: 2026-03-25

Highlights:
- Expanded thesis heading detection to recognize special top-level sections such as `MUC LUC`, `DANH MUC TU VIET TAT`, `LOI CAM ON`, `LOI MO DAU`, and `PHAN A: ...` as chapter-level headings.
- Improved thesis heading scan so numbering copied from PDF like `1. 1`, `1. 1. 1`, and `2. 1. 4` is detected at the correct level.
- Normalized detected thesis headings during style application, including cleanup such as `1. 1. 1` -> `1.1.1` and `CHUONG 1 :` -> `CHUONG 1:`.
- Upgraded the Quick Utils prefix-normalization tool to support more numbering formats and custom separators, with clearer in-app guidance.
- Adjusted `FIX TAT CA - Chuan do an` so it no longer auto-creates bullet lists or numbered lists.
