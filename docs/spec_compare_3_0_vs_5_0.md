# HWP Spec Compare: 3.0 vs 5.0

## Inputs
- `3.0.pdf` (title: Hwp Document File Formats 3.0 / HWPML, revision 1.2:20141105)
- `한글문서파일형식_5.0_revision1.3.pdf` (revision 1.3:20181108)
- Code baseline: `src/main.js`

## Quick summary
- `3.0.pdf` is split into two large parts:
  - Part I: legacy `한글 3.x` binary structure (paragraph/object/special-char oriented)
  - Part II: `HWPML` XML structure
- `5.0.pdf` is record-centric (`TagID/Level/Size`) and explicitly defines `HWPTAG_*` sets for DocInfo/BodyText.
- Current viewer implementation is aligned to 5.0-style record streams, not to legacy 3.x binary blocks nor HWPML XML tree.

## Structural differences

### 3.0 major chapters (from TOC)
- `I. 한글 3.x 문서 파일 구조`
- `4. 문단 자료 구조`
- `5. 문단 모양 자료 구조`
- `6. 글자 모양 자료 구조`
- `7~9. 정보/추가 정보 블록`
- `10. 특수 문자 자료 구조`
- `11. 그리기 개체 자료 구조`
- `12. OLE 개체 자료 구조`
- `II. HWPML 구조`

### 5.0 major chapters (from TOC)
- `I. 한글 5.0 파일 구조`
- `3. 한글 파일 구조` (스토리지 단위)
- `4. 데이터 레코드`
- `4.2. 문서 정보 레코드`
- `4.3. 본문 레코드`
- `4.4. 문서 이력 관리`

### Key delta
- 3.0 spec exposes object-specific binary structures and HWPML element definitions.
- 5.0 spec unifies body/doc data through generic record headers and `HWPTAG_*` taxonomy.
- 5.0 includes newer concepts explicitly (layout compatibility, compatible document, track-change, doc history, etc.).

## Tag-level comparison vs current code

### `HWPTAG_*` extraction from 3.0 / 5.0 PDFs
- `3.0.pdf`: `HWPTAG_*` not found in extracted text (legacy/HWPML-centric document structure)
- `5.0.pdf`: 52 unique `HWPTAG_*` names found

### Implemented in `src/main.js`
- `RECORD_TAGS` currently maps 38 tag names
- Coverage against 5.0 extracted tag list: `38/52 = 73.1%`

### Missing (present in 5.0 doc text, not mapped in `RECORD_TAGS`)
- `HWPTAG_CHART_DATA`
- `HWPTAG_FORBIDDEN_CHAR`
- `HWPTAG_FORM_OBJECT`
- `HWPTAG_MEMO_LIST`
- `HWPTAG_MEMO_SHAPE`
- `HWPTAG_SHAPE_COMPONENT_TEXTART`
- `HWPTAG_SHAPE_COMPONENT_UNKNOWN`
- `HWPTAG_TRACKCHANGE`
- `HWPTAG_TRACK_CHANGE`
- `HWPTAG_TRACK_CHANGE_AUTHOR`
- `HWPTAG_VIDEO_DATA`
- `HWPTAG_VIDEO_TDATA`
- (plus symbolic definitions shown in text such as `HWPTAG_BEGIN`, `HWPTAG_CTRL_HEAD`)

## Reality check on sample files
- `3.0.hwp` in this directory is OLE/CFB and has `FileHeader` version `5.0.5.0` (compressed).
- `BodyText/Section0` in `3.0.hwp` parses as 5.x style records (`66,67,68,71,73,74,75...`).
- So: `3.0.pdf` as spec document does not necessarily mean `3.0.hwp` sample requires a legacy 3.x parser path.

## Practical implication for web viewer
- Current architecture should stay 5.x-record-first.
- If true legacy 3.x binaries must be supported later, add a second parser pipeline:
  1. `legacy-3x` binary parser (non-5.x record schema)
  2. `hwpml` XML parser (for XML-encoded document path)
- Immediate value for current codebase is to close 5.0 gaps first (especially table-adjacent controls and media/chart tags).

## Suggested next implementation order
1. Complete 5.0 table-related descendants (`LIST_HEADER`, cell paragraph ownership, border/fill resolution) and validate against `rnd.hwp`.
2. Add missing 5.0 tag mappings/parsers for media and track-change-related records.
3. Add runtime capability detection and warning banner (`5.x record`, `legacy 3.x`, `HWPML`) before parse.
4. Only then start dedicated legacy 3.x parser branch if required by actual input corpus.

## Implementation status (2026-02-22)
- Completed `5.0` tag coverage expansion in code:
  - Added parsers/mappings for track-change, memo/forbidden-char, chart/video, textart/form-object families.
- Added runtime format profiling banner:
  - `HWP 5.x Record Stream`, `Legacy 3.x Binary`, `HWPML XML`, unknown fallback.
- Added legacy fallback parser branch:
  - non-record binaries now expose `legacy3-heuristic` ASCII/UTF-16 pseudo-records for inspection.
- Current extracted-tag coverage vs `5.0` PDF text:
  - `49/52` (`94.2%`) implemented in `RECORD_TAGS`.
  - Remaining unmodeled names are symbolic/alias identifiers in spec text: `HWPTAG_BEGIN`, `HWPTAG_CTRL_HEAD`, `HWPTAG_VIDEO_TDATA`.
