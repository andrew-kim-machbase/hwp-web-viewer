# HWP Raw Viewer (MVP)

Browser viewer for raw `.hwp` internals:

- CFB/OLE stream tree
- HWP record table (tag/level/size)
- Paragraph text extraction from `HWPTAG_PARA_TEXT`
- Payload hex view

## Run

```bash
npm install
npm run dev
```

Open the local Vite URL, then select a `.hwp` file.

## PDF Diff Loop

Run automated preview-vs-PDF comparison (captures preview pages with Playwright and compares against rendered PDF pages):

```bash
npm run diff:preview -- --max-pages 30
```

Notes:

- By default, the script does **not** reuse an already-running dev server on the same URL (to avoid stale comparisons).
- If `http://127.0.0.1:4173` is already in use, either:
  - pass a different URL, e.g. `--url http://127.0.0.1:4188`, or
  - explicitly opt in to reuse: `--reuse-server`

Outputs:

- `artifacts/diff/<case>/preview/preview-page-XXX.png`
- `artifacts/diff/<case>/report.json`
- `artifacts/diff/summary.json`

## Current parsing scope

- Parses `FileHeader` (`signature`, `version`, `flags`)
- Uses compressed bit (`flags bit0`) to decode `DocInfo`, `BodyText/Section*`
- Uses distributable bit (`flags bit2`) to decode `ViewText/Section*`:
  - reads head record `HWPTAG_DISTRIBUTE_DOC_DATA(28)`
  - derives key with MSVC `rand()` compatible stream decoder
  - decrypts tail with `AES-128-ECB`
  - inflates raw-deflate payload before record parsing
- Reads 4-byte record header (`10-bit tag`, `10-bit level`, `12-bit size`)
- Supports extended header when `size == 0xFFF`
- Extracts UTF-16 paragraph text with basic control handling (`[CTRL:*]` placeholder)
- Links records by paragraph (`HWPTAG_PARA_HEADER=66`)
- Decodes `HWPTAG_PARA_CHAR_SHAPE(68)` runs (`startPos`, `charShapeId`)
- Decodes `HWPTAG_PARA_LINE_SEG(69)` segments (`text pos`, `height`, `flags`)
- Parses `DocInfo` mappings and resolves:
  - `paraStyleId` -> style name / refs
  - `paraShapeId` -> paragraph shape summary
  - `charShapeId` -> primary font + size from `CHAR_SHAPE`
- Decodes `CHAR_SHAPE` detail fields:
  - attribute bits (italic/bold/underline/outline/shadow/kerning, etc.)
  - colors (`text`, `underline`, `shade`, `shadow`, `strikeout`)
  - shadow offset / border-fill reference
- Decodes `PARA_SHAPE` detail fields:
  - property1/2/3 bitfields (align, heading type/level, spacing type, auto-spacing flags)
  - border spacing values and modern/legacy line-spacing fields
- Adds a `DocInfo Catalog` UI panel:
  - FaceName / Style / CharShape / ParaShape tables
  - TrackChange / TrackChangeAuthor / ForbiddenChar / MemoShape tables
- Adds `Document Preview (Beta)`:
  - Renders BodyText paragraphs with mapped alignment/font/size/bold/italic/color
  - Applies paragraph shape spacing/indent/margins for closer reading layout
  - Applies `PARA_CHAR_SHAPE(68)` runs to inline text spans (font family/size/weight/style/color/underline/strike)
  - Applies `PARA_LINE_SEG(69)` starts/metrics to line-block rendering (line split, horizontal offset, line-height spacing)
  - Splits preview into multiple pages using `PARA_LINE_SEG(69)` first-line flags (`page-first-line`, `column-first-line`)
  - Parses `DocInfo` `TAB_DEF(22)`, `NUMBERING(23)`, `BULLET(24)` and applies list markers + tab-stop spacing in preview text
  - Parses `DocInfo` `BIN_DATA(18)` and resolves embedded image streams (`Root Entry/BinData/BIN*`)
  - For picture graphics (`gso` + `HWPTAG_SHAPE_COMPONENT_PICTURE(85)`), resolves `binDataId` and renders real `<img>` preview when decodable
  - Uses DocInfo style references to approximate visual reading flow
  - Renders control blocks from `CTRL_HEADER(71)` with kind/type separation:
    - layout (`secd`, `cold`, `head`, `foot`)
    - block (`tbl `, `gso `, `eqed`)
    - note (`fn  `, `en  `)
  - Summarizes per-control subtree records (level-based descendants) and raw word previews
  - For table controls (`tbl ` + `HWPTAG_TABLE(77)`), renders a beta grid preview:
    - parses row/col/cell metadata, cell spans, row sizes, inner margins, zones
    - supports compact table variants where cell metadata is provided via descendant `HWPTAG_LIST_HEADER(72)` records
    - maps descendant paragraph text into cells using each cell's paragraph-count metadata
    - displays an HTML table with merged-cell structure, cell text, and per-cell metrics
    - applies row-height based table chunking for overflow pagination
  - For section controls (`secd` + `HWPTAG_PAGE_DEF(73)`), applies paper/margin-derived page width to preview sheet
  - For column controls (`cold`), parses column count/gap/direction/width hints and applies multi-column flow segments in preview
  - For graphic controls (`gso `), decodes object-flow hints and renders floating/inline object blocks with text wrap approximation
  - Applies content-height overflow pagination:
    - splits oversized control-only paragraphs across pages
    - keeps absolute-overlay graphics out of normal flow height accumulation
  - Parses additional body/object records:
    - `HWPTAG_SHAPE_COMPONENT_TEXTART(90)`
    - `HWPTAG_FORM_OBJECT(91)`
    - `HWPTAG_MEMO_SHAPE(92)` / `HWPTAG_MEMO_LIST(93)`
    - `HWPTAG_CHART_DATA(95)`
    - `HWPTAG_VIDEO_DATA(98)` (alias `HWPTAG_VIDEO_TDATA`)
    - `HWPTAG_SHAPE_COMPONENT_UNKNOWN(115)`
  - Parses additional DocInfo records:
    - `HWPTAG_TRACKCHANGE(32)`
    - `HWPTAG_TRACK_CHANGE(96)`
    - `HWPTAG_TRACK_CHANGE_AUTHOR(97)`
    - `HWPTAG_FORBIDDEN_CHAR(94)`
  - Uses reader-first rendering by default (layout-debug controls hidden unless debug mode)
  - For distributable documents, preview source automatically switches from `BodyText` to decrypted `ViewText`
 - Adds runtime format profiling banner:
   - `HWP 5.x Record Stream`
   - `Legacy 3.x Binary`
   - `HWPML XML`
   - unknown fallback
 - Adds fallback `legacy3-heuristic` analysis path for non-record raw binaries:
   - extracts ASCII/UTF-16 text blocks as pseudo-records for inspection

## Notes

- This is a raw-format viewer, not a full layout renderer.
- `ViewText/Section*` streams may not parse as plain record streams depending on document mode.
- Preview does not yet replicate full HWP pagination/table/object layout.
- True legacy 3.x binary layout reproduction is partial; current branch focuses on inspection-grade parsing and detection.

## References used

- https://tech.hancom.com/python-hwp-parsing-1/
- https://tech.hancom.com/python-hwp-parsing-2/
- https://tech.hancom.com/python-hwpx-parsing-2/
