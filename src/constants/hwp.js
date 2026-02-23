export const RECORD_TAGS = {
  16: "HWPTAG_DOCUMENT_PROPERTIES",
  17: "HWPTAG_ID_MAPPINGS",
  18: "HWPTAG_BIN_DATA",
  19: "HWPTAG_FACE_NAME",
  20: "HWPTAG_BORDER_FILL",
  21: "HWPTAG_CHAR_SHAPE",
  22: "HWPTAG_TAB_DEF",
  23: "HWPTAG_NUMBERING",
  24: "HWPTAG_BULLET",
  25: "HWPTAG_PARA_SHAPE",
  26: "HWPTAG_STYLE",
  27: "HWPTAG_DOC_DATA",
  28: "HWPTAG_DISTRIBUTE_DOC_DATA",
  30: "HWPTAG_COMPATIBLE_DOCUMENT",
  31: "HWPTAG_LAYOUT_COMPATIBILITY",
  32: "HWPTAG_TRACKCHANGE",
  66: "HWPTAG_PARA_HEADER",
  67: "HWPTAG_PARA_TEXT",
  68: "HWPTAG_PARA_CHAR_SHAPE",
  69: "HWPTAG_PARA_LINE_SEG",
  70: "HWPTAG_PARA_RANGE_TAG",
  71: "HWPTAG_CTRL_HEADER",
  72: "HWPTAG_LIST_HEADER",
  73: "HWPTAG_PAGE_DEF",
  74: "HWPTAG_FOOTNOTE_SHAPE",
  75: "HWPTAG_PAGE_BORDER_FILL",
  76: "HWPTAG_SHAPE_COMPONENT",
  77: "HWPTAG_TABLE",
  78: "HWPTAG_SHAPE_COMPONENT_LINE",
  79: "HWPTAG_SHAPE_COMPONENT_RECTANGLE",
  80: "HWPTAG_SHAPE_COMPONENT_ELLIPSE",
  81: "HWPTAG_SHAPE_COMPONENT_ARC",
  82: "HWPTAG_SHAPE_COMPONENT_POLYGON",
  83: "HWPTAG_SHAPE_COMPONENT_CURVE",
  84: "HWPTAG_SHAPE_COMPONENT_OLE",
  85: "HWPTAG_SHAPE_COMPONENT_PICTURE",
  86: "HWPTAG_SHAPE_COMPONENT_CONTAINER",
  87: "HWPTAG_CTRL_DATA",
  88: "HWPTAG_EQEDIT",
  90: "HWPTAG_SHAPE_COMPONENT_TEXTART",
  91: "HWPTAG_FORM_OBJECT",
  92: "HWPTAG_MEMO_SHAPE",
  93: "HWPTAG_MEMO_LIST",
  94: "HWPTAG_FORBIDDEN_CHAR",
  95: "HWPTAG_CHART_DATA",
  96: "HWPTAG_TRACK_CHANGE",
  97: "HWPTAG_TRACK_CHANGE_AUTHOR",
  98: "HWPTAG_VIDEO_DATA",
  115: "HWPTAG_SHAPE_COMPONENT_UNKNOWN",
  896: "LEGACY3_ASCII_BLOCK",
  897: "LEGACY3_UTF16_BLOCK",
};

export const RECORD_TAG_ALIASES = {
  HWPTAG_CTRL_HEAD: 71,
  HWPTAG_VIDEO_TDATA: 98,
};

export const DISTRIBUTE_DOC_RECORD_TAG = 28;

export const LINE_SEG_FLAGS = {
  0x00000001: "page-first-line",
  0x00000002: "column-first-line",
  0x00010000: "empty-segment",
  0x00020000: "line-first-segment",
  0x00040000: "line-last-segment",
  0x00080000: "auto-hyphenation",
  0x00100000: "indentation",
  0x00200000: "para-head-shape",
  0x80000000: "internal-property",
};
export const LINE_SEG_PAGE_FIRST_BIT = 0x00000001;
export const LINE_SEG_COLUMN_FIRST_BIT = 0x00000002;
export const PARA_SPLIT_SECTION_BIT = 0x01;
export const PARA_SPLIT_COLUMNS_DEF_BIT = 0x02;
export const PARA_SPLIT_PAGE_BIT = 0x04;
export const PARA_SPLIT_COLUMN_BIT = 0x08;
export const PAGE_Y_RESET_THRESHOLD = 1200;
export const PAGE_CONTENT_FALLBACK_HEIGHT_PX = 980;
export const PAGE_CONTENT_BOTTOM_GUARD_PX = 10;
export const PARAGRAPH_BLOCK_MIN_PX = 14;
export const TABLE_CHUNK_BASE_PX = 10;
export const TABLE_ROW_FALLBACK_PX = 22;
export const TABLE_ROW_MIN_PX = 9;
export const TEXT_PAGINATION_SCALE = 1.08;

export const SCRIPT_LANGS = ["ko", "en", "hanja", "jp", "other", "symbol", "user"];
export const UTF16LE_DECODER = new TextDecoder("utf-16le");

export const UNDERLINE_TYPE_NAMES = {
  0: "none",
  1: "under",
  3: "over",
};

export const OUTLINE_TYPE_NAMES = {
  0: "none",
  1: "solid",
  2: "dotted",
  3: "thick-solid",
  4: "dashed",
  5: "dash-dot",
  6: "dash-dot-dot",
};

export const SHADOW_TYPE_NAMES = {
  0: "none",
  1: "discrete",
  2: "continuous",
};

export const ACCENT_MARK_NAMES = {
  0: "none",
  1: "dot-filled",
  2: "dot-empty",
  3: "caron",
  4: "tilde",
  5: "middle-dot",
  6: "colon",
};

export const PARA_ALIGN_NAMES = {
  0: "justify",
  1: "left",
  2: "right",
  3: "center",
  4: "distribute",
  5: "divide",
};

export const PARA_LATIN_BREAK_NAMES = {
  0: "word",
  1: "hyphen",
  2: "char",
};

export const PARA_KOREAN_BREAK_NAMES = {
  0: "eojul",
  1: "char",
};

export const PARA_VERTICAL_ALIGN_NAMES = {
  0: "font",
  1: "top",
  2: "middle",
  3: "bottom",
};

export const PARA_HEADING_TYPE_NAMES = {
  0: "none",
  1: "outline",
  2: "number",
  3: "bullet",
};

export const PARA_LINE_SPACING_TYPE_NAMES = {
  0: "font-ratio",
  1: "fixed",
  2: "margin-only",
  3: "minimum",
};

export const CTRL_ID_NAMES = {
  secd: "Section Definition",
  cold: "Column Definition",
  head: "Header",
  foot: "Footer",
  "tbl ": "Table",
  "gso ": "Graphic Object",
  "eqed": "Equation",
  "fn  ": "Footnote",
  "en  ": "Endnote",
  atno: "Auto Number",
  bokm: "Bookmark",
  " %d": "Hidden Comment",
};

export const CONTROL_KIND_NAMES = {
  layout: "Layout",
  block: "Block",
  note: "Note",
  inline: "Inline",
};

export const CONTROL_KIND_BY_ID = {
  secd: "layout",
  cold: "layout",
  head: "layout",
  foot: "layout",
  "tbl ": "block",
  "gso ": "block",
  eqed: "block",
  "fn  ": "note",
  "en  ": "note",
};

export const LIST_TEXT_DIRECTION_NAMES = {
  0: "horizontal",
  1: "vertical",
};

export const LIST_LINE_BREAK_NAMES = {
  0: "normal",
  1: "keep-line",
  2: "expand-width",
};

export const LIST_VERTICAL_ALIGN_NAMES = {
  0: "top",
  1: "center",
  2: "bottom",
};

export const PAGE_BINDING_NAMES = {
  0: "left",
  1: "mirrored",
  2: "top",
};

export const COLUMN_TYPE_NAMES = {
  0: "normal",
  1: "distribute",
  2: "parallel",
};

export const COLUMN_DIRECTION_NAMES = {
  0: "left-to-right",
  1: "right-to-left",
  2: "mirror",
};

export const TAB_ALIGN_NAMES = {
  0: "left",
  1: "right",
  2: "center",
  3: "decimal",
};

export const TAB_LEADER_NAMES = {
  0: "none",
  1: "dot",
  2: "middle-dot",
  3: "hyphen",
  4: "underline",
};

export const BIN_STORAGE_TYPE_NAMES = {
  0: "link",
  1: "embedding",
  2: "storage",
};

export const BIN_COMPRESSION_NAMES = {
  0: "default",
  1: "compress",
  2: "decompress",
};

export const CONTROL_CHAR_SIZE = {
  0x00: 1,
  0x01: 8,
  0x02: 8,
  0x03: 8,
  0x04: 8,
  0x05: 8,
  0x06: 8,
  0x07: 8,
  0x08: 8,
  0x09: 8,
  0x0a: 1,
  0x0b: 8,
  0x0c: 8,
  0x0d: 1,
  0x0e: 8,
  0x0f: 8,
  0x10: 8,
  0x11: 8,
  0x12: 8,
  0x13: 8,
  0x14: 8,
  0x15: 8,
  0x16: 8,
  0x17: 8,
  0x18: 1,
  0x1e: 1,
  0x1f: 1,
};

export const IMAGE_MIME_BY_EXT = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  jpe: "image/jpeg",
  jfif: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
  svg: "image/svg+xml",
  tif: "image/tiff",
  tiff: "image/tiff",
  wmf: "image/wmf",
  emf: "image/emf",
  mp4: "video/mp4",
  m4v: "video/mp4",
  webm: "video/webm",
  mov: "video/quicktime",
  avi: "video/x-msvideo",
  wmv: "video/x-ms-wmv",
  mpg: "video/mpeg",
  mpeg: "video/mpeg",
};

export const SHAPE_OBJECT_CTRL_NAMES = {
  "$lin": "Line",
  "$rec": "Rectangle",
  "$ell": "Ellipse",
  "$arc": "Arc",
  "$pol": "Polygon",
  "$cur": "Curve",
  "$pic": "Picture",
  "$ole": "OLE",
  "$con": "Group",
  "tbl ": "Table",
  eqed: "Equation",
};

export const GRAPHIC_COMMON_RECORD_TAGS = new Set([78, 79, 80, 81, 82, 83, 84, 85, 86]);
export const GRAPHIC_DETAIL_RECORD_TAGS = new Set([76, 78, 79, 80, 81, 82, 83, 84, 85, 86, 90, 91, 95, 98, 115]);

export const OBJECT_TEXT_FLOW_NAMES = {
  0: "square",
  1: "tight",
  2: "through",
  3: "top-and-bottom",
  4: "behind-text",
  5: "in-front-text",
};

export const OBJECT_TEXT_SIDE_NAMES = {
  0: "both",
  1: "left-only",
  2: "right-only",
  3: "largest-only",
};

export const OBJECT_VERT_REL_TO_NAMES = {
  0: "paper",
  1: "page",
  2: "paragraph",
};

export const OBJECT_HORZ_REL_TO_NAMES = {
  0: "paper",
  1: "page",
  2: "column",
};

export const OBJECT_VERT_ALIGN_NAMES = {
  0: "top",
  1: "center",
  2: "bottom",
  3: "inside",
  4: "outside",
};

export const OBJECT_HORZ_ALIGN_NAMES = {
  0: "left",
  1: "center",
  2: "right",
  3: "inside",
  4: "outside",
};

export const PREVIEW_DEBUG = false;
export const PREVIEW_FONT_SCALE = 0.92;
