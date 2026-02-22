import { escapeHtml } from "../utils/html.js";
import {
  BIN_STORAGE_TYPE_NAMES,
  BIN_COMPRESSION_NAMES,
  PARA_ALIGN_NAMES,
  PARA_LINE_SPACING_TYPE_NAMES,
  PARA_HEADING_TYPE_NAMES,
} from "../constants/hwp.js";

export function renderDocInfoPanelContent(docInfoPanelEl, doc, formatters = {}) {
  const docInfo = doc?.docInfo;
  const binDataById = doc?.binDataById ?? new Map();
  const {
    formatStyleName = (style) => style?.name || "",
    formatParaShapeSummary = () => "",
    describeDocCharShape = () => "",
    colorRefToText = () => "-",
  } = formatters;

  if (!docInfo) {
    docInfoPanelEl.innerHTML = `<div class="empty">DocInfo stream is not available.</div>`;
    return;
  }

  const faceRows = docInfo.faceNames
    .map(
      (face) => `
      <tr>
        <td>${face.id}</td>
        <td>${escapeHtml(face.primaryName || "-")}</td>
        <td>${escapeHtml(face.alternativeName || "-")}</td>
      </tr>
    `
    )
    .join("");

  const styleRows = docInfo.styles
    .map((style) => {
      const paraShape = docInfo.paraShapeById.get(style.paraShapeId);
      const charShape = docInfo.charShapeById.get(style.charShapeId);
      return `
        <tr>
          <td>${style.id}</td>
          <td>${escapeHtml(formatStyleName(style))}</td>
          <td>${style.styleType ?? "-"}</td>
          <td>${style.nextStyleId ?? "-"}</td>
          <td>${style.paraShapeId ?? "-"}</td>
          <td>${escapeHtml(paraShape ? formatParaShapeSummary(paraShape) : "-")}</td>
          <td>${style.charShapeId ?? "-"}</td>
          <td>${escapeHtml(charShape ? describeDocCharShape(charShape) : "-")}</td>
        </tr>
      `;
    })
    .join("");

  const tabDefRows = docInfo.tabDefs
    .map((tabDef) => {
      const stops = tabDef.tabStops.length
        ? tabDef.tabStops
            .map((stop) => `${Math.round(stop.positionPx)}px/${stop.alignName}${stop.leader !== 0 ? `/${stop.leaderName}` : ""}`)
            .join(", ")
        : "-";
      return `
        <tr>
          <td>${tabDef.id}</td>
          <td>${tabDef.count}</td>
          <td>${escapeHtml(stops)}</td>
        </tr>
      `;
    })
    .join("");

  const numberingRows = docInfo.numberings
    .map((numbering) => {
      const levels = numbering.levels
        .map((level) => `${level.level}:${level.format || "-"}`)
        .join(" | ");
      return `
        <tr>
          <td>${numbering.id}</td>
          <td>${escapeHtml(levels || "-")}</td>
        </tr>
      `;
    })
    .join("");

  const bulletRows = docInfo.bullets
    .map((bullet) => {
      return `
        <tr>
          <td>${bullet.id}</td>
          <td>${escapeHtml(bullet.marker || "-")}</td>
          <td>${bullet.charCode ? `U+${bullet.charCode.toString(16).toUpperCase().padStart(4, "0")}` : "-"}</td>
        </tr>
      `;
    })
    .join("");

  const memoShapeRows = docInfo.memoShapes
    .map((item) => {
      const preview = item.strings?.length ? item.strings.join(" | ") : "-";
      return `
        <tr>
          <td>${item.id}</td>
          <td>${item.property != null ? `0x${item.property.toString(16)}` : "-"}</td>
          <td>${item.width ?? "-"}</td>
          <td>${item.margin ?? "-"}</td>
          <td>${escapeHtml(preview)}</td>
        </tr>
      `;
    })
    .join("");

  const forbiddenRows = docInfo.forbiddenChars
    .map((item) => {
      const sample = item.sample?.length
        ? item.sample.map((char) => `${char.char}(U+${char.code.toString(16).toUpperCase().padStart(4, "0")})`).join(", ")
        : "-";
      return `
        <tr>
          <td>${item.id}</td>
          <td>${item.countByHeader ?? "-"}</td>
          <td>${escapeHtml(sample)}</td>
        </tr>
      `;
    })
    .join("");

  const trackChangeRows = docInfo.trackChanges
    .map((item) => {
      return `
        <tr>
          <td>${item.id}</td>
          <td>${item.tag}</td>
          <td>${item.flags != null ? `0x${item.flags.toString(16)}` : "-"}</td>
          <td>${item.authorId ?? "-"}</td>
          <td>${escapeHtml(item.summary || "-")}</td>
        </tr>
      `;
    })
    .join("");

  const trackAuthorRows = docInfo.trackChangeAuthors
    .map((item) => {
      const texts = item.strings?.length ? item.strings.join(" | ") : "-";
      return `
        <tr>
          <td>${item.id}</td>
          <td>${item.authorIndex ?? "-"}</td>
          <td>${escapeHtml(item.name || "-")}</td>
          <td>${escapeHtml(texts)}</td>
        </tr>
      `;
    })
    .join("");

  const binDataRows = docInfo.binDataRecords
    .map((item) => {
      const resolved = item.id != null ? binDataById.get(item.id) ?? null : null;
      const ext = item.extension || resolved?.extension || "-";
      const storage = item.storageTypeName || BIN_STORAGE_TYPE_NAMES[item.storageType] || item.storageType;
      const compression = item.compressionName || BIN_COMPRESSION_NAMES[item.compression] || item.compression;
      const stream = resolved?.streamPath ? resolved.streamPath.replace(/^Root Entry\//, "") : "-";
      const decodeInfo = resolved?.decodeStrategy ? `${resolved.decodeStrategy}${resolved.format ? `:${resolved.format}` : ""}` : "-";
      const sizeInfo = resolved ? `${resolved.rawSize}/${resolved.decodedSize}` : "-";

      return `
        <tr>
          <td>${item.id ?? "-"}</td>
          <td>${escapeHtml(String(storage ?? "-"))}</td>
          <td>${escapeHtml(String(compression ?? "-"))}</td>
          <td>${escapeHtml(ext)}</td>
          <td>${escapeHtml(stream)}</td>
          <td>${escapeHtml(sizeInfo)}</td>
          <td>${escapeHtml(decodeInfo)}</td>
        </tr>
      `;
    })
    .join("");

  const charShapeRows = docInfo.charShapes
    .map((shape) => {
      const attrText = shape.attributeBits?.labels?.join(", ") || "-";
      return `
        <tr>
          <td>${shape.id}</td>
          <td>${escapeHtml(shape.primaryFont || "-")}</td>
          <td>${shape.baseSize != null ? (shape.baseSize / 100).toFixed(1) : "-"}</td>
          <td>${escapeHtml(attrText)}</td>
          <td>${escapeHtml(colorRefToText(shape.textColor))}</td>
          <td>${shape.borderFillId ?? "-"}</td>
        </tr>
      `;
    })
    .join("");

  const paraShapeRows = docInfo.paraShapes
    .map((shape) => {
      const attr = shape.property1Bits;
      const align = attr ? PARA_ALIGN_NAMES[attr.align] ?? attr.align : "-";
      const spacingType =
        shape.property3Bits?.lineSpacingTypeName ??
        PARA_LINE_SPACING_TYPE_NAMES[shape.property1Bits?.lineSpacingTypeLegacy] ??
        shape.property1Bits?.lineSpacingTypeLegacy ??
        "-";
      const heading = attr ? PARA_HEADING_TYPE_NAMES[attr.headingType] ?? attr.headingType : "-";
      return `
        <tr>
          <td>${shape.id}</td>
          <td>${escapeHtml(String(align))}</td>
          <td>${shape.leftMargin ?? "-"}</td>
          <td>${shape.rightMargin ?? "-"}</td>
          <td>${shape.indent ?? "-"}</td>
          <td>${escapeHtml(String(spacingType))}</td>
          <td>${shape.lineSpacing ?? "-"}</td>
          <td>${shape.borderFillId ?? "-"}</td>
          <td>${escapeHtml(String(heading))}</td>
        </tr>
      `;
    })
    .join("");

  docInfoPanelEl.innerHTML = `
    <div class="docinfo-top">
      <span class="chip">FaceName ${docInfo.faceNames.length}</span>
      <span class="chip">BinData ${docInfo.binDataRecords.length}</span>
      <span class="chip">TabDef ${docInfo.tabDefs.length}</span>
      <span class="chip">Numbering ${docInfo.numberings.length}</span>
      <span class="chip">Bullet ${docInfo.bullets.length}</span>
      <span class="chip">TrackChange ${docInfo.trackChanges.length}</span>
      <span class="chip">TrackAuthor ${docInfo.trackChangeAuthors.length}</span>
      <span class="chip">Forbidden ${docInfo.forbiddenChars.length}</span>
      <span class="chip">MemoShape ${docInfo.memoShapes.length}</span>
      <span class="chip">Style ${docInfo.styles.length}</span>
      <span class="chip">CharShape ${docInfo.charShapes.length}</span>
      <span class="chip">ParaShape ${docInfo.paraShapes.length}</span>
      <span class="chip">IDMappings ${docInfo.idMappings?.counts?.length ?? 0}</span>
    </div>

    <details class="docinfo-details" open>
      <summary>Face Names</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Primary</th><th>Alternative</th></tr>
          </thead>
          <tbody>${faceRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Bin Data</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Storage</th><th>Compression</th><th>Ext</th><th>Stream</th><th>Raw/Decoded</th><th>Decode</th></tr>
          </thead>
          <tbody>${binDataRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Tab Defs</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Stops</th><th>Preview</th></tr>
          </thead>
          <tbody>${tabDefRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Numberings</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Level Formats</th></tr>
          </thead>
          <tbody>${numberingRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Bullets</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Marker</th><th>Code</th></tr>
          </thead>
          <tbody>${bulletRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Track Changes</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Tag</th><th>Flags</th><th>AuthorID</th><th>Summary</th></tr>
          </thead>
          <tbody>${trackChangeRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Track Change Authors</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>AuthorIdx</th><th>Name</th><th>Strings</th></tr>
          </thead>
          <tbody>${trackAuthorRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Forbidden Characters</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Count</th><th>Sample</th></tr>
          </thead>
          <tbody>${forbiddenRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Memo Shapes</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Property</th><th>Width</th><th>Margin</th><th>Preview</th></tr>
          </thead>
          <tbody>${memoShapeRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Styles</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr>
              <th>ID</th><th>Name</th><th>Type</th><th>Next</th><th>ParaShapeID</th><th>ParaShape</th><th>CharShapeID</th><th>CharShape</th>
            </tr>
          </thead>
          <tbody>${styleRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Char Shapes</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Primary Font</th><th>Size(pt)</th><th>Attributes</th><th>TextColor</th><th>BorderFill</th></tr>
          </thead>
          <tbody>${charShapeRows}</tbody>
        </table>
      </div>
    </details>

    <details class="docinfo-details">
      <summary>Para Shapes</summary>
      <div class="record-table-wrap">
        <table class="record-table">
          <thead>
            <tr><th>ID</th><th>Align</th><th>Left</th><th>Right</th><th>Indent</th><th>SpacingType</th><th>LineSpacing</th><th>BorderFill</th><th>Heading</th></tr>
          </thead>
          <tbody>${paraShapeRows}</tbody>
        </table>
      </div>
    </details>
  `;
}
