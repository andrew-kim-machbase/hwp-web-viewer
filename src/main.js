import CryptoJS from "crypto-js";
import "./styles.css";
import { toUint8Array, extractReadableAscii, bytesToAscii } from "./utils/bytes.js";
import { safeInflate } from "./utils/compression.js";
import { formatHex, formatBytes } from "./utils/format.js";
import { clamp, hwpToPx, hwpToMm } from "./utils/numeric.js";
import { escapeHtml, escapeHtmlAttr } from "./utils/html.js";
import { renderAppShell } from "./ui/appShell.js";
import { buildDocumentFromBytes } from "./parser/documentLoader.js";
import { isLikelyRecordStreamPath, mayBeCompressedRecordStreamPath } from "./parser/streamPath.js";
import { renderDocInfoPanelContent } from "./render/docInfoPanel.js";
import { renderSummaryPanel } from "./render/summary.js";
import { renderRecordPanelView, renderStreamListPanel } from "./render/streamPanels.js";
import {
  RECORD_TAGS,
  RECORD_TAG_ALIASES,
  DISTRIBUTE_DOC_RECORD_TAG,
  LINE_SEG_FLAGS,
  LINE_SEG_PAGE_FIRST_BIT,
  LINE_SEG_COLUMN_FIRST_BIT,
  PARA_SPLIT_SECTION_BIT,
  PARA_SPLIT_COLUMNS_DEF_BIT,
  PARA_SPLIT_PAGE_BIT,
  PARA_SPLIT_COLUMN_BIT,
  PAGE_Y_RESET_THRESHOLD,
  PAGE_CONTENT_FALLBACK_HEIGHT_PX,
  PAGE_CONTENT_BOTTOM_GUARD_PX,
  PARAGRAPH_BLOCK_MIN_PX,
  TABLE_CHUNK_BASE_PX,
  TABLE_ROW_FALLBACK_PX,
  TABLE_ROW_MIN_PX,
  TEXT_PAGINATION_SCALE,
  SCRIPT_LANGS,
  UTF16LE_DECODER,
  UNDERLINE_TYPE_NAMES,
  OUTLINE_TYPE_NAMES,
  SHADOW_TYPE_NAMES,
  ACCENT_MARK_NAMES,
  PARA_ALIGN_NAMES,
  PARA_LATIN_BREAK_NAMES,
  PARA_KOREAN_BREAK_NAMES,
  PARA_VERTICAL_ALIGN_NAMES,
  PARA_HEADING_TYPE_NAMES,
  PARA_LINE_SPACING_TYPE_NAMES,
  CTRL_ID_NAMES,
  CONTROL_KIND_NAMES,
  CONTROL_KIND_BY_ID,
  LIST_TEXT_DIRECTION_NAMES,
  LIST_LINE_BREAK_NAMES,
  LIST_VERTICAL_ALIGN_NAMES,
  PAGE_BINDING_NAMES,
  COLUMN_TYPE_NAMES,
  COLUMN_DIRECTION_NAMES,
  TAB_ALIGN_NAMES,
  TAB_LEADER_NAMES,
  BIN_STORAGE_TYPE_NAMES,
  BIN_COMPRESSION_NAMES,
  CONTROL_CHAR_SIZE,
  IMAGE_MIME_BY_EXT,
  SHAPE_OBJECT_CTRL_NAMES,
  GRAPHIC_COMMON_RECORD_TAGS,
  GRAPHIC_DETAIL_RECORD_TAGS,
  OBJECT_TEXT_FLOW_NAMES,
  OBJECT_TEXT_SIDE_NAMES,
  OBJECT_VERT_REL_TO_NAMES,
  OBJECT_HORZ_REL_TO_NAMES,
  OBJECT_VERT_ALIGN_NAMES,
  OBJECT_HORZ_ALIGN_NAMES,
  PREVIEW_DEBUG,
  PREVIEW_FONT_SCALE,
} from "./constants/hwp.js";

const state = {
  doc: null,
  selectedEntryIndex: null,
  selectedRecordIndex: null,
};

const app = document.querySelector("#app");
const { fileInput, summaryEl, docInfoPanelEl, previewPanelEl, streamListEl, recordPanelEl, detailPanelEl } =
  renderAppShell(app);

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    await loadHwp(file);
    render();
  } catch (error) {
    releaseDocResources(state.doc);
    state.doc = null;
    state.selectedEntryIndex = null;
    state.selectedRecordIndex = null;
    summaryEl.innerHTML = `<div class="error">Failed to parse file: ${escapeHtml(error.message)}</div>`;
    docInfoPanelEl.innerHTML = `<div class="empty">No DocInfo data.</div>`;
    previewPanelEl.innerHTML = `<div class="empty">No preview data.</div>`;
    streamListEl.innerHTML = "";
    recordPanelEl.innerHTML = `<div class="empty">No data.</div>`;
    detailPanelEl.innerHTML = `<div class="empty">No data.</div>`;
  }
});

async function loadHwp(file) {
  const previousDoc = state.doc;
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const { doc, selectedEntryIndex } = buildDocumentFromBytes({
    fileName: file.name,
    bytes,
    parseDocInfo,
    buildBinDataCatalog,
  });

  releaseDocResources(previousDoc);
  state.doc = doc;
  state.selectedEntryIndex = selectedEntryIndex;
  state.selectedRecordIndex = null;
}

function render() {
  renderSummaryPanel(summaryEl, state.doc);
  renderDocInfoPanel();
  renderStreamList();
  renderRecordPanel();
  renderPreviewPanel();
}

function renderDocInfoPanel() {
  renderDocInfoPanelContent(docInfoPanelEl, state.doc, {
    formatStyleName,
    formatParaShapeSummary,
    describeDocCharShape,
    colorRefToText,
  });
}

function renderPreviewPanel() {
  const model = buildPreviewModel();
  if (!model || !model.paragraphs.length) {
    const mode = state.doc?.formatInfo?.mode ?? "unknown";
    if (mode === "legacy-3x-binary" || mode === "hwpml-xml") {
      previewPanelEl.innerHTML = `<div class="empty">Preview is limited for this format profile (${escapeHtml(
        state.doc?.formatInfo?.label || mode
      )}). Use stream/detail panes for raw inspection.</div>`;
      return;
    }
    previewPanelEl.innerHTML = `<div class="empty">No BodyText paragraphs available for preview.</div>`;
    return;
  }

  const pages = buildPreviewPagesWithOverflow(model.paragraphs, model.pageLayout);
  const sheetClass = PREVIEW_DEBUG ? "preview-sheet preview-sheet-debug" : "preview-sheet preview-sheet-doc";
  const pageStyle = buildPreviewSheetStyle(model.pageLayout);

  const renderedPages = pages
    .map((page) => {
      const pageParagraphs = page.paragraphs;
      const segments = buildPreviewSegments(pageParagraphs);
      const segmentItems = segments
        .map((segment) => {
          const segmentBody = segment.paragraphs.map((item) => renderPreviewParagraph(item)).filter(Boolean).join("");
          if (!segmentBody) {
            return "";
          }

          const segmentFlags = [];
          if (segment.columnStart) {
            segmentFlags.push("column-start");
          }
          const columnInfo =
            PREVIEW_DEBUG && (segment.columnLayout?.count > 1 || segmentFlags.length)
              ? `<div class="preview-flow-meta">
                  ${segment.columnLayout?.count > 1 ? `<span>${escapeHtml(`columns ${segment.columnLayout.count}`)}</span>` : ""}
                  ${segment.columnLayout?.count > 1 ? `<span>${escapeHtml(`gap ${segment.columnLayout.gap}`)}</span>` : ""}
                  ${segment.columnLayout?.count > 1 ? `<span>${escapeHtml(segment.columnLayout.directionName)}</span>` : ""}
                  ${segmentFlags.map((flag) => `<span>${escapeHtml(flag)}</span>`).join("")}
                </div>`
              : "";

          const flowClass = segment.columnStart ? "preview-flow preview-flow-column-start" : "preview-flow";
          return `
            <section class="${flowClass}">
              ${columnInfo}
              <div class="preview-flow-body" style="${escapeHtmlAttr(buildPreviewFlowStyle(segment.columnLayout))}">
                ${segmentBody}
              </div>
            </section>
          `;
        })
        .filter(Boolean)
        .join("");

      const pageHead = PREVIEW_DEBUG
        ? `<div class="preview-page-head">
            <span>${escapeHtml(`page ${page.index}`)}</span>
            <span>${escapeHtml(`paragraphs ${page.paragraphs.length}`)}</span>
          </div>`
        : "";

      if (!segmentItems) {
        return null;
      }

      return `
        <article class="${sheetClass} preview-page" style="${escapeHtmlAttr(pageStyle)}">
          ${pageHead}
          ${segmentItems}
        </article>
      `;
    })
    .filter(Boolean);
  const items = renderedPages.join("");

  const chips = [
    `<span class="chip">sections ${model.sectionCount}</span>`,
    `<span class="chip">paragraphs ${model.paragraphs.length}</span>`,
    `<span class="chip">pages ${renderedPages.length}</span>`,
  ];
  if (model.columnLayout?.count > 1) {
    chips.push(`<span class="chip">columns ${model.columnLayout.count}</span>`);
  }
  if (model.pageLayout) {
    chips.push(
      `<span class="chip">${escapeHtml(
        `${model.pageLayout.orientation} ${Math.round(model.pageLayout.paperWidthMm)}x${Math.round(model.pageLayout.paperHeightMm)}mm`
      )}</span>`
    );
  }
  if (PREVIEW_DEBUG) {
    chips.push(`<span class="chip">controls ${model.controlCount}</span>`);
    chips.push(`<span class="chip">debug</span>`);
  } else {
    chips.push(`<span class="chip">reader mode</span>`);
  }

  previewPanelEl.innerHTML = `
    <div class="preview-meta">${chips.join("")}</div>
    <div class="preview-pages">${items}</div>
  `;
}

function resolveGraphicObjectDimensions(graphicInfo) {
  const objectWidth = Number.isFinite(graphicInfo?.objectCommon?.width) && graphicInfo.objectCommon.width > 0 ? graphicInfo.objectCommon.width : 0;
  const objectHeight =
    Number.isFinite(graphicInfo?.objectCommon?.height) && graphicInfo.objectCommon.height > 0 ? graphicInfo.objectCommon.height : 0;
  const componentWidth =
    Number.isFinite(graphicInfo?.shapeComponent?.currentWidth) && graphicInfo.shapeComponent.currentWidth > 0
      ? graphicInfo.shapeComponent.currentWidth
      : 0;
  const componentHeight =
    Number.isFinite(graphicInfo?.shapeComponent?.currentHeight) && graphicInfo.shapeComponent.currentHeight > 0
      ? graphicInfo.shapeComponent.currentHeight
      : 0;

  if (objectWidth > 0 && objectHeight > 0) {
    return {
      width: objectWidth,
      height: objectHeight,
    };
  }

  if (objectWidth > 0 || objectHeight > 0) {
    return {
      width: objectWidth > 0 ? objectWidth : componentWidth,
      height: objectHeight > 0 ? objectHeight : componentHeight,
    };
  }

  return {
    width: Math.max(componentWidth, 0),
    height: Math.max(componentHeight, 0),
  };
}

function computeControlRenderHints(controls) {
  if (!controls?.length) {
    return [];
  }

  const hints = controls.map(() => ({
    topGapPx: 0,
    shiftXPx: 0,
    absolute: false,
    absoluteTopPx: 0,
    absoluteXPx: 0,
    absoluteAlign: "left",
    absoluteFlow: null,
  }));
  let previousBottomY = null;

  for (let i = 0; i < controls.length; i += 1) {
    const control = controls[i];
    if (control?.ctrlId !== "gso " || !control.graphicInfo?.objectCommon) {
      continue;
    }

    const object = control.graphicInfo.objectCommon;
    const yOffset = Number.isFinite(object.yOffset) ? object.yOffset : null;
    const xOffset = Number.isFinite(object.xOffset) ? object.xOffset : 0;
    const { height } = resolveGraphicObjectDimensions(control.graphicInfo);
    const allowOffsetLayout = !object.propertyBits?.likeCharacter;
    const flow = object.propertyBits?.textFlow;
    const isOverlayFlow = flow === 2 || flow === 3;
    const hasAbsoluteAnchor =
      object.propertyBits?.vertRelToName &&
      object.propertyBits?.horzRelToName &&
      (object.propertyBits.vertRelToName === "paper" || object.propertyBits.vertRelToName === "page") &&
      (object.propertyBits.horzRelToName === "paper" || object.propertyBits.horzRelToName === "page");

    if (allowOffsetLayout && isOverlayFlow && hasAbsoluteAnchor && yOffset != null) {
      hints[i].absolute = true;
      hints[i].absoluteTopPx = clamp(hwpToPx(yOffset), -240, 2400);
      hints[i].absoluteXPx = clamp(hwpToPx(xOffset), -720, 1800);
      hints[i].absoluteAlign = object.propertyBits?.horzAlignName || "left";
      hints[i].absoluteFlow = flow;
      continue;
    }

    if (allowOffsetLayout && yOffset != null) {
      const gapY = Number.isFinite(previousBottomY) ? Math.max(0, yOffset - previousBottomY) : Math.max(0, yOffset);
      hints[i].topGapPx = clamp(hwpToPx(gapY), 0, 1500);
      const bottomY = yOffset + Math.max(0, height || 0);
      previousBottomY = Number.isFinite(previousBottomY) ? Math.max(previousBottomY, bottomY) : bottomY;
    }
    hints[i].shiftXPx = allowOffsetLayout ? clamp(hwpToPx(xOffset), -320, 320) : 0;
  }

  return hints;
}

function renderPreviewControlBlock(control, renderHint = null) {
  if (!isRenderableControl(control)) {
    return "";
  }

  const previewBody = control.tableInfo
    ? renderTableControlPreview(control.tableInfo, control.tableSplitInfo ?? null)
    : control.graphicInfo
      ? renderGraphicControlPreview(control.graphicInfo)
      : control.sectionInfo
        ? renderSectionControlPreview(control.sectionInfo)
        : control.columnInfo
          ? renderColumnControlPreview(control.columnInfo)
      : "";

  if (!PREVIEW_DEBUG) {
    if (previewBody) {
      const inlineStyles = [];
      if (renderHint?.absolute) {
        inlineStyles.push("position:absolute");
        inlineStyles.push(`top:${renderHint.absoluteTopPx.toFixed(1)}px`);
        if (renderHint.absoluteAlign === "center") {
          inlineStyles.push(`left:calc(50% + ${renderHint.absoluteXPx.toFixed(1)}px)`);
          inlineStyles.push("transform:translateX(-50%)");
        } else if (renderHint.absoluteAlign === "right") {
          inlineStyles.push(`right:${Math.max(0, renderHint.absoluteXPx).toFixed(1)}px`);
        } else {
          inlineStyles.push(`left:${renderHint.absoluteXPx.toFixed(1)}px`);
        }
        inlineStyles.push(`z-index:${renderHint.absoluteFlow === 2 ? 0 : 3}`);
        inlineStyles.push("margin-top:0");
      } else {
        if (renderHint?.topGapPx >= 0.2) {
          inlineStyles.push(`margin-top:${renderHint.topGapPx.toFixed(1)}px`);
        }
        if (Number.isFinite(renderHint?.shiftXPx) && Math.abs(renderHint.shiftXPx) >= 0.2) {
          inlineStyles.push("position:relative");
          inlineStyles.push(`left:${renderHint.shiftXPx.toFixed(1)}px`);
        }
      }
      const styleAttr = inlineStyles.length ? ` style="${escapeHtmlAttr(inlineStyles.join(";"))}"` : "";
      return `<article class="control-block doc-control kind-${escapeHtmlAttr(control.kind)}"${styleAttr}>${previewBody}</article>`;
    }
    if (control.ctrlId === "gso ") {
      return "";
    }
    if (control.ctrlId === "fn  " || control.ctrlId === "en  ") {
      return `<span class="doc-note-marker">[${escapeHtml(control.name)}]</span>`;
    }
    if (control.name === "Unknown Control" || !control.ctrlId?.trim()) {
      return "";
    }
    return `<span class="doc-inline-control">${escapeHtml(control.name)}</span>`;
  }

  const kindName = CONTROL_KIND_NAMES[control.kind] ?? CONTROL_KIND_NAMES.inline;
  const metrics = [
    `id ${control.ctrlId || "?"}`,
    `level ${control.level}`,
    `payload ${control.payloadSize}B`,
    `sub ${control.subRecordCount}`,
  ];
  if (control.inlineCount > 0) {
    metrics.push(`inline ${control.inlineCount}`);
  }
  if (control.recordIndex != null) {
    metrics.push(`record #${control.recordIndex}`);
  }

  const tags = control.tagSummary?.length
    ? `<div class="control-block-tags">
        ${control.tagSummary
          .map((tag) => `<span class="control-block-tag">${escapeHtml(tag)}</span>`)
          .join("")}
      </div>`
    : "";

  const details = control.detailLines?.length
    ? `<pre class="control-block-data">${escapeHtml(control.detailLines.join("\n"))}</pre>`
    : "";

  return `
    <article class="control-block kind-${escapeHtmlAttr(control.kind)}">
      <div class="control-block-head">
        <span class="control-block-code">${escapeHtml(control.ctrlId || "?")}</span>
        <strong>${escapeHtml(control.name)}</strong>
        <span class="control-block-kind">${escapeHtml(kindName)}</span>
      </div>
      <div class="control-block-metrics">
        ${metrics.map((value) => `<span>${escapeHtml(value)}</span>`).join("")}
      </div>
      ${tags}
      ${previewBody}
      ${details}
    </article>
  `;
}

function isRenderableControl(control) {
  if (PREVIEW_DEBUG) {
    return true;
  }
  if (!control) {
    return false;
  }
  if (control.ctrlId === "secd" || control.ctrlId === "cold" || control.ctrlId === "head" || control.ctrlId === "foot") {
    return false;
  }
  if (control.ctrlId === "atno" || control.ctrlId === "bokm" || control.ctrlId === " %d") {
    return false;
  }
  return true;
}

function renderPreviewParagraph(item) {
  const useLineLayout = (item.textLines?.length ?? 0) > 0;
  const textRenderOptions = {
    listMarker: item.listMarker,
    listType: item.listType,
    tabStops: item.tabStops,
    tabStopsPx: item.tabStopsPx,
    baseStyle: item.style,
  };
  const textCss = [
    `text-align:${item.style.textAlign}`,
    `font-family:${item.style.fontFamily}`,
    `font-size:${item.style.fontSizePt.toFixed(1)}pt`,
    `font-style:${item.style.fontStyle}`,
    `font-weight:${item.style.fontWeight}`,
    `color:${item.style.color}`,
    `line-height:${item.style.lineHeight}`,
    `margin:${item.style.spacingBeforePx.toFixed(1)}px ${item.style.marginRightPx.toFixed(1)}px ${item.style.spacingAfterPx.toFixed(
      1
    )}px ${item.style.marginLeftPx.toFixed(1)}px`,
    `text-indent:${(useLineLayout ? 0 : item.style.textIndentPx).toFixed(1)}px`,
  ].join(";");

  const hasText = (item.previewText ?? "").length > 0 || useLineLayout;
  const textHtml = useLineLayout
    ? renderStyledTextLines(item.textLines, textRenderOptions)
    : item.textRuns?.length
      ? (() => {
          const marker = renderListMarker(item.listMarker, item.listType);
          const body = renderStyledTextRuns(item.textRuns, {
            ...textRenderOptions,
            initialCursorPx: estimateListMarkerAdvancePx(item.listMarker, item.listType, item.style),
          });
          return marker ? `${marker}${body}` : body;
        })()
      : (() => {
          const marker = renderListMarker(item.listMarker, item.listType);
          const body = textToHtml(item.previewText);
          return marker ? `${marker}${body}` : body;
        })();
  const controlRenderHints = computeControlRenderHints(item.controls);
  const renderedControls = item.controls
    .map((ctrl, index) => renderPreviewControlBlock(ctrl, controlRenderHints[index] ?? null))
    .filter(Boolean)
    .join("");
  if (!hasText && !renderedControls) {
    return "";
  }

  if (!hasText && renderedControls && !PREVIEW_DEBUG) {
    return `
      <article class="preview-paragraph preview-control-paragraph">
        ${renderedControls}
      </article>
    `;
  }

  return `
    <article class="preview-paragraph">
      ${
        hasText
          ? `<p class="preview-text ${useLineLayout ? "preview-text-lines" : ""}" style="${escapeHtmlAttr(textCss)}">${textHtml}</p>`
          : PREVIEW_DEBUG
            ? `<p class="preview-text-empty">(control-only paragraph)</p>`
            : ""
      }
      ${renderedControls ? `<div class="preview-control-stack">${renderedControls}</div>` : ""}
      ${
        PREVIEW_DEBUG
          ? `<div class="preview-note">
              <span>#${item.previewIndex + 1}</span>
              <span>${escapeHtml(item.sectionLabel)}</span>
              <span>paraShape=${item.refs.paraShapeId ?? "-"}</span>
              <span>style=${item.refs.paraStyleId ?? "-"}</span>
              <span>charShape=${item.refs.charShapeId ?? "-"}</span>
              <span>controls=${item.controls.length}</span>
            </div>`
          : ""
      }
    </article>
  `;
}

function renderStyledTextRuns(textRuns, options = {}) {
  const tabStops = normalizeTabStops(options.tabStops, options.tabStopsPx);
  const tabStopsPx = tabStops.map((tab) => tab.positionPx);
  const baseFontSizePt = clamp(options.baseStyle?.fontSizePt ?? 10, 6, 72);
  const defaultTabPx = estimateDefaultTabWidthPx(baseFontSizePt);
  const out = [];
  let cursorPx = Number.isFinite(options.initialCursorPx) ? options.initialCursorPx : 0;

  for (const run of textRuns ?? []) {
    if (run?.tab) {
      const tabInfo = computeNextTabAdvance(cursorPx, tabStops, defaultTabPx);
      const tabWidth = tabInfo.advancePx;
      const leaderHtml = buildTabLeaderHtml(tabWidth, baseFontSizePt, tabInfo.stop?.leaderName);
      out.push(`<span class="doc-tab-stop" style="min-width:${tabWidth.toFixed(1)}px">${leaderHtml}</span>`);
      cursorPx += tabWidth;
      continue;
    }

    const text = normalizeDisplayText(run?.text ?? "");
    if (!text) {
      continue;
    }
    const html = textToHtml(text);
    if (!run.css) {
      out.push(html);
    } else {
      out.push(`<span style="${escapeHtmlAttr(run.css)}">${html}</span>`);
    }

    const runFontSizePt = resolveRunFontSizePt(run.css, baseFontSizePt);
    cursorPx = estimateTextCursorPx(text, cursorPx, runFontSizePt, tabStopsPx, defaultTabPx);
  }

  return out.join("");
}

function normalizeTabStops(tabStops, tabStopsPx) {
  if (Array.isArray(tabStops) && tabStops.length) {
    return tabStops
      .filter((tab) => Number.isFinite(tab?.positionPx) && tab.positionPx > 0)
      .map((tab) => ({
        positionPx: tab.positionPx,
        alignName: tab.alignName ?? TAB_ALIGN_NAMES[tab.align] ?? "left",
        leaderName: tab.leaderName ?? TAB_LEADER_NAMES[tab.leader] ?? "none",
      }))
      .sort((a, b) => a.positionPx - b.positionPx);
  }
  return [...(tabStopsPx ?? [])]
    .filter((value) => Number.isFinite(value) && value > 0)
    .sort((a, b) => a - b)
    .map((positionPx) => ({
      positionPx,
      alignName: "left",
      leaderName: "none",
    }));
}

function computeNextTabAdvance(cursorPx, tabStops, defaultTabPx) {
  for (const tabStop of tabStops) {
    if (tabStop.positionPx > cursorPx + 0.2) {
      return {
        advancePx: clamp(tabStop.positionPx - cursorPx, 8, 420),
        stop: tabStop,
      };
    }
  }
  const step = Math.max(8, defaultTabPx);
  const next = Math.floor(cursorPx / step + 1) * step;
  return {
    advancePx: clamp(next - cursorPx, 8, 420),
    stop: null,
  };
}

function buildTabLeaderHtml(tabWidthPx, fontSizePt, leaderName) {
  const leaderKind = tabLeaderKind(leaderName);
  if (!leaderKind) {
    return "";
  }
  const fontPx = clamp((fontSizePt * 96) / 72, 8, 120);
  const heightEm = clamp(fontPx / 18, 0.5, 1.1).toFixed(2);
  return `<span class="doc-tab-leader doc-tab-leader-${leaderKind}" style="height:${heightEm}em"></span>`;
}

function tabLeaderKind(leaderName) {
  if (leaderName === "dot" || leaderName === "middle-dot") {
    return "dot";
  }
  if (leaderName === "hyphen") {
    return "dash";
  }
  if (leaderName === "underline") {
    return "underline";
  }
  return null;
}

function renderStyledTextLines(textLines, options = {}) {
  const markerHtml = renderListMarker(options.listMarker, options.listType);
  return (textLines ?? [])
    .map((line, index) => {
      const normalizedRuns = normalizeTabRunsForLine(line.runs);
      const hasTabRuns = normalizedRuns.some((run) => Boolean(run?.tab));
      const lineOptions =
        index === 0 && markerHtml
          ? {
              ...options,
              initialCursorPx: estimateListMarkerAdvancePx(options.listMarker, options.listType, options.baseStyle),
            }
          : options;
      const textHtml = normalizedRuns?.length ? renderStyledTextRuns(normalizedRuns, lineOptions) : "&nbsp;";
      const mergedText = index === 0 && markerHtml ? `${markerHtml}${textHtml}` : textHtml;
      const lineClass = hasTabRuns ? "preview-line preview-line-tabbed" : "preview-line";
      const attrs = line.css ? ` style="${escapeHtmlAttr(line.css)}"` : "";
      return `<span class="${lineClass}"${attrs}>${mergedText}</span>`;
    })
    .join("");
}

function normalizeTabRunsForLine(runs) {
  if (!Array.isArray(runs) || !runs.length) {
    return [];
  }
  const tabIndexes = [];
  for (let i = 0; i < runs.length; i += 1) {
    if (runs[i]?.tab) {
      tabIndexes.push(i);
    }
  }
  if (tabIndexes.length <= 1) {
    return runs;
  }
  const keepIndex = tabIndexes[tabIndexes.length - 1];
  return runs.map((run, index) => {
    if (!run?.tab || index === keepIndex) {
      return run;
    }
    // Older documents may emit multiple tab controls inside TOC labels.
    // Keep only the final leader tab and normalize earlier tabs to spaces.
    return {
      text: " ",
      css: "",
    };
  });
}

function renderListMarker(marker, listType = "") {
  if (!marker) {
    return "";
  }
  const markerClass = listType ? ` preview-list-marker-${escapeHtmlAttr(listType)}` : "";
  return `<span class="preview-list-marker${markerClass}">${escapeHtml(normalizeDisplayText(marker))}</span>`;
}

function estimateListMarkerAdvancePx(marker, listType, baseStyle) {
  if (!marker) {
    return 0;
  }
  const fontSizePt = clamp(baseStyle?.fontSizePt ?? 10, 6, 72);
  const fontPx = (fontSizePt * 96) / 72;
  const markerUnits =
    listType === "bullet"
      ? 1.35
      : Math.max(2.3, Math.min(4.6, String(marker).length * 0.58 + 0.9));
  const marginUnits = 0.34;
  return clamp(fontPx * (markerUnits + marginUnits), 10, 320);
}

function estimateDefaultTabWidthPx(fontSizePt) {
  const fontPx = clamp((fontSizePt * 96) / 72, 8, 96);
  return clamp(fontPx * 4, 20, 220);
}

function computeTabAdvancePx(cursorPx, tabStopsPx, defaultTabPx) {
  for (const tabStop of tabStopsPx) {
    if (tabStop > cursorPx + 0.2) {
      return clamp(tabStop - cursorPx, 8, 420);
    }
  }
  const step = Math.max(8, defaultTabPx);
  const next = Math.floor(cursorPx / step + 1) * step;
  return clamp(next - cursorPx, 8, 420);
}

function resolveRunFontSizePt(css, fallbackPt) {
  if (!css) {
    return fallbackPt;
  }
  const match = css.match(/font-size:([0-9.]+)pt/);
  if (!match) {
    return fallbackPt;
  }
  const value = Number(match[1]);
  return Number.isFinite(value) ? clamp(value, 6, 96) : fallbackPt;
}

function estimateTextCursorPx(text, startPx, fontSizePt, tabStopsPx, defaultTabPx) {
  const fontPx = clamp((fontSizePt * 96) / 72, 8, 120);
  let cursorPx = Number.isFinite(startPx) ? startPx : 0;
  for (const char of text) {
    if (char === "\n") {
      cursorPx = 0;
      continue;
    }
    if (char === "\t") {
      cursorPx += computeTabAdvancePx(cursorPx, tabStopsPx, defaultTabPx);
      continue;
    }
    if (char === " ") {
      cursorPx += fontPx * 0.32;
      continue;
    }
    const code = char.charCodeAt(0);
    const wide =
      (code >= 0x1100 && code <= 0x11ff) ||
      (code >= 0x3130 && code <= 0x318f) ||
      (code >= 0xac00 && code <= 0xd7a3) ||
      (code >= 0x3000 && code <= 0x303f) ||
      (code >= 0x4e00 && code <= 0x9fff);
    cursorPx += fontPx * (wide ? 0.92 : 0.56);
  }
  return cursorPx;
}

function tokensToPlainText(tokens) {
  return normalizeDisplayText((tokens ?? []).map((token) => token.text ?? "").join(""));
}

function buildStyledTextRuns(tokens, charRuns, docInfo, baseStyle, fallbackCharShapeId = null) {
  if (!tokens?.length) {
    return [];
  }

  const sortedRuns = [...(charRuns ?? [])]
    .filter((run) => Number.isFinite(run.startPos) && run.startPos >= 0)
    .sort((a, b) => a.startPos - b.startPos);

  let runIndex = -1;
  let currentShapeId = fallbackCharShapeId;
  if (sortedRuns.length && sortedRuns[0].startPos <= 0) {
    runIndex = 0;
    currentShapeId = sortedRuns[0].charShapeId;
  }

  const cssCache = new Map();
  const out = [];

  for (const token of tokens) {
    const logicalPos = Number.isFinite(token.logicalPos) ? token.logicalPos : 0;
    while (runIndex + 1 < sortedRuns.length && logicalPos >= sortedRuns[runIndex + 1].startPos) {
      runIndex += 1;
      currentShapeId = sortedRuns[runIndex].charShapeId;
    }

    const css = resolveCharShapeInlineCss(currentShapeId, docInfo, baseStyle, cssCache);
    if (token.type === "tab") {
      out.push({
        tab: true,
        css,
      });
      continue;
    }
    const text = token.text ?? "";
    if (!text) {
      continue;
    }

    const last = out[out.length - 1];
    if (last && last.css === css) {
      last.text += text;
    } else {
      out.push({ text, css });
    }
  }

  return out;
}

function buildStyledTextLines(tokens, charRuns, lineSegments, docInfo, baseStyle, fallbackCharShapeId = null) {
  if (!tokens?.length) {
    return [];
  }

  const orderedTokens = [...tokens].sort((a, b) => {
    const aPos = Number.isFinite(a.logicalPos) ? a.logicalPos : 0;
    const bPos = Number.isFinite(b.logicalPos) ? b.logicalPos : 0;
    return aPos - bPos;
  });
  if (!orderedTokens.length) {
    return [];
  }

  const sortedSegments = [...(lineSegments ?? [])]
    .filter((segment) => Number.isFinite(segment.textStartPos))
    .sort((a, b) => a.textStartPos - b.textStartPos);
  const hasSegmentInfo = sortedSegments.length > 0;
  const baseHoriz = hasSegmentInfo ? sortedSegments[0].horizStartPos ?? 0 : 0;

  let segmentIndex = hasSegmentInfo ? 0 : -1;
  let activeSegment = segmentIndex >= 0 ? sortedSegments[segmentIndex] : null;
  let currentTokens = [];
  const lines = [];

  const pushLine = (force = false) => {
    if (!force && currentTokens.length === 0) {
      return;
    }
    const tokensForLine = currentTokens;
    currentTokens = [];
    lines.push({
      runs: buildStyledTextRuns(tokensForLine, charRuns, docInfo, baseStyle, fallbackCharShapeId),
      css: buildLineInlineCss(activeSegment, baseHoriz, baseStyle),
      estimatedHeightPx: estimateLineHeightPx(activeSegment, baseStyle),
      pageStart: Boolean((activeSegment?.flags ?? 0) & LINE_SEG_PAGE_FIRST_BIT),
      columnStart: Boolean((activeSegment?.flags ?? 0) & LINE_SEG_COLUMN_FIRST_BIT),
    });
  };

  for (const token of orderedTokens) {
    const logicalPos = Number.isFinite(token.logicalPos) ? token.logicalPos : 0;
    while (segmentIndex >= 0 && segmentIndex + 1 < sortedSegments.length && logicalPos >= sortedSegments[segmentIndex + 1].textStartPos) {
      pushLine(false);
      segmentIndex += 1;
      activeSegment = sortedSegments[segmentIndex];
    }

    if (token.type === "newline") {
      pushLine(true);
      continue;
    }
    currentTokens.push(token);
  }
  pushLine(false);

  if (!lines.length) {
    return [];
  }
  return lines;
}

function estimateLineHeightPx(segment, baseStyle) {
  const lineHeightRatio = Number.parseFloat(baseStyle?.lineHeight ?? "1.6");
  const ratio = Number.isFinite(lineHeightRatio) ? clamp(lineHeightRatio, 1.0, 2.6) : 1.6;
  const baseFontPx = clamp((baseStyle?.fontSizePt ?? 10) * (96 / 72), 8, 96);
  let result = baseFontPx * ratio;

  if (segment) {
    const minHeightUnit = segment.textHeight > 0 ? segment.textHeight : segment.lineHeight;
    if (Number.isFinite(minHeightUnit) && minHeightUnit > 0) {
      result = Math.max(result, clamp(hwpToPx(minHeightUnit), 10, 200));
    }
    if (Number.isFinite(segment.lineSpaceBelow) && segment.lineSpaceBelow > 0) {
      result += clamp(hwpToPx(segment.lineSpaceBelow), 0, 72);
    }
  }

  return clamp(result, 10, 240);
}

function buildLineInlineCss(segment, baseHoriz, baseStyle) {
  if (!segment) {
    return "";
  }

  const styles = [];
  const lineShiftPx = clamp(hwpToPx((segment.horizStartPos ?? 0) - baseHoriz), -260, 380);
  if (Math.abs(lineShiftPx) >= 0.2) {
    styles.push(`margin-left:${lineShiftPx.toFixed(1)}px`);
  }

  const minHeightUnit = segment.textHeight > 0 ? segment.textHeight : segment.lineHeight;
  if (Number.isFinite(minHeightUnit) && minHeightUnit > 0) {
    const minHeightPx = clamp(hwpToPx(minHeightUnit), Math.max(10, baseStyle.fontSizePt * 0.9), 180);
    styles.push(`min-height:${minHeightPx.toFixed(1)}px`);
  }

  const spaceBelowPx = Number.isFinite(segment.lineSpaceBelow) ? clamp(hwpToPx(segment.lineSpaceBelow), 0, 72) : 0;
  if (spaceBelowPx >= 0.2) {
    styles.push(`padding-bottom:${spaceBelowPx.toFixed(1)}px`);
  }

  return styles.join(";");
}

function resolveCharShapeInlineCss(charShapeId, docInfo, baseStyle, cache) {
  const key = charShapeId == null ? "null" : String(charShapeId);
  if (cache.has(key)) {
    return cache.get(key);
  }

  const shape = charShapeId != null ? docInfo?.charShapeById?.get(charShapeId) ?? null : null;
  if (!shape) {
    cache.set(key, "");
    return "";
  }

  const parts = [];
  const fontFamily = toCssFontFamily(shape.primaryFont);
  if (fontFamily && fontFamily !== baseStyle.fontFamily) {
    parts.push(`font-family:${fontFamily}`);
  }

  const sizePt =
    shape.baseSize != null && shape.baseSize > 0
      ? clamp((shape.baseSize / 100) * PREVIEW_FONT_SCALE, 6, 72)
      : baseStyle.fontSizePt;
  if (Math.abs(sizePt - baseStyle.fontSizePt) >= 0.2) {
    parts.push(`font-size:${sizePt.toFixed(1)}pt`);
  }

  const fontWeight = shape.attributeBits?.bold ? 600 : 400;
  if (fontWeight !== baseStyle.fontWeight) {
    parts.push(`font-weight:${fontWeight}`);
  }

  const fontStyle = shape.attributeBits?.italic ? "italic" : "normal";
  if (fontStyle !== baseStyle.fontStyle) {
    parts.push(`font-style:${fontStyle}`);
  }

  const color = !shape.textColor || shape.textColor.auto ? baseStyle.color : shape.textColor.hex;
  if (color !== baseStyle.color) {
    parts.push(`color:${color}`);
  }

  const decorationLines = [];
  if (shape.attributeBits?.underlineType && shape.attributeBits.underlineType !== 0) {
    decorationLines.push("underline");
  }
  if (shape.attributeBits?.strikeoutType && shape.attributeBits.strikeoutType !== 0) {
    decorationLines.push("line-through");
  }
  if (decorationLines.length) {
    parts.push(`text-decoration-line:${decorationLines.join(" ")}`);
  }

  if (decorationLines.includes("underline") && shape.underlineColor && !shape.underlineColor.auto) {
    parts.push(`text-decoration-color:${shape.underlineColor.hex}`);
  }

  if (shape.attributeBits?.superscript) {
    parts.push("vertical-align:super");
  } else if (shape.attributeBits?.subscript) {
    parts.push("vertical-align:sub");
  }

  const css = parts.join(";");
  cache.set(key, css);
  return css;
}

function buildPreviewSheetStyle(pageLayout) {
  if (!pageLayout) {
    return "";
  }
  const styles = [
    `width:${pageLayout.paperWidthPx.toFixed(1)}px`,
    `height:${pageLayout.paperHeightPx.toFixed(1)}px`,
    `min-height:${pageLayout.paperHeightPx.toFixed(1)}px`,
    `box-sizing:border-box`,
    `margin:0 auto`,
    `padding:${pageLayout.topPaddingPx.toFixed(1)}px ${pageLayout.rightPaddingPx.toFixed(1)}px ${pageLayout.bottomPaddingPx.toFixed(
      1
    )}px ${pageLayout.leftPaddingPx.toFixed(1)}px`,
  ];
  return styles.join(";");
}

function buildPreviewSegments(paragraphs) {
  const segments = [];
  for (const paragraph of paragraphs) {
    const key = columnLayoutKey(paragraph.columnLayout);
    const current = segments[segments.length - 1];
    const forceColumnSplit = Boolean(paragraph.breakHints?.columnStart) && Boolean(current?.paragraphs?.length);
    if (!current || current.key !== key || forceColumnSplit) {
      segments.push({
        key,
        columnLayout: paragraph.columnLayout ?? null,
        columnStart: Boolean(paragraph.breakHints?.columnStart),
        paragraphs: [paragraph],
      });
      continue;
    }
    current.paragraphs.push(paragraph);
  }
  return segments;
}

function buildPreviewPages(paragraphs) {
  const pages = [];
  let current = null;

  for (let i = 0; i < paragraphs.length; i += 1) {
    const paragraph = { ...paragraphs[i], previewIndex: i };
    const shouldStartNewPage = Boolean(paragraph.breakHints?.pageStart) && Boolean(current?.paragraphs?.length);
    if (!current || shouldStartNewPage) {
      current = {
        index: pages.length + 1,
        paragraphs: [],
      };
      pages.push(current);
    }
    current.paragraphs.push(paragraph);
  }
  return pages;
}

function buildPreviewPagesWithOverflow(paragraphs, pageLayout) {
  const basePages = trimLeadingEmptyPages(buildPreviewPages(paragraphs));
  if (!basePages.length) {
    return [];
  }
  const contentHeightPx = resolvePageContentHeightPx(pageLayout, paragraphs);
  const paged = [];
  for (const basePage of basePages) {
    const overflowPages = paginatePageParagraphsByHeight(basePage.paragraphs, contentHeightPx);
    for (const page of overflowPages) {
      paged.push({
        index: paged.length + 1,
        paragraphs: page.paragraphs,
      });
    }
  }
  return paged;
}

function resolvePageContentHeightPx(pageLayout, paragraphs = []) {
  const tableRatio = estimateTableParagraphRatio(paragraphs);
  const tableDensityScale = computeTableDensityPageScale(tableRatio);
  let base = PAGE_CONTENT_FALLBACK_HEIGHT_PX;
  if (!pageLayout) {
    return clamp(base * tableDensityScale, 320, 3200);
  }
  if (Number.isFinite(pageLayout.paperHeightPx) && Number.isFinite(pageLayout.topPaddingPx) && Number.isFinite(pageLayout.bottomPaddingPx)) {
    base = clamp(pageLayout.paperHeightPx - pageLayout.topPaddingPx - pageLayout.bottomPaddingPx, 320, 3200);
    return clamp(base * tableDensityScale, 320, 3200);
  }
  return clamp(base * tableDensityScale, 320, 3200);
}

function estimateTableParagraphRatio(paragraphs) {
  if (!paragraphs?.length) {
    return 0;
  }
  let total = 0;
  let withTable = 0;
  for (const paragraph of paragraphs) {
    total += 1;
    if ((paragraph?.controls ?? []).some((control) => Boolean(control?.tableInfo))) {
      withTable += 1;
    }
  }
  if (!total) {
    return 0;
  }
  return withTable / total;
}

function computeTableDensityPageScale(tableParagraphRatio) {
  if (!Number.isFinite(tableParagraphRatio)) {
    return 1;
  }
  const pivot = 0.22;
  if (tableParagraphRatio >= pivot) {
    return clamp(1 + (tableParagraphRatio - pivot) * 0.32, 0.96, 1.05);
  }
  return clamp(1 - (pivot - tableParagraphRatio) * 0.8, 0.96, 1.05);
}

function paginatePageParagraphsByHeight(paragraphs, contentHeightPx) {
  const state = {
    pages: [],
    current: {
      paragraphs: [],
      usedHeightPx: 0,
    },
    contentHeightPx,
  };

  for (const paragraph of paragraphs) {
    if (paragraph?.breakHints?.pageStart && state.current.paragraphs.length) {
      startNewPaginationPage(state);
    }
    appendParagraphWithOverflow(state, paragraph);
  }

  if (state.current.paragraphs.length) {
    state.pages.push(state.current);
  }

  return state.pages;
}

function startNewPaginationPage(state) {
  if (state.current.paragraphs.length) {
    state.pages.push(state.current);
  }
  state.current = {
    paragraphs: [],
    usedHeightPx: 0,
  };
}

function addPaginationParagraph(state, paragraph, estimatedHeightPx) {
  state.current.paragraphs.push(paragraph);
  const consume = Number.isFinite(estimatedHeightPx) ? clamp(estimatedHeightPx, 0, state.contentHeightPx) : PARAGRAPH_BLOCK_MIN_PX;
  state.current.usedHeightPx += consume;
}

function appendParagraphWithOverflow(state, paragraph) {
  if (!paragraph) {
    return;
  }
  const hasText = (paragraph.previewTextNormalized ?? "").length > 0 || (paragraph.textLines?.length ?? 0) > 0;
  const controls = paragraph.controls ?? [];

  if (!hasText && controls.length > 0) {
    appendControlOnlyParagraphWithOverflow(state, paragraph);
    return;
  }

  const estimatedHeight = estimateParagraphHeightPx(paragraph);
  if (
    estimatedHeight > 0 &&
    state.current.paragraphs.length &&
    state.current.usedHeightPx + estimatedHeight > state.contentHeightPx - PAGE_CONTENT_BOTTOM_GUARD_PX
  ) {
    startNewPaginationPage(state);
  }
  addPaginationParagraph(state, paragraph, estimatedHeight);
}

function appendControlOnlyParagraphWithOverflow(state, paragraph) {
  const controls = paragraph.controls ?? [];
  if (!controls.length) {
    addPaginationParagraph(state, paragraph, estimateParagraphHeightPx(paragraph));
    return;
  }

  if (controls.length === 1 && controls[0].tableInfo) {
    appendTableControlParagraphWithOverflow(state, paragraph, controls[0]);
    return;
  }

  controls.forEach((control, index) => {
    const breakHints = index === 0 ? paragraph.breakHints : { pageStart: false, columnStart: false };
    const sliced = cloneParagraphWithControls(paragraph, [control], breakHints);
    if (control.tableInfo) {
      appendTableControlParagraphWithOverflow(state, sliced, control);
      return;
    }
    const estimatedHeight = estimateParagraphHeightPx(sliced);
    if (
      estimatedHeight > 0 &&
      state.current.paragraphs.length &&
      state.current.usedHeightPx + estimatedHeight > state.contentHeightPx - PAGE_CONTENT_BOTTOM_GUARD_PX
    ) {
      startNewPaginationPage(state);
    }
    addPaginationParagraph(state, sliced, estimatedHeight);
  });
}

function appendTableControlParagraphWithOverflow(state, paragraph, tableControl) {
  const tableInfo = tableControl?.tableInfo;
  if (!tableInfo || !Number.isFinite(tableInfo.rows) || tableInfo.rows <= 0) {
    const estimatedHeight = estimateParagraphHeightPx(paragraph);
    if (
      estimatedHeight > 0 &&
      state.current.paragraphs.length &&
      state.current.usedHeightPx + estimatedHeight > state.contentHeightPx - PAGE_CONTENT_BOTTOM_GUARD_PX
    ) {
      startNewPaginationPage(state);
    }
    addPaginationParagraph(state, paragraph, estimatedHeight);
    return;
  }

  const rowHeights = estimateTableRowHeightsPx(tableInfo);
  let rowStart = 0;

  while (rowStart < tableInfo.rows) {
    const available = state.contentHeightPx - state.current.usedHeightPx - PAGE_CONTENT_BOTTOM_GUARD_PX;
    if (available < TABLE_CHUNK_BASE_PX + TABLE_ROW_MIN_PX && state.current.paragraphs.length) {
      startNewPaginationPage(state);
      continue;
    }

    let fitRows = pickTableRowsForHeight(rowHeights, rowStart, available);
    if (fitRows <= 0) {
      if (state.current.paragraphs.length) {
        startNewPaginationPage(state);
        continue;
      }
      fitRows = 1;
    }

    const rowEnd = Math.min(tableInfo.rows, rowStart + fitRows);
    const splitInfo =
      rowStart > 0 || rowEnd < tableInfo.rows
        ? {
            startRow: rowStart,
            endRow: rowEnd,
            totalRows: tableInfo.rows,
          }
        : null;
    const slicedTableInfo = splitInfo ? sliceTableInfoByRowRange(tableInfo, rowStart, rowEnd) : tableInfo;
    const slicedControl = {
      ...tableControl,
      tableInfo: slicedTableInfo,
      tableSplitInfo: splitInfo,
    };
    const breakHints = rowStart === 0 ? paragraph.breakHints : { pageStart: false, columnStart: false };
    const slicedParagraph = cloneParagraphWithControls(paragraph, [slicedControl], breakHints);
    const estimatedHeight = estimateParagraphHeightPx(slicedParagraph);

    if (
      estimatedHeight > 0 &&
      state.current.paragraphs.length &&
      state.current.usedHeightPx + estimatedHeight > state.contentHeightPx - PAGE_CONTENT_BOTTOM_GUARD_PX
    ) {
      startNewPaginationPage(state);
      continue;
    }

    addPaginationParagraph(state, slicedParagraph, estimatedHeight);
    rowStart = rowEnd;
  }
}

function cloneParagraphWithControls(paragraph, controls, breakHints = paragraph.breakHints) {
  return {
    ...paragraph,
    controls,
    breakHints,
  };
}

function estimateParagraphHeightPx(paragraph) {
  const textHeight = estimateParagraphTextHeightPx(paragraph);
  const controls = paragraph?.controls ?? [];
  const controlsHeight = controls.reduce((sum, control) => sum + estimateControlHeightPx(control), 0);
  const stackGap = controls.length > 1 ? (controls.length - 1) * 4 : 0;
  const stackTopGap = controls.length && textHeight > 0 ? 6 : 0;
  const total = textHeight + stackTopGap + stackGap + controlsHeight;
  return clamp(total, 0, 3600);
}

function estimateParagraphTextHeightPx(paragraph) {
  const style = paragraph?.style ?? null;
  const spacingBeforePx = Number.isFinite(style?.spacingBeforePx) ? clamp(style.spacingBeforePx, 0, 220) : 0;
  const spacingAfterPx = Number.isFinite(style?.spacingAfterPx) ? clamp(style.spacingAfterPx, 0, 220) : 0;

  if ((paragraph?.textLines?.length ?? 0) > 0) {
    const linesHeight = paragraph.textLines.reduce((sum, line) => {
      if (Number.isFinite(line?.estimatedHeightPx)) {
        return sum + clamp(line.estimatedHeightPx, 8, 280);
      }
      return sum + 18;
    }, 0);
    return clamp((linesHeight + spacingBeforePx + spacingAfterPx) * TEXT_PAGINATION_SCALE, 0, 2200);
  }

  const text = paragraph?.previewTextNormalized ?? paragraph?.previewText ?? "";
  if (!text) {
    return 0;
  }

  const lineHeightRatio = Number.parseFloat(style?.lineHeight ?? "1.6");
  const ratio = Number.isFinite(lineHeightRatio) ? clamp(lineHeightRatio, 1.0, 2.6) : 1.6;
  const fontPx = clamp((style?.fontSizePt ?? 10) * (96 / 72), 8, 96);
  const lineHeightPx = fontPx * ratio;
  const approxCharsPerLine = 38;
  const lineCount = Math.max(1, Math.ceil(text.length / approxCharsPerLine));
  return clamp((lineCount * lineHeightPx + spacingBeforePx + spacingAfterPx) * TEXT_PAGINATION_SCALE, 0, 2200);
}

function estimateControlHeightPx(control) {
  if (!control) {
    return 0;
  }
  if (control.tableInfo) {
    return estimateTableHeightPx(control.tableInfo);
  }
  if (control.graphicInfo) {
    if (isAbsoluteOverlayGraphic(control)) {
      return 0;
    }
    const { height } = resolveGraphicObjectDimensions(control.graphicInfo);
    const bodyHeight = clamp(hwpToPx(Math.max(0, height)), 12, 2400);
    return clamp(bodyHeight + 8, 12, 2800);
  }
  if (control.ctrlId === "fn  " || control.ctrlId === "en  ") {
    return 12;
  }
  if (control.ctrlId === "secd" || control.ctrlId === "cold" || control.ctrlId === "head" || control.ctrlId === "foot") {
    return 0;
  }
  return 24;
}

function isAbsoluteOverlayGraphic(control) {
  if (control?.ctrlId !== "gso " || !control.graphicInfo?.objectCommon?.propertyBits) {
    return false;
  }
  const bits = control.graphicInfo.objectCommon.propertyBits;
  const flow = bits.textFlow;
  if (!(flow === 2 || flow === 3) || bits.likeCharacter) {
    return false;
  }
  const vRel = bits.vertRelToName;
  const hRel = bits.horzRelToName;
  return (vRel === "paper" || vRel === "page") && (hRel === "paper" || hRel === "page");
}

function estimateTableHeightPx(tableInfo) {
  const rowHeights = estimateTableRowHeightsPx(tableInfo);
  if (!rowHeights.length) {
    return 0;
  }
  return clamp(TABLE_CHUNK_BASE_PX + rowHeights.reduce((sum, value) => sum + value, 0), 16, 6400);
}

function estimateTableRowHeightsPx(tableInfo) {
  const rowCount = Number.isFinite(tableInfo?.rows) ? Math.max(0, tableInfo.rows) : 0;
  if (!rowCount) {
    return [];
  }
  const rowHeights = Array.from({ length: rowCount }, () => 0);
  const rowSizes = Array.isArray(tableInfo?.rowSizes) ? tableInfo.rowSizes : [];
  const sizedRows = rowSizes.reduce((sum, rowSize) => sum + (Number.isFinite(rowSize) && rowSize > 0 ? 1 : 0), 0);
  const rowSizeCoverage = rowCount > 0 ? sizedRows / rowCount : 0;
  const rowSizeScale = rowSizeCoverage >= 0.7 ? 0.88 : rowSizeCoverage >= 0.45 ? 0.93 : 1.0;
  const inferredLinePx = rowSizeCoverage <= 0.25 ? 15.5 : 14.0;
  const fallbackRowPx = rowSizeCoverage <= 0.25 ? 25 : TABLE_ROW_FALLBACK_PX;
  for (let i = 0; i < rowCount; i += 1) {
    const rowSize = rowSizes[i];
    if (Number.isFinite(rowSize) && rowSize > 0) {
      rowHeights[i] = Math.max(rowHeights[i], clamp(hwpToPx(rowSize) * rowSizeScale, TABLE_ROW_MIN_PX, 360));
    }
  }

  const cells = Array.isArray(tableInfo?.cells) ? tableInfo.cells : [];
  for (const cell of cells) {
    if (!cell || cell.covered) {
      continue;
    }
    if (!Number.isFinite(cell.row) || cell.row < 0 || cell.row >= rowCount) {
      continue;
    }
    const rowSpan = Math.max(1, cell.rowSpan || 1);
    let cellHeightPx = Number.isFinite(cell.height) && cell.height > 0 ? hwpToPx(cell.height) : 0;
    if (!(cellHeightPx > 0) && cell.text) {
      const lines = Math.max(1, String(cell.text).split(/\n+/).length);
      cellHeightPx = lines * inferredLinePx + 8;
    }
    if (!(cellHeightPx > 0)) {
      continue;
    }
    const perRow = clamp((cellHeightPx / rowSpan) * rowSizeScale, TABLE_ROW_MIN_PX, 320);
    const rowEnd = Math.min(rowCount, cell.row + rowSpan);
    for (let row = cell.row; row < rowEnd; row += 1) {
      rowHeights[row] = Math.max(rowHeights[row], perRow);
    }
  }

  for (let i = 0; i < rowHeights.length; i += 1) {
    if (!(rowHeights[i] > 0)) {
      rowHeights[i] = fallbackRowPx;
    }
  }

  return rowHeights;
}

function pickTableRowsForHeight(rowHeights, rowStart, availableHeightPx) {
  if (!Array.isArray(rowHeights) || rowStart < 0 || rowStart >= rowHeights.length) {
    return 0;
  }
  if (!(availableHeightPx > TABLE_CHUNK_BASE_PX + TABLE_ROW_MIN_PX)) {
    return 0;
  }
  let used = TABLE_CHUNK_BASE_PX;
  let rows = 0;
  for (let i = rowStart; i < rowHeights.length; i += 1) {
    const rowHeight = clamp(rowHeights[i], TABLE_ROW_MIN_PX, 320);
    if (rows > 0 && used + rowHeight > availableHeightPx) {
      break;
    }
    if (rows === 0 && used + rowHeight > availableHeightPx) {
      return 1;
    }
    used += rowHeight;
    rows += 1;
  }
  return rows;
}

function sliceTableInfoByRowRange(tableInfo, rowStart, rowEnd) {
  const start = clamp(Math.floor(rowStart), 0, Math.max(0, tableInfo.rows - 1));
  const end = clamp(Math.floor(rowEnd), start + 1, tableInfo.rows);
  const rows = Math.max(1, end - start);
  const cells = (tableInfo.cells ?? [])
    .map((cell) => {
      if (!cell || cell.covered) {
        return null;
      }
      const cellStart = cell.row;
      const cellEnd = cell.row + Math.max(1, cell.rowSpan);
      if (cellEnd <= start || cellStart >= end) {
        return null;
      }
      const clippedStart = Math.max(cellStart, start);
      const clippedEnd = Math.min(cellEnd, end);
      const clippedSpan = Math.max(1, clippedEnd - clippedStart);
      return {
        ...cell,
        row: clippedStart - start,
        rowSpan: clippedSpan,
      };
    })
    .filter(Boolean);

  return {
    ...tableInfo,
    rows,
    rowSizes: Array.isArray(tableInfo.rowSizes) ? tableInfo.rowSizes.slice(start, end) : [],
    cells,
    zones: [],
  };
}

function trimLeadingEmptyPages(pages) {
  if (!pages?.length) {
    return [];
  }
  let start = 0;
  while (start < pages.length && !hasRenderablePageContent(pages[start])) {
    start += 1;
  }
  if (start <= 0) {
    return pages;
  }
  return pages.slice(start);
}

function hasRenderablePageContent(page) {
  if (!page?.paragraphs?.length) {
    return false;
  }
  return page.paragraphs.some((paragraph) => {
    const hasText = (paragraph?.previewTextNormalized ?? "").length > 0;
    const hasControls = (paragraph?.controls?.length ?? 0) > 0;
    return hasText || hasControls;
  });
}

function columnLayoutKey(layout) {
  if (!layout) {
    return "none";
  }
  const widths = layout.widths?.length ? layout.widths.join(",") : "";
  return `${layout.count}|${layout.gap}|${layout.direction}|${layout.sameWidth ? 1 : 0}|${widths}`;
}

function buildPreviewFlowStyle(columnLayout) {
  if (!columnLayout || columnLayout.count <= 1) {
    return "";
  }
  const styles = [
    `column-count:${Math.max(2, columnLayout.count)}`,
    `column-gap:${escapeCssLength(columnLayout.gapPx, 24)}px`,
    `column-fill:balance`,
  ];
  if (columnLayout.direction === 1) {
    styles.push("direction:rtl");
  }
  return styles.join(";");
}

function escapeCssLength(value, fallback = 0) {
  if (!Number.isFinite(value)) {
    return fallback;
  }
  return clamp(value, 0, 240).toFixed(1);
}

function defaultColumnLayout() {
  return {
    count: 1,
    kind: "normal",
    kindName: "normal",
    direction: 0,
    directionName: "left-to-right",
    sameWidth: true,
    gap: 0,
    gapPx: 24,
    widths: [],
    widthsPx: [],
    separatorType: 0,
    separatorWidth: 0,
    separatorColor: null,
  };
}

function buildPreviewModel() {
  const doc = state.doc;
  if (!doc) {
    return null;
  }
  if (doc.formatInfo?.mode === "legacy-3x-binary" || doc.formatInfo?.mode === "hwpml-xml") {
    return {
      sectionCount: 0,
      controlCount: 0,
      pageLayout: null,
      columnLayout: null,
      paragraphs: [],
    };
  }
  if (doc.previewModel) {
    return doc.previewModel;
  }

  const sectionStreams = getPreviewSectionStreams(doc.entries, doc.fileHeader);
  const paragraphs = [];
  const numberingState = new Map();
  let pageLayout = null;
  let columnLayout = defaultColumnLayout();

  for (const stream of sectionStreams) {
    const analysis = getStreamAnalysis(stream);
    let previousTopLevelLineY = null;
    let currentPageLayout = pageLayout;
    let currentColumnLayout = columnLayout;
    for (const paragraph of analysis.paragraphs) {
      const header = paragraph.headerRecord
        ? parseParaHeader(getRecordPayload(analysis.decoded, paragraph.headerRecord))
        : null;
      const paragraphTextData = paragraph.textRecord
        ? decodeParagraphText(getRecordPayload(analysis.decoded, paragraph.textRecord))
        : { text: "", extensions: [], tokens: [], logicalLength: 0 };
      const charRuns = paragraph.charShapeRecord
        ? parseParaCharShape(getRecordPayload(analysis.decoded, paragraph.charShapeRecord))
        : [];
      const ctrlHeaders = paragraph.ctrlHeaderRecords?.length
        ? paragraph.ctrlHeaderRecords.map((record) => parseCtrlHeader(getRecordPayload(analysis.decoded, record), record))
        : [];
      const allControls = buildPreviewControlBlocks(analysis, ctrlHeaders, paragraphTextData.extensions, doc);
      for (const control of allControls) {
        if (control.sectionInfo?.pageLayout) {
          currentPageLayout = control.sectionInfo.pageLayout;
          if (!pageLayout) {
            pageLayout = currentPageLayout;
          }
        }
        if (control.columnInfo?.layout) {
          currentColumnLayout = control.columnInfo.layout;
          columnLayout = currentColumnLayout;
        }
      }
      const controls = dedupePreviewControls(allControls.filter((control) => isRenderableControl(control)));

      const docInfo = doc.docInfo;
      const paraShape = header ? docInfo?.paraShapeById.get(header.paraShapeId) : null;
      const styleRef = header ? docInfo?.styleById.get(header.paraStyleId) : null;
      let charShapeId = charRuns.length ? charRuns[0].charShapeId : null;
      if (charShapeId == null && styleRef) {
        charShapeId = styleRef.charShapeId ?? null;
      }
      const charShape = charShapeId != null ? docInfo?.charShapeById.get(charShapeId) : null;

      const visibleTokens = (paragraphTextData.tokens ?? []).filter((token) => token.type !== "control");
      const visibleText = tokensToPlainText(visibleTokens);
      const normalizedText = normalizePreviewText(visibleText);
      const lineSegments = paragraph.lineSegRecord
        ? parseParaLineSeg(getRecordPayload(analysis.decoded, paragraph.lineSegRecord))
        : [];
      const breakHints = extractParagraphBreakHints(lineSegments, header);
      const firstLineY = lineSegments.length ? lineSegments[0].lineVerticalPos : null;
      const headerLevel = paragraph.headerRecord?.level ?? 0;
      const canUseYPageHeuristic = headerLevel <= 1;
      if (
        !breakHints.pageStart &&
        canUseYPageHeuristic &&
        Number.isFinite(firstLineY) &&
        Number.isFinite(previousTopLevelLineY) &&
        firstLineY + PAGE_Y_RESET_THRESHOLD < previousTopLevelLineY
      ) {
        breakHints.pageStart = true;
      }
      if (canUseYPageHeuristic && Number.isFinite(firstLineY)) {
        const maxLineY = lineSegments.reduce((max, segment) => {
          if (!Number.isFinite(segment.lineVerticalPos)) {
            return max;
          }
          return Math.max(max, segment.lineVerticalPos);
        }, firstLineY);
        previousTopLevelLineY = maxLineY;
      }

      const structuralBreakOnly = breakHints.pageStart || breakHints.columnStart;
      if (!visibleText.length && controls.length === 0 && !structuralBreakOnly) {
        continue;
      }

      const previewStyle = resolvePreviewStyle(paraShape, charShape);
      const listContext = resolveParagraphListContext(paraShape, docInfo, numberingState);
      const textRuns = buildStyledTextRuns(visibleTokens, charRuns, docInfo, previewStyle, charShapeId);
      const textLines =
        lineSegments.length > 0
          ? buildStyledTextLines(visibleTokens, charRuns, lineSegments, docInfo, previewStyle, charShapeId)
          : [];
      paragraphs.push({
        sectionLabel: toSectionLabel(stream.fullPath),
        previewText: visibleText,
        previewTextNormalized: normalizedText,
        textRuns,
        textLines,
        listMarker: listContext.marker,
        listType: listContext.type,
        listLevel: listContext.level,
        tabStops: listContext.tabStops,
        tabStopsPx: listContext.tabStopsPx,
        breakHints,
        controls,
        pageLayout: currentPageLayout,
        columnLayout: currentColumnLayout,
        refs: {
          paraShapeId: header?.paraShapeId ?? null,
          paraStyleId: header?.paraStyleId ?? null,
          charShapeId,
        },
        style: previewStyle,
      });
    }
  }

  const controlCount = paragraphs.reduce((sum, paragraph) => sum + paragraph.controls.length, 0);
  const model = {
    sectionCount: sectionStreams.length,
    controlCount,
    pageLayout,
    columnLayout,
    paragraphs,
  };
  doc.previewModel = model;
  return model;
}

function extractParagraphBreakHints(lineSegments, paraHeader = null) {
  if (!lineSegments?.length) {
    const splitType = paraHeader?.splitType ?? 0;
    return {
      pageStart: Boolean((splitType & PARA_SPLIT_PAGE_BIT) || (splitType & PARA_SPLIT_SECTION_BIT)),
      columnStart: Boolean((splitType & PARA_SPLIT_COLUMN_BIT) || (splitType & PARA_SPLIT_COLUMNS_DEF_BIT)),
    };
  }

  const sorted = [...lineSegments]
    .filter((segment) => Number.isFinite(segment.textStartPos))
    .sort((a, b) => a.textStartPos - b.textStartPos);
  if (!sorted.length) {
    return {
      pageStart: false,
      columnStart: false,
    };
  }

  const hasPageFirst = sorted.some((segment) => Boolean((segment.flags ?? 0) & LINE_SEG_PAGE_FIRST_BIT));
  const hasColumnFirst = sorted.some((segment) => Boolean((segment.flags ?? 0) & LINE_SEG_COLUMN_FIRST_BIT));
  const splitType = paraHeader?.splitType ?? 0;
  const headerPageBreak = Boolean((splitType & PARA_SPLIT_PAGE_BIT) || (splitType & PARA_SPLIT_SECTION_BIT));
  const headerColumnBreak = Boolean((splitType & PARA_SPLIT_COLUMN_BIT) || (splitType & PARA_SPLIT_COLUMNS_DEF_BIT));
  return {
    pageStart: hasPageFirst || headerPageBreak,
    columnStart: hasColumnFirst || headerColumnBreak,
  };
}

function resolveParagraphListContext(paraShape, docInfo, numberingState) {
  const tabDef = paraShape?.tabDefId != null ? docInfo?.tabDefById?.get(paraShape.tabDefId) ?? null : null;
  const tabStops = tabDef?.tabStops?.length
    ? tabDef.tabStops
        .filter((tab) => Number.isFinite(tab?.positionPx) && tab.positionPx > 0)
        .map((tab) => ({
          positionPx: tab.positionPx,
          align: tab.align ?? 0,
          alignName: tab.alignName ?? TAB_ALIGN_NAMES[tab.align] ?? `align${tab.align ?? 0}`,
          leader: tab.leader ?? 0,
          leaderName: tab.leaderName ?? TAB_LEADER_NAMES[tab.leader] ?? `leader${tab.leader ?? 0}`,
        }))
        .sort((a, b) => a.positionPx - b.positionPx)
    : [];
  const tabStopsPx = tabStops.map((tab) => tab.positionPx);

  const headingType = paraShape?.property1Bits?.headingType ?? 0;
  const listLevel = clamp(paraShape?.property1Bits?.headingLevel || 1, 1, 7);
  const listId = paraShape?.numberingId;
  if (headingType === 2) {
    const marker = buildNumberingMarker(listId, listLevel, docInfo, numberingState);
    return {
      type: "number",
      marker,
      level: listLevel,
      tabStops,
      tabStopsPx,
    };
  }
  if (headingType === 3) {
    const marker = buildBulletMarker(listId, docInfo);
    return {
      type: "bullet",
      marker,
      level: listLevel,
      tabStops,
      tabStopsPx,
    };
  }

  return {
    type: "",
    marker: null,
    level: null,
    tabStops,
    tabStopsPx,
  };
}

function buildNumberingMarker(listId, listLevel, docInfo, numberingState) {
  if (listId == null) {
    return null;
  }
  const numbering = docInfo?.numberingById?.get(listId) ?? null;
  if (!numbering) {
    return null;
  }

  const counters = numberingState.get(listId) ?? Array.from({ length: 7 }, () => 0);
  numberingState.set(listId, counters);

  const levelIndex = clamp((listLevel ?? 1) - 1, 0, 6);
  counters[levelIndex] = (counters[levelIndex] || 0) + 1;
  for (let i = levelIndex + 1; i < counters.length; i += 1) {
    counters[i] = 0;
  }
  for (let i = 0; i < levelIndex; i += 1) {
    if (!counters[i]) {
      counters[i] = 1;
    }
  }

  const format = numbering.levels[levelIndex]?.format?.trim() || `^${levelIndex + 1}.`;
  const marker = format
    .replace(/\^([1-7])/g, (_, level) => String(counters[Number(level) - 1] || 1))
    .replace(/\u0000/g, "")
    .trim();
  return marker || `${counters[levelIndex]}.`;
}

function buildBulletMarker(listId, docInfo) {
  if (listId == null) {
    return null;
  }
  const bullet = docInfo?.bulletById?.get(listId) ?? null;
  if (!bullet?.marker) {
    return "•";
  }
  return bullet.marker;
}

function buildPreviewControlBlocks(analysis, ctrlHeaders, textExtensions, doc = state.doc) {
  if (!ctrlHeaders.length) {
    return [];
  }

  const extensionCounts = countControlExtensions(textExtensions);
  return ctrlHeaders.map((ctrl) => {
    const descendants = collectControlDescendantRecords(analysis.records, ctrl.recordIndex);
    const listHeaderRecords = descendants.filter((record) => record.tag === 72);
    const listHeader = listHeaderRecords[0] ?? null;
    const table = descendants.find((record) => record.tag === 77);
    const shape = descendants.find((record) => record.tag === 76);
    const ctrlRecord = ctrl.recordIndex != null ? analysis.records[ctrl.recordIndex] ?? null : null;
    const ctrlPayload = ctrlRecord ? getRecordPayload(analysis.decoded, ctrlRecord) : null;
    const detailLines = [`headerWords=${ctrl.wordsText || "-"}`];
    if (listHeader) {
      const listInfo = parseListHeader(getRecordPayload(analysis.decoded, listHeader));
      if (listInfo) {
        detailLines.push(
          `list paragraphs=${listInfo.paragraphCount}, direction=${listInfo.textDirectionName}, break=${listInfo.lineBreakName}, valign=${listInfo.verticalAlignName}`
        );
      } else {
        detailLines.push(`listHeaderWords=${payloadWordPreview(getRecordPayload(analysis.decoded, listHeader))}`);
      }
    }
    if (table) {
      detailLines.push(`tableWords=${payloadWordPreview(getRecordPayload(analysis.decoded, table))}`);
    }
    if (shape) {
      detailLines.push(`shapeWords=${payloadWordPreview(getRecordPayload(analysis.decoded, shape))}`);
    }
    const sectionInfo = ctrl.ctrlId === "secd" ? parseSectionControl(descendants, analysis.decoded) : null;
    if (sectionInfo?.pageDef) {
      detailLines.push(
        `page ${sectionInfo.pageDef.orientation}, paper=${sectionInfo.pageDef.paperWidth}x${sectionInfo.pageDef.paperHeight}, text=${sectionInfo.pageDef.textWidth}x${sectionInfo.pageDef.textHeight}`
      );
      detailLines.push(
        `page margin LRTB=${sectionInfo.pageDef.leftMargin}/${sectionInfo.pageDef.rightMargin}/${sectionInfo.pageDef.topMargin}/${sectionInfo.pageDef.bottomMargin}`
      );
    }

    const graphicInfo = ctrl.ctrlId === "gso " ? parseGraphicControl(descendants, analysis.decoded, doc, ctrlPayload) : null;
    if (graphicInfo) {
      const dims = graphicInfo.objectCommon
        ? `${graphicInfo.objectCommon.width}x${graphicInfo.objectCommon.height}`
        : graphicInfo.shapeComponent
          ? `${graphicInfo.shapeComponent.currentWidth}x${graphicInfo.shapeComponent.currentHeight}`
          : "-";
      detailLines.push(`graphic type=${graphicInfo.objectTypeName}, size=${dims}, textParas=${graphicInfo.textParagraphs.length}`);
      if (graphicInfo.objectCommon?.propertyBits) {
        detailLines.push(
          `graphic flow=${graphicInfo.objectCommon.propertyBits.textFlowName}, side=${graphicInfo.objectCommon.propertyBits.textSideName}, inline=${
            graphicInfo.objectCommon.propertyBits.likeCharacter ? "yes" : "no"
          }`
        );
      }
      if (graphicInfo.shapeTagNames.length) {
        detailLines.push(`graphic tags=${graphicInfo.shapeTagNames.join("|")}`);
      }
      if (graphicInfo.pictureInfo?.binDataId) {
        detailLines.push(`graphic picture binDataId=${graphicInfo.pictureInfo.binDataId}`);
      }
      if (graphicInfo.videoInfo?.binDataId) {
        detailLines.push(`graphic video binDataId=${graphicInfo.videoInfo.binDataId}`);
      }
      if (graphicInfo.chartInfo?.summary) {
        detailLines.push(`graphic chart=${graphicInfo.chartInfo.summary}`);
      }
      if (graphicInfo.textArtInfo?.text) {
        detailLines.push(`graphic textArt=${graphicInfo.textArtInfo.text.slice(0, 120)}`);
      }
      if (graphicInfo.formObjectInfo?.name) {
        detailLines.push(`graphic formObject=${graphicInfo.formObjectInfo.name}`);
      }
      if (graphicInfo.imageRef) {
        detailLines.push(
          `graphic image=${graphicInfo.imageRef.format || graphicInfo.imageRef.mime}, decode=${graphicInfo.imageRef.decodeStrategy}, bytes=${graphicInfo.imageRef.rawSize}/${graphicInfo.imageRef.decodedSize}`
        );
      }
      if (graphicInfo.mediaRef && graphicInfo.mediaRef !== graphicInfo.imageRef) {
        detailLines.push(
          `graphic media=${graphicInfo.mediaRef.format || graphicInfo.mediaRef.mime}, decode=${graphicInfo.mediaRef.decodeStrategy}, bytes=${graphicInfo.mediaRef.rawSize}/${graphicInfo.mediaRef.decodedSize}`
        );
      }
    }

    const columnInfo = ctrl.ctrlId === "cold" ? parseColumnControl(ctrlPayload) : null;
    if (columnInfo?.layout) {
      detailLines.push(
        `columns count=${columnInfo.layout.count}, type=${columnInfo.layout.kindName}, direction=${columnInfo.layout.directionName}, sameWidth=${
          columnInfo.layout.sameWidth ? "yes" : "no"
        }`
      );
      detailLines.push(
        `columns gap=${columnInfo.layout.gap} (${columnInfo.layout.gapPx.toFixed(1)}px), widths=${
          columnInfo.layout.widths.length ? columnInfo.layout.widths.join("/") : "equal"
        }`
      );
      if (columnInfo.layout.widths.length) {
        detailLines.push(`columns widths=${columnInfo.layout.widths.join(",")}`);
      }
      if (columnInfo.layout.separatorType || columnInfo.layout.separatorWidth) {
        detailLines.push(
          `columns separator type=${columnInfo.layout.separatorType}, width=${columnInfo.layout.separatorWidth}, color=${
            columnInfo.layout.separatorColor ? columnInfo.layout.separatorColor.hex : "-"
          }`
        );
      }
    }

    let tableInfo = null;
    if (ctrl.ctrlId === "tbl " || table) {
      tableInfo = table ? parseTableControl(getRecordPayload(analysis.decoded, table)) : null;
      const listHeaders = listHeaderRecords
        .map((record) => parseListHeader(getRecordPayload(analysis.decoded, record)))
        .filter(Boolean);
      const cellTableInfo = buildTableInfoFromListHeaders(listHeaders, tableInfo);
      if (cellTableInfo) {
        tableInfo = cellTableInfo;
      }
      if (tableInfo) {
        tableInfo = attachTableCellTexts(tableInfo, descendants, analysis.decoded);
      }
      if (tableInfo) {
        detailLines.push(
          `table rows=${tableInfo.rows}, cols=${tableInfo.cols}, cells=${tableInfo.cells.length}, spacing=${tableInfo.cellSpacing}`
        );
        detailLines.push(`table borderFill=${tableInfo.borderFillId}, zones=${tableInfo.zones.length}`);
        detailLines.push(`table paraPool=${tableInfo.paragraphPool}, paraMapped=${tableInfo.mappedParagraphs}`);
      } else if (table || ctrl.ctrlId === "tbl ") {
        detailLines.push(`table parse failed`);
      }
    }

    return {
      ...ctrl,
      kind: classifyControlKind(ctrl.ctrlId),
      subRecordCount: descendants.length,
      tagSummary: summarizeRecordTags(descendants),
      detailLines,
      sectionInfo,
      graphicInfo,
      columnInfo,
      tableInfo,
      inlineCount: extensionCounts.get(ctrl.ctrlId || "") ?? 0,
    };
  });
}

function dedupePreviewControls(controls) {
  if (!controls?.length) {
    return [];
  }
  const seen = new Set();
  const kept = [];

  for (let i = controls.length - 1; i >= 0; i -= 1) {
    const control = controls[i];
    const signature = getControlDedupSignature(control);
    if (signature && seen.has(signature)) {
      continue;
    }
    if (signature) {
      seen.add(signature);
    }
    kept.push(control);
  }

  kept.reverse();
  return kept;
}

function getControlDedupSignature(control) {
  if (!control) {
    return "";
  }
  if (control.ctrlId !== "gso " || !control.graphicInfo?.imageRef?.id) {
    return "";
  }

  const objectWidth =
    Number.isFinite(control.graphicInfo.objectCommon?.width) && control.graphicInfo.objectCommon.width > 0
      ? control.graphicInfo.objectCommon.width
      : 0;
  const objectHeight =
    Number.isFinite(control.graphicInfo.objectCommon?.height) && control.graphicInfo.objectCommon.height > 0
      ? control.graphicInfo.objectCommon.height
      : 0;
  const componentWidth =
    Number.isFinite(control.graphicInfo.shapeComponent?.currentWidth) && control.graphicInfo.shapeComponent.currentWidth > 0
      ? control.graphicInfo.shapeComponent.currentWidth
      : 0;
  const componentHeight =
    Number.isFinite(control.graphicInfo.shapeComponent?.currentHeight) && control.graphicInfo.shapeComponent.currentHeight > 0
      ? control.graphicInfo.shapeComponent.currentHeight
      : 0;
  const width = Math.max(objectWidth, componentWidth, 0);
  const height = Math.max(objectHeight, componentHeight, 0);

  const widthBucket = Math.round(width / 200);
  const heightBucket = Math.round(height / 200);
  return `gso-image:${control.graphicInfo.imageRef.id}:${widthBucket}:${heightBucket}`;
}

function countControlExtensions(extensions) {
  const counts = new Map();
  for (const extension of extensions ?? []) {
    const key = extension.ctrlId || "";
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  return counts;
}

function collectControlDescendantRecords(records, ctrlRecordIndex) {
  if (!records?.length || ctrlRecordIndex == null) {
    return [];
  }
  const start = records.findIndex((record) => record.index === ctrlRecordIndex);
  if (start < 0) {
    return [];
  }

  const baseLevel = records[start].level;
  const descendants = [];
  for (let i = start + 1; i < records.length; i += 1) {
    const record = records[i];
    if (record.level <= baseLevel) {
      break;
    }
    descendants.push(record);
  }
  return descendants;
}

function summarizeRecordTags(records) {
  if (!records.length) {
    return [];
  }

  const counts = new Map();
  for (const record of records) {
    const name = getRecordTagName(record.tag);
    counts.set(name, (counts.get(name) ?? 0) + 1);
  }

  return Array.from(counts.entries()).map(([name, count]) => `${name}×${count}`);
}

function payloadWordPreview(payload, maxWords = 4) {
  if (!payload?.length) {
    return "-";
  }

  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const words = [];
  for (let offset = 0; offset + 3 < payload.length && words.length < maxWords; offset += 4) {
    words.push(`0x${dv.getUint32(offset, true).toString(16)}`);
  }
  if (words.length) {
    return words.join(", ");
  }

  return Array.from(payload.subarray(0, Math.min(payload.length, 8)))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join(" ");
}

function classifyControlKind(ctrlId) {
  if (!ctrlId) {
    return "inline";
  }
  return CONTROL_KIND_BY_ID[ctrlId] ?? "inline";
}

function getRecordTagName(tag) {
  const name = RECORD_TAGS[tag] ?? `TAG_${tag}`;
  const aliases = Object.entries(RECORD_TAG_ALIASES)
    .filter(([, value]) => value === tag)
    .map(([alias]) => alias);
  if (!aliases.length) {
    return name;
  }
  return `${name} (${aliases.join(",")})`;
}

function parseListHeader(payload) {
  if (!payload || payload.length < 8) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const paragraphCount = dv.getUint16(0, true);
  const unknown1 = dv.getUint16(2, true);
  const property = dv.getUint32(4, true);
  const textDirection = property & 0x7;
  const lineBreak = (property >>> 3) & 0x3;
  const verticalAlign = (property >>> 5) & 0x3;

  const result = {
    paragraphCount,
    unknown1,
    property,
    textDirection,
    lineBreak,
    verticalAlign,
    textDirectionName: LIST_TEXT_DIRECTION_NAMES[textDirection] ?? `dir${textDirection}`,
    lineBreakName: LIST_LINE_BREAK_NAMES[lineBreak] ?? `break${lineBreak}`,
    verticalAlignName: LIST_VERTICAL_ALIGN_NAMES[verticalAlign] ?? `v${verticalAlign}`,
    payloadSize: payload.length,
    cell: null,
  };

  if (payload.length >= 34) {
    const col = dv.getUint16(8, true);
    const row = dv.getUint16(10, true);
    const colSpanRaw = dv.getUint16(12, true);
    const rowSpanRaw = dv.getUint16(14, true);
    const width = dv.getInt32(16, true);
    const height = dv.getInt32(20, true);
    const margin = [
      dv.getInt16(24, true),
      dv.getInt16(26, true),
      dv.getInt16(28, true),
      dv.getInt16(30, true),
    ];
    const borderFillId = dv.getUint16(32, true);
    const colSpan = colSpanRaw > 0 ? colSpanRaw : 1;
    const rowSpan = rowSpanRaw > 0 ? rowSpanRaw : 1;
    if (
      col < 8192 &&
      row < 8192 &&
      colSpan <= 2048 &&
      rowSpan <= 2048 &&
      Number.isFinite(width) &&
      Number.isFinite(height) &&
      Math.abs(width) < 100000000 &&
      Math.abs(height) < 100000000
    ) {
      result.cell = {
        paragraphCount: Math.max(0, paragraphCount),
        listProperty: property,
        col,
        row,
        colSpan,
        rowSpan,
        width,
        height,
        margin,
        borderFillId,
      };
    }
  };
  return result;
}

function parseColumnControl(payload) {
  if (!payload || payload.length < 14) {
    return null;
  }

  // CTRL_HEADER payload layout: [4-byte CtrlID][control-specific bytes...]
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const bodyOffset = 4;
  if (bodyOffset + 10 > payload.length) {
    return null;
  }

  const propertyLow = dv.getUint16(bodyOffset, true);
  const gap = dv.getInt16(bodyOffset + 2, true);
  const countByBits = clamp((propertyLow >>> 2) & 0xff, 1, 32);
  const sameWidth = Boolean((propertyLow >>> 12) & 0x1);

  let cursor = bodyOffset + 4;
  const widths = [];
  if (!sameWidth && countByBits > 0) {
    const reservedTail = 8;
    const maxWidthCount = Math.max(0, Math.floor((payload.length - cursor - reservedTail) / 2));
    const widthCount = Math.min(countByBits, maxWidthCount);
    for (let i = 0; i < widthCount; i += 1) {
      widths.push(dv.getInt16(cursor, true));
      cursor += 2;
    }
  }

  if (cursor + 8 > payload.length) {
    return null;
  }

  const propertyHigh = dv.getUint16(cursor, true);
  const property = (propertyLow | (propertyHigh << 16)) >>> 0;
  const separatorType = dv.getUint8(cursor + 2);
  const separatorWidth = dv.getUint8(cursor + 3);
  const separatorColor = parseColorRef(dv.getUint32(cursor + 4, true));

  const kind = property & 0x3;
  const count = clamp((property >>> 2) & 0xff, 1, 32);
  const direction = (property >>> 10) & 0x3;
  const equalWidthByProperty = Boolean((property >>> 12) & 0x1);
  const effectiveSameWidth = sameWidth || equalWidthByProperty;

  const widthsPx = widths.map((value) => clamp(hwpToPx(value), 24, 780));
  const layout = {
    count,
    kind,
    kindName: COLUMN_TYPE_NAMES[kind] ?? `type${kind}`,
    direction,
    directionName: COLUMN_DIRECTION_NAMES[direction] ?? `dir${direction}`,
    sameWidth: effectiveSameWidth,
    gap,
    gapPx: clamp(hwpToPx(gap), 0, 140),
    widths,
    widthsPx,
    separatorType,
    separatorWidth,
    separatorColor,
    property,
  };

  return {
    propertyLow,
    propertyHigh,
    layout,
  };
}

function renderColumnControlPreview(columnInfo) {
  const layout = columnInfo?.layout;
  if (!layout) {
    return "";
  }
  const widthInfo =
    !layout.sameWidth && layout.widths.length
      ? `<div class="control-column-widths">${escapeHtml(layout.widths.map((value) => String(value)).join(" / "))}</div>`
      : "";
  const separatorInfo =
    layout.separatorType || layout.separatorWidth
      ? `<span>${escapeHtml(
          `separator type=${layout.separatorType} width=${layout.separatorWidth} color=${layout.separatorColor?.hex ?? "-"}`
        )}</span>`
      : "";

  return `
    <div class="control-column-wrap">
      <div class="control-column-row">
        <span>${escapeHtml(`columns ${layout.count}`)}</span>
        <span>${escapeHtml(layout.kindName)}</span>
        <span>${escapeHtml(layout.directionName)}</span>
      </div>
      <div class="control-column-row">
        <span>${escapeHtml(`gap ${layout.gap} (${layout.gapPx.toFixed(1)}px)`)}</span>
        <span>${escapeHtml(`sameWidth ${layout.sameWidth ? "on" : "off"}`)}</span>
        <span>${escapeHtml(`kind ${layout.kindName}`)}</span>
      </div>
      ${separatorInfo ? `<div class="control-column-row">${separatorInfo}</div>` : ""}
      ${widthInfo}
    </div>
  `;
}

function parseSectionControl(descendants, decoded) {
  const pageDefRecord = descendants.find((record) => record.tag === 73) ?? null;
  if (!pageDefRecord) {
    return null;
  }
  const pageDef = parsePageDef(getRecordPayload(decoded, pageDefRecord));
  if (!pageDef) {
    return { pageDef: null, pageLayout: null };
  }
  return {
    pageDef,
    pageLayout: toPageLayout(pageDef),
  };
}

function parsePageDef(payload) {
  if (!payload || payload.length < 40) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const paperWidth = dv.getInt32(0, true);
  const paperHeight = dv.getInt32(4, true);
  const leftMargin = dv.getInt32(8, true);
  const rightMargin = dv.getInt32(12, true);
  const topMargin = dv.getInt32(16, true);
  const bottomMargin = dv.getInt32(20, true);
  const headerMargin = dv.getInt32(24, true);
  const footerMargin = dv.getInt32(28, true);
  const gutterMargin = dv.getInt32(32, true);
  const property = dv.getUint32(36, true);
  const orientationBit = property & 0x1;
  const bindingMode = (property >>> 1) & 0x3;
  const orientation = orientationBit ? "landscape" : "portrait";
  const textWidth = Math.max(1, paperWidth - leftMargin - rightMargin);
  const textHeight = Math.max(1, paperHeight - topMargin - bottomMargin);

  return {
    paperWidth,
    paperHeight,
    leftMargin,
    rightMargin,
    topMargin,
    bottomMargin,
    headerMargin,
    footerMargin,
    gutterMargin,
    property,
    orientation,
    bindingMode,
    bindingName: PAGE_BINDING_NAMES[bindingMode] ?? `binding${bindingMode}`,
    textWidth,
    textHeight,
  };
}

function toPageLayout(pageDef) {
  if (!pageDef) {
    return null;
  }
  const paperWidthPx = hwpToPx(pageDef.paperWidth);
  const paperHeightPx = hwpToPx(pageDef.paperHeight);
  const textWidthPx = clamp(hwpToPx(pageDef.textWidth), 360, 1280);
  const leftPaddingPx = clamp(hwpToPx(pageDef.leftMargin), 0, 220);
  const rightPaddingPx = clamp(hwpToPx(pageDef.rightMargin), 0, 220);
  const topPaddingPx = clamp(hwpToPx(pageDef.topMargin), 0, 260);
  const bottomPaddingPx = clamp(hwpToPx(pageDef.bottomMargin), 0, 260);

  return {
    orientation: pageDef.orientation,
    paperWidthPx,
    paperHeightPx,
    textWidthPx,
    leftPaddingPx,
    rightPaddingPx,
    topPaddingPx,
    bottomPaddingPx,
    paperWidthMm: hwpToMm(pageDef.paperWidth),
    paperHeightMm: hwpToMm(pageDef.paperHeight),
  };
}

function renderSectionControlPreview(sectionInfo) {
  if (!sectionInfo?.pageDef) {
    return "";
  }
  const def = sectionInfo.pageDef;
  return `
    <div class="control-page-wrap">
      <div class="control-page-row">
        <span>${escapeHtml(`${def.orientation} / ${def.bindingName}`)}</span>
        <span>${escapeHtml(`paper ${Math.round(hwpToMm(def.paperWidth))}x${Math.round(hwpToMm(def.paperHeight))} mm`)}</span>
      </div>
      <div class="control-page-row">
        <span>${escapeHtml(`text ${Math.round(hwpToMm(def.textWidth))}x${Math.round(hwpToMm(def.textHeight))} mm`)}</span>
        <span>${escapeHtml(`margins ${def.leftMargin}/${def.rightMargin}/${def.topMargin}/${def.bottomMargin}`)}</span>
      </div>
    </div>
  `;
}

function renderGraphicControlPreview(graphicInfo) {
  const { width, height } = resolveGraphicObjectDimensions(graphicInfo);
  const xOffset = graphicInfo.objectCommon?.xOffset ?? graphicInfo.shapeComponent?.groupOffsetX ?? 0;
  const yOffset = graphicInfo.objectCommon?.yOffset ?? graphicInfo.shapeComponent?.groupOffsetY ?? 0;
  const widthPx = clamp(hwpToPx(Math.max(0, width)), 24, 1600);
  const heightPx = clamp(hwpToPx(Math.max(0, height)), 16, 2200);
  const rotation = graphicInfo.shapeComponent?.rotation ?? 0;
  const bits = graphicInfo.objectCommon?.propertyBits ?? null;
  const textFlow = bits?.textFlow ?? null;
  const textSide = bits?.textSide ?? null;
  const supportsWrapFlow = textFlow === 0 || textFlow === 4 || textFlow === 5;
  const isBlockFlow = textFlow === 1;
  const isBehindFlow = textFlow === 2;
  const isFrontFlow = textFlow === 3;
  let floatSide = "";
  if (supportsWrapFlow) {
    if (textSide === 1) {
      floatSide = "left";
    } else if (textSide === 2 || textSide === 3) {
      floatSide = "right";
    }
  }

  const wrapStyles = [];
  if (!PREVIEW_DEBUG) {
    if (floatSide) {
      wrapStyles.push(`max-width:${clamp(widthPx + 24, 80, 1600).toFixed(1)}px`);
      wrapStyles.push(`float:${floatSide}`);
      wrapStyles.push(floatSide === "left" ? "margin:0.2rem 0.9rem 0.6rem 0" : "margin:0.2rem 0 0.6rem 0.9rem");
    } else if (isBlockFlow) {
      wrapStyles.push("clear:both");
      wrapStyles.push("margin:0.45rem auto");
    } else if (isBehindFlow || isFrontFlow) {
      wrapStyles.push("clear:both");
      wrapStyles.push("margin:0.45rem auto");
    }
  }

  const textBlocks = graphicInfo.textBlocks ?? [];
  const sampleText = graphicInfo.textParagraphs
    .map((value) => normalizePreviewText(value))
    .filter((value) => value.length > 0)
    .join("\n");
  const mediaWidthPx = clamp(widthPx, 18, 1600);
  const mediaHeightPx = clamp(heightPx, 18, 2200);
  const flowLabel = bits?.textFlowName ? `${bits.textFlowName}/${bits.textSideName}` : "";
  const imageRef = graphicInfo.imageRef ?? null;
  const mediaRef = graphicInfo.mediaRef ?? null;
  const imageTag = imageRef
    ? `<img class="control-object-image" src="${escapeHtmlAttr(imageRef.url)}" alt="${escapeHtmlAttr(
        graphicInfo.objectTypeName || "image"
      )}" loading="lazy" decoding="async" style="width:${mediaWidthPx.toFixed(1)}px;max-width:100%;height:auto;max-height:${mediaHeightPx.toFixed(
        1
      )}px">`
    : null;
  const videoTag =
    !imageTag && mediaRef?.mime?.startsWith("video/")
      ? `<video class="control-object-video" src="${escapeHtmlAttr(
          mediaRef.url
        )}" controls preload="metadata" style="width:${mediaWidthPx.toFixed(1)}px;max-width:100%;height:auto;max-height:${mediaHeightPx.toFixed(
          1
      )}px"></video>`
      : null;
  const hasRenderableMedia = Boolean(imageTag || videoTag);
  const hasRenderableTextBlocks = textBlocks.length > 0;
  const hasVisibleText = sampleText.length > 0;
  const fallbackStage = `
    <div class="control-object-box" style="width:${widthPx.toFixed(1)}px;height:${heightPx.toFixed(1)}px;transform:rotate(${rotation}deg)">
      <span>${escapeHtml(graphicInfo.objectTypeName)}</span>
    </div>
  `;

  if (!PREVIEW_DEBUG && !hasRenderableMedia && !hasRenderableTextBlocks) {
    return "";
  }
  const renderText = hasVisibleText && (PREVIEW_DEBUG || (!hasRenderableMedia && !hasRenderableTextBlocks));

  const figureStyles = [...wrapStyles];
  if (!PREVIEW_DEBUG && hasRenderableMedia) {
    figureStyles.push("border:0");
    figureStyles.push("background:transparent");
    figureStyles.push("padding:0");
  }
  if (!PREVIEW_DEBUG && hasRenderableTextBlocks && !hasRenderableMedia) {
    figureStyles.push("border:0");
    figureStyles.push("background:transparent");
    figureStyles.push("padding:0");
    figureStyles.push("margin-top:0");
  }
  const stageStyles = [];
  if (!PREVIEW_DEBUG && hasRenderableMedia) {
    stageStyles.push("border:0");
    stageStyles.push("background:transparent");
    stageStyles.push("min-height:0");
    stageStyles.push("padding:0");
    stageStyles.push("display:flex");
    stageStyles.push("justify-content:center");
    stageStyles.push("align-items:flex-start");
  }
  const textPadding = graphicInfo.textBoxListInfo?.padding ?? [0, 0, 0, 0];
  const textPaddingPx = [
    clamp(hwpToPx(textPadding[0] ?? 0), 0, 120),
    clamp(hwpToPx(textPadding[1] ?? 0), 0, 120),
    clamp(hwpToPx(textPadding[2] ?? 0), 0, 120),
    clamp(hwpToPx(textPadding[3] ?? 0), 0, 120),
  ];
  const wrapRelaxEligible =
    !hasRenderableMedia &&
    hasRenderableTextBlocks &&
    widthPx >= 520 &&
    textBlocks.length > 0 &&
    textBlocks.length <= 3 &&
    textBlocks.some((block) => normalizePreviewText(block?.text ?? "").includes("/"));
  const textBoxWidthPx = wrapRelaxEligible ? clamp(widthPx * 1.14, widthPx, 680) : widthPx;
  const textVerticalAlign = graphicInfo.textBoxListInfo?.verticalAlignName;
  const textBoxJustify = textVerticalAlign === "middle" ? "center" : textVerticalAlign === "bottom" ? "flex-end" : "flex-start";
  const textBoxMargin =
    bits?.horzAlignName === "right" ? "margin-left:auto" : bits?.horzAlignName === "left" ? "margin-right:auto" : "margin:0 auto";
  const textBlocksHtml = hasRenderableTextBlocks
    ? textBlocks
        .map((block) => {
          const textRenderOptions = {
            tabStopsPx: block.tabStopsPx ?? [],
            baseStyle: block.style,
          };
          const blockTextHtml = block.textRuns?.length ? renderStyledTextRuns(block.textRuns, textRenderOptions) : textToHtml(block.text ?? "");
          const spacingAfterPx = clamp(block.style.spacingAfterPx ?? 0, 0, 28);
          const textCss = [
            `text-align:${block.style.textAlign}`,
            `font-family:${block.style.fontFamily}`,
            `font-size:${block.style.fontSizePt.toFixed(1)}pt`,
            `font-style:${block.style.fontStyle}`,
            `font-weight:${block.style.fontWeight}`,
            `color:${block.style.color}`,
            `line-height:${block.style.lineHeight}`,
            `margin:0 0 ${spacingAfterPx.toFixed(1)}px 0`,
            `text-indent:0`,
          ].join(";");
          return `<p class="preview-text control-object-paragraph" style="${escapeHtmlAttr(textCss)}">${blockTextHtml}</p>`;
        })
        .join("")
    : "";
  const textBoxHtml = hasRenderableTextBlocks
    ? `<div class="control-object-textbox" style="width:${textBoxWidthPx.toFixed(1)}px;min-height:${heightPx.toFixed(
        1
      )}px;padding:${textPaddingPx[2].toFixed(
        1
      )}px ${textPaddingPx[1].toFixed(1)}px ${textPaddingPx[3].toFixed(1)}px ${textPaddingPx[0].toFixed(1)}px;justify-content:${textBoxJustify};${textBoxMargin}">
        ${textBlocksHtml}
      </div>`
    : "";

  return `
    <figure class="control-object-wrap ${PREVIEW_DEBUG ? "" : "doc-object-wrap"}" style="${escapeHtmlAttr(figureStyles.join(";"))}">
      ${
        PREVIEW_DEBUG
          ? `<div class="control-object-meta">
              <span>${escapeHtml(graphicInfo.objectTypeName)}</span>
              <span>${escapeHtml(`offset ${xOffset},${yOffset}`)}</span>
              <span>${escapeHtml(`size ${width}x${height}`)}</span>
              ${flowLabel ? `<span>${escapeHtml(`flow ${flowLabel}`)}</span>` : ""}
              ${graphicInfo.pictureInfo?.binDataId ? `<span>${escapeHtml(`binData ${graphicInfo.pictureInfo.binDataId}`)}</span>` : ""}
              ${graphicInfo.videoInfo?.binDataId ? `<span>${escapeHtml(`videoBin ${graphicInfo.videoInfo.binDataId}`)}</span>` : ""}
              ${graphicInfo.chartInfo?.summary ? `<span>${escapeHtml(`chart ${graphicInfo.chartInfo.summary}`)}</span>` : ""}
              ${imageRef?.decodeStrategy ? `<span>${escapeHtml(`decode ${imageRef.decodeStrategy}`)}</span>` : ""}
              ${mediaRef?.decodeStrategy ? `<span>${escapeHtml(`media ${mediaRef.decodeStrategy}`)}</span>` : ""}
            </div>`
          : ""
      }
      ${
        PREVIEW_DEBUG || hasRenderableMedia
          ? `<div class="control-object-stage" style="${escapeHtmlAttr(stageStyles.join(";"))}">${imageTag || videoTag || fallbackStage}</div>`
          : ""
      }
      ${textBoxHtml}
      ${renderText ? `<div class="control-object-text">${textToHtml(sampleText.slice(0, 420))}</div>` : ""}
    </figure>
  `;
}

function parseGraphicControl(descendants, decoded, doc = state.doc, ctrlPayload = null) {
  const shapeComponentRecord = descendants.find((record) => record.tag === 76) ?? null;
  const shapeComponent = shapeComponentRecord
    ? parseShapeComponent(getRecordPayload(decoded, shapeComponentRecord))
    : null;

  const headerObjectCommon = ctrlPayload ? parseObjectCommon(ctrlPayload) : null;
  let objectCommon =
    headerObjectCommon && (headerObjectCommon.width !== 0 || headerObjectCommon.height !== 0) ? headerObjectCommon : null;
  if (!objectCommon) {
    for (const record of descendants) {
      if (!GRAPHIC_COMMON_RECORD_TAGS.has(record.tag)) {
        continue;
      }
      const candidate = parseObjectCommon(getRecordPayload(decoded, record));
      if (!candidate) {
        continue;
      }
      if (candidate.width === 0 && candidate.height === 0) {
        continue;
      }
      objectCommon = candidate;
      break;
    }
  }

  const shapeTagNames = descendants
    .filter((record) => GRAPHIC_DETAIL_RECORD_TAGS.has(record.tag))
    .map((record) => getRecordTagName(record.tag));
  const textBlocks = extractControlParagraphBlocks(descendants, decoded, doc);
  const textParagraphs = textBlocks.length ? textBlocks.map((block) => block.text) : extractControlParagraphTexts(descendants, decoded, doc);
  const textBoxListHeader = descendants.find((record) => record.tag === 72) ?? null;
  const textBoxListInfo = textBoxListHeader ? parseListHeader(getRecordPayload(decoded, textBoxListHeader)) : null;
  const objectTypeId = shapeComponent?.objectCtrlId || objectCommon?.ctrlId || "";
  const objectTypeName = SHAPE_OBJECT_CTRL_NAMES[objectTypeId] || objectTypeId || "Graphic";
  const pictureRecord = descendants.find((record) => record.tag === 85) ?? null;
  const chartRecord = descendants.find((record) => record.tag === 95) ?? null;
  const videoRecord = descendants.find((record) => record.tag === 98) ?? null;
  const textArtRecord = descendants.find((record) => record.tag === 90) ?? null;
  const formObjectRecord = descendants.find((record) => record.tag === 91) ?? null;
  const pictureInfo = pictureRecord ? parseShapePicture(getRecordPayload(decoded, pictureRecord), doc?.binDataById) : null;
  const chartInfo = chartRecord ? parseChartData(getRecordPayload(decoded, chartRecord)) : null;
  const videoInfo = videoRecord ? parseVideoData(getRecordPayload(decoded, videoRecord), doc?.binDataById) : null;
  const textArtInfo = textArtRecord ? parseTextArtData(getRecordPayload(decoded, textArtRecord)) : null;
  const formObjectInfo = formObjectRecord ? parseFormObjectData(getRecordPayload(decoded, formObjectRecord)) : null;
  const mediaRef = videoInfo?.binDataId ? resolveBinDataRef(videoInfo.binDataId, doc) : null;
  const pictureRef = pictureInfo?.binDataId ? resolveBinDataImageRef(pictureInfo.binDataId, doc) : null;
  const imageRef = pictureRef || (mediaRef?.mime?.startsWith("image/") ? mediaRef : null);

  return {
    shapeComponent,
    objectCommon,
    shapeTagNames,
    objectTypeId,
    objectTypeName,
    textParagraphs,
    textBlocks,
    textBoxListInfo,
    pictureInfo,
    chartInfo,
    videoInfo,
    textArtInfo,
    formObjectInfo,
    imageRef,
    mediaRef,
  };
}

function parseShapePicture(payload, knownBinDataMap = null) {
  if (!payload || payload.length < 73) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  let binDataId = dv.getUint16(71, true);
  if (knownBinDataMap && !knownBinDataMap.has(binDataId)) {
    const inferred = inferPictureBinDataId(payload, knownBinDataMap);
    if (inferred != null) {
      binDataId = inferred;
    }
  }

  return {
    brightness: dv.getInt8(68),
    contrast: dv.getInt8(69),
    effect: payload[70],
    binDataId,
    borderTransparency: payload.length >= 74 ? payload[73] : null,
    instanceId: payload.length >= 78 ? dv.getUint32(74, true) : null,
    payloadSize: payload.length,
  };
}

function parseChartData(payload) {
  if (!payload?.length) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const chartType = payload.length >= 4 ? dv.getUint32(0, true) : null;
  const chartId = payload.length >= 8 ? dv.getUint32(4, true) : null;
  const strings = findLengthPrefixedUtf16Strings(payload, 0).map((item) => item.text).slice(0, 5);
  const summaryText = strings[0] || extractReadableAscii(payload, 80) || "";
  return {
    chartType,
    chartId,
    strings,
    summary: summaryText || (chartType != null ? `type=${chartType}` : "chart"),
    payloadSize: payload.length,
  };
}

function parseVideoData(payload, knownBinDataMap = null) {
  if (!payload?.length) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const flags = payload.length >= 4 ? dv.getUint32(0, true) : null;
  const strings = findLengthPrefixedUtf16Strings(payload, 0).map((item) => item.text).slice(0, 6);
  const asciiPath = extractReadableAscii(payload, 180);

  let binDataId = null;
  if (knownBinDataMap?.size) {
    let best = null;
    for (let offset = 0; offset + 1 < payload.length; offset += 1) {
      const candidate = dv.getUint16(offset, true);
      if (!knownBinDataMap.has(candidate)) {
        continue;
      }
      const score = Math.abs(offset - 4);
      if (!best || score < best.score) {
        best = { id: candidate, score };
      }
    }
    binDataId = best?.id ?? null;
  }

  return {
    flags,
    binDataId,
    strings,
    mediaPath: strings[0] || asciiPath || "",
    payloadSize: payload.length,
  };
}

function parseTextArtData(payload) {
  if (!payload?.length) {
    return null;
  }
  const strings = findLengthPrefixedUtf16Strings(payload, 0).map((item) => item.text).slice(0, 3);
  const text = strings[0] || extractReadableAscii(payload, 100) || "";
  return {
    text,
    strings,
    payloadSize: payload.length,
  };
}

function parseFormObjectData(payload) {
  if (!payload?.length) {
    return null;
  }
  const strings = findLengthPrefixedUtf16Strings(payload, 0).map((item) => item.text).slice(0, 5);
  return {
    name: strings[0] || extractReadableAscii(payload, 80) || "",
    help: strings[1] || "",
    strings,
    payloadSize: payload.length,
  };
}

function inferPictureBinDataId(payload, knownBinDataMap) {
  if (!knownBinDataMap?.size || payload.length < 2) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  let best = null;
  for (let offset = 0; offset + 1 < payload.length; offset += 1) {
    const candidate = dv.getUint16(offset, true);
    if (!knownBinDataMap.has(candidate)) {
      continue;
    }
    const distance = Math.abs(offset - 71);
    if (!best || distance < best.distance) {
      best = {
        id: candidate,
        distance,
      };
    }
  }
  return best?.id ?? null;
}

function parseShapeComponent(payload) {
  if (!payload || payload.length < 46) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const rawId = bytesToAscii(payload.subarray(0, 4));
  const objectCtrlId = rawId ? rawId.split("").reverse().join("") : "";
  const groupOffsetX = dv.getInt32(4, true);
  const groupOffsetY = dv.getInt32(8, true);
  const groupLevel = dv.getUint16(12, true);
  const localVersion = dv.getUint16(14, true);
  const initialWidth = dv.getUint32(16, true);
  const initialHeight = dv.getUint32(20, true);
  const currentWidth = dv.getUint32(24, true);
  const currentHeight = dv.getUint32(28, true);
  const property = dv.getUint32(32, true);
  const rotation = dv.getInt16(36, true);
  const rotationCenterX = dv.getInt32(38, true);
  const rotationCenterY = dv.getInt32(42, true);
  let matrixPairCount = null;
  if (payload.length >= 48) {
    matrixPairCount = dv.getUint16(46, true);
  }

  return {
    objectCtrlId,
    groupOffsetX,
    groupOffsetY,
    groupLevel,
    localVersion,
    initialWidth,
    initialHeight,
    currentWidth,
    currentHeight,
    property,
    rotation,
    rotationCenterX,
    rotationCenterY,
    matrixPairCount,
  };
}

function renderTableControlPreview(tableInfo, splitInfo = null) {
  const grid = buildTableGrid(tableInfo.rows, tableInfo.cols, tableInfo.cells);
  if (!grid.rows.length || !grid.cols.length) {
    return `<div class="control-table-empty">Unable to build table grid.</div>`;
  }

  const colgroup = grid.cols
    .map((width) => {
      if (!width || width <= 0) {
        return "<col>";
      }
      const clamped = clamp(width / 80, 20, 360);
      return `<col style="width:${clamped.toFixed(1)}px">`;
    })
    .join("");

  const bodyRows = grid.rows
    .map((row, rowIndex) => {
      const cells = row
        .map((cell) => {
          if (!cell) {
            return `<td class="control-table-cell empty">${PREVIEW_DEBUG ? "" : "&nbsp;"}</td>`;
          }
          if (cell.covered) {
            return "";
          }

          const spanAttrs = [
            cell.colSpan > 1 ? `colspan="${cell.colSpan}"` : "",
            cell.rowSpan > 1 ? `rowspan="${cell.rowSpan}"` : "",
          ]
            .filter(Boolean)
            .join(" ");
          const cellText = cell.text ? textToHtml(cell.text.slice(0, 1024)) : "";
          if (!PREVIEW_DEBUG) {
            return `
              <td class="control-table-cell" ${spanAttrs}>
                ${cellText ? `<div class="control-table-cell-text">${cellText}</div>` : "&nbsp;"}
              </td>
            `;
          }

          const spanLabel = cell.colSpan > 1 || cell.rowSpan > 1 ? `span ${cell.colSpan}x${cell.rowSpan}` : "span 1x1";
          const widthLabel = cell.width ? `w ${cell.width}` : "w -";
          const heightLabel = cell.height ? `h ${cell.height}` : "h -";
          const paraLabel = `p ${cell.paragraphCount}`;

          return `
            <td class="control-table-cell" ${spanAttrs}>
              <div class="control-table-cell-pos">R${cell.row + 1}C${cell.col + 1}</div>
              <div class="control-table-cell-meta">${escapeHtml(`${spanLabel} · ${widthLabel} · ${heightLabel} · ${paraLabel}`)}</div>
              ${cellText ? `<div class="control-table-cell-text">${cellText}</div>` : ""}
            </td>
          `;
        })
        .join("");
      return `<tr data-row="${rowIndex + 1}">${cells}</tr>`;
    })
    .join("");

  return `
    <div class="control-table-wrap ${PREVIEW_DEBUG ? "" : "doc-table-wrap"}">
      ${
        splitInfo
          ? `<div class="control-table-split-note">${escapeHtml(
              `rows ${splitInfo.startRow + 1}-${splitInfo.endRow} / ${splitInfo.totalRows}`
            )}</div>`
          : ""
      }
      ${
        PREVIEW_DEBUG
          ? `<div class="control-table-meta">
              <span>rows ${tableInfo.rows}</span>
              <span>cols ${tableInfo.cols}</span>
              <span>cells ${tableInfo.cells.length}</span>
              ${
                splitInfo
                  ? `<span>${escapeHtml(`split ${splitInfo.startRow + 1}-${splitInfo.endRow}/${splitInfo.totalRows}`)}</span>`
                  : ""
              }
            </div>`
          : ""
      }
      <table class="control-table-grid">
        <colgroup>${colgroup}</colgroup>
        <tbody>${bodyRows}</tbody>
      </table>
    </div>
  `;
}

function buildTableGrid(rowCount, colCount, cells) {
  if (!Number.isFinite(rowCount) || !Number.isFinite(colCount) || rowCount <= 0 || colCount <= 0) {
    return { rows: [], cols: [] };
  }

  const rows = Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => null));
  const cols = Array.from({ length: colCount }, () => 0);
  const sorted = [...cells].sort((a, b) => a.row - b.row || a.col - b.col);

  for (const cell of sorted) {
    if (cell.row < 0 || cell.col < 0 || cell.row >= rowCount || cell.col >= colCount) {
      continue;
    }
    if (rows[cell.row][cell.col]) {
      continue;
    }

    rows[cell.row][cell.col] = cell;
    if (cell.colSpan === 1 && cell.width > 0) {
      cols[cell.col] = Math.max(cols[cell.col], cell.width);
    }

    for (let rr = cell.row; rr < Math.min(rowCount, cell.row + cell.rowSpan); rr += 1) {
      for (let cc = cell.col; cc < Math.min(colCount, cell.col + cell.colSpan); cc += 1) {
        if (rr === cell.row && cc === cell.col) {
          continue;
        }
        if (!rows[rr][cc]) {
          rows[rr][cc] = { covered: true };
        }
      }
    }
  }

  const fallbackWidth = cols.every((value) => value <= 0) ? 2100 : 0;
  for (let i = 0; i < cols.length; i += 1) {
    if (cols[i] <= 0) {
      cols[i] = fallbackWidth;
    }
  }

  return { rows, cols };
}

function parseTableControl(payload) {
  if (!payload || payload.length < 8) {
    return null;
  }

  // HWPTAG_TABLE payload itself starts with table body fields (no ObjectCommon prefix).
  const tableBody = parseTablePropertiesAt(payload, 0);
  if (tableBody) {
    return {
      ...tableBody,
      objectInfo: null,
    };
  }

  return parseCompactTableControl(payload);
}

function parseCompactTableControl(payload) {
  if (!payload || payload.length < 8) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const property = payload.length >= 4 ? dv.getUint32(0, true) : 0;
  const rows = payload.length >= 6 ? dv.getUint16(4, true) : 0;
  const cols = payload.length >= 8 ? dv.getUint16(6, true) : 0;
  if (rows <= 0 || cols <= 0 || rows > 4096 || cols > 1024) {
    return null;
  }
  return {
    property,
    rows,
    cols,
    cellSpacing: 0,
    innerMargin: [0, 0, 0, 0],
    rowSizes: [],
    borderFillId: 0,
    zones: [],
    cells: [],
    parseOffset: 0,
    payloadOffset: 0,
    tailBytes: 0,
    cellAddressBase: 0,
    objectInfo: null,
  };
}

function buildTableInfoFromListHeaders(listHeaders, tableInfo = null) {
  if (!listHeaders?.length) {
    return tableInfo;
  }
  const expectedRows = tableInfo?.rows && tableInfo.rows > 0 ? tableInfo.rows : 4096;
  const expectedCols = tableInfo?.cols && tableInfo.cols > 0 ? tableInfo.cols : 4096;
  const cellEntries = listHeaders
    .map((header) => header.cell)
    .filter(Boolean)
    .filter(
      (cell) =>
        cell.col >= 0 &&
        cell.row >= 0 &&
        cell.col < Math.max(8192, expectedCols * 8) &&
        cell.row < Math.max(8192, expectedRows * 8) &&
        cell.colSpan > 0 &&
        cell.rowSpan > 0 &&
        cell.colSpan <= 2048 &&
        cell.rowSpan <= 2048
    );
  if (!cellEntries.length) {
    return tableInfo;
  }

  const normalized = normalizeTableCellAddresses(
    cellEntries,
    expectedRows,
    expectedCols
  );
  const cells = normalized.cells;
  const rowsByCells = cells.reduce((max, cell) => Math.max(max, cell.row + Math.max(1, cell.rowSpan)), 0);
  const colsByCells = cells.reduce((max, cell) => Math.max(max, cell.col + Math.max(1, cell.colSpan)), 0);
  const rows = Math.max(tableInfo?.rows ?? 0, rowsByCells);
  const cols = Math.max(tableInfo?.cols ?? 0, colsByCells);
  const innerMargin =
    tableInfo?.innerMargin && tableInfo.innerMargin.length === 4
      ? tableInfo.innerMargin
      : cells[0]?.margin?.length === 4
        ? [...cells[0].margin]
        : [0, 0, 0, 0];
  const borderFillId =
    tableInfo?.borderFillId && tableInfo.borderFillId > 0
      ? tableInfo.borderFillId
      : cells.find((cell) => cell.borderFillId > 0)?.borderFillId ?? 0;

  return {
    ...(tableInfo ?? {}),
    property: tableInfo?.property ?? 0,
    rows,
    cols,
    cellSpacing: tableInfo?.cellSpacing ?? 0,
    innerMargin,
    rowSizes: tableInfo?.rowSizes ?? [],
    borderFillId,
    zones: tableInfo?.zones ?? [],
    cells,
    cellAddressBase: normalized.addressBase,
  };
}

function parseObjectCommon(payload, offset = 0) {
  if (offset + 46 > payload.length) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const rawCtrl = bytesToAscii(payload.subarray(offset, offset + 4));
  const ctrlId = rawCtrl ? rawCtrl.split("").reverse().join("") : "";
  const property = dv.getUint32(offset + 4, true);
  const propertyBits = decodeObjectCommonProperty(property);
  const yOffset = dv.getInt32(offset + 8, true);
  const xOffset = dv.getInt32(offset + 12, true);
  const width = dv.getInt32(offset + 16, true);
  const height = dv.getInt32(offset + 20, true);
  const zOrder = dv.getInt32(offset + 24, true);
  const outerMargin = [
    dv.getInt16(offset + 28, true),
    dv.getInt16(offset + 30, true),
    dv.getInt16(offset + 32, true),
    dv.getInt16(offset + 34, true),
  ];
  const instanceId = dv.getUint32(offset + 36, true);
  const lockToPage = dv.getInt32(offset + 40, true);
  const descLen = dv.getUint16(offset + 44, true);
  const descStart = offset + 46;
  const descEnd = descStart + descLen * 2;
  const description =
    descLen > 0 && descEnd <= payload.length
      ? UTF16LE_DECODER.decode(payload.subarray(descStart, descEnd)).replace(/\0/g, "")
      : "";

  return {
    ctrlId,
    property,
    propertyBits,
    xOffset,
    yOffset,
    width,
    height,
    zOrder,
    outerMargin,
    instanceId,
    lockToPage,
    description,
  };
}

function decodeObjectCommonProperty(property) {
  const value = property >>> 0;
  const likeCharacter = Boolean(value & (1 << 0));
  const affectLineSpacing = Boolean(value & (1 << 2));
  const vertRelTo = bitsValue(value, 3, 2);
  const vertAlign = bitsValue(value, 5, 3);
  const horzRelTo = bitsValue(value, 8, 2);
  const horzAlign = bitsValue(value, 10, 3);
  const restrictInBody = Boolean(value & (1 << 13));
  const allowOverlap = Boolean(value & (1 << 14));
  const widthRelTo = bitsValue(value, 15, 3);
  const heightRelTo = bitsValue(value, 18, 2);
  const protectSize = Boolean(value & (1 << 20));
  const textFlow = bitsValue(value, 21, 3);
  const textSide = bitsValue(value, 24, 2);
  const numberCategory = bitsValue(value, 26, 3);

  return {
    likeCharacter,
    affectLineSpacing,
    vertRelTo,
    vertRelToName: OBJECT_VERT_REL_TO_NAMES[vertRelTo] ?? `vRel${vertRelTo}`,
    vertAlign,
    vertAlignName: OBJECT_VERT_ALIGN_NAMES[vertAlign] ?? `vAlign${vertAlign}`,
    horzRelTo,
    horzRelToName: OBJECT_HORZ_REL_TO_NAMES[horzRelTo] ?? `hRel${horzRelTo}`,
    horzAlign,
    horzAlignName: OBJECT_HORZ_ALIGN_NAMES[horzAlign] ?? `hAlign${horzAlign}`,
    restrictInBody,
    allowOverlap,
    widthRelTo,
    heightRelTo,
    protectSize,
    textFlow,
    textSide,
    numberCategory,
    textFlowName: OBJECT_TEXT_FLOW_NAMES[textFlow] ?? `flow${textFlow}`,
    textSideName: OBJECT_TEXT_SIDE_NAMES[textSide] ?? `side${textSide}`,
  };
}

function parseTablePropertiesAt(payload, offset) {
  if (offset + 22 > payload.length) {
    return null;
  }

  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const property = dv.getUint32(offset, true);
  const rows = dv.getUint16(offset + 4, true);
  const cols = dv.getUint16(offset + 6, true);
  const cellSpacing = dv.getInt16(offset + 8, true);
  const innerMargin = [
    dv.getInt16(offset + 10, true),
    dv.getInt16(offset + 12, true),
    dv.getInt16(offset + 14, true),
    dv.getInt16(offset + 16, true),
  ];
  if (rows <= 0 || cols <= 0 || rows > 4096 || cols > 1024) {
    return null;
  }

  let cursor = offset + 18;
  if (cursor + rows * 2 + 4 > payload.length) {
    return null;
  }

  const rowSizes = [];
  for (let i = 0; i < rows; i += 1) {
    rowSizes.push(dv.getInt16(cursor, true));
    cursor += 2;
  }

  const borderFillId = dv.getUint16(cursor, true);
  cursor += 2;
  const zoneCount = dv.getUint16(cursor, true);
  cursor += 2;
  if (zoneCount > 16384 || cursor + zoneCount * 10 > payload.length) {
    return null;
  }

  const zones = [];
  for (let i = 0; i < zoneCount; i += 1) {
    const startCol = dv.getUint16(cursor, true);
    const startRow = dv.getUint16(cursor + 2, true);
    const endCol = dv.getUint16(cursor + 4, true);
    const endRow = dv.getUint16(cursor + 6, true);
    const zoneBorderFillId = dv.getUint16(cursor + 8, true);
    zones.push({
      startCol,
      startRow,
      endCol,
      endRow,
      borderFillId: zoneBorderFillId,
    });
    cursor += 10;
  }

  const remaining = payload.length - cursor;
  if (remaining < 0) {
    return null;
  }

  const cellCount = Math.floor(remaining / 32);
  if (cellCount < 0) {
    return null;
  }
  const tailBytes = remaining - cellCount * 32;
  const cells = [];
  for (let i = 0; i < cellCount; i += 1) {
    const base = cursor + i * 32;
    const paragraphCount = dv.getInt16(base, true);
    const listProperty = dv.getUint32(base + 2, true);
    const col = dv.getUint16(base + 6, true);
    const row = dv.getUint16(base + 8, true);
    const colSpanRaw = dv.getUint16(base + 10, true);
    const rowSpanRaw = dv.getUint16(base + 12, true);
    const width = dv.getInt32(base + 14, true);
    const height = dv.getInt32(base + 18, true);
    const margin = [
      dv.getInt16(base + 22, true),
      dv.getInt16(base + 24, true),
      dv.getInt16(base + 26, true),
      dv.getInt16(base + 28, true),
    ];
    const cellBorderFillId = dv.getUint16(base + 30, true);

    cells.push({
      paragraphCount: Math.max(0, paragraphCount),
      listProperty,
      col,
      row,
      colSpan: colSpanRaw > 0 ? colSpanRaw : 1,
      rowSpan: rowSpanRaw > 0 ? rowSpanRaw : 1,
      width,
      height,
      margin,
      borderFillId: cellBorderFillId,
    });
  }

  const normalized = normalizeTableCellAddresses(cells, rows, cols);

  return {
    property,
    rows,
    cols,
    cellSpacing,
    innerMargin,
    rowSizes,
    borderFillId,
    zones,
    cells: normalized.cells,
    cellAddressBase: normalized.addressBase,
    parseOffset: offset,
    tailBytes,
    payloadOffset: offset,
  };
}

function normalizeTableCellAddresses(cells, rows, cols) {
  const zeroBasedHits = cells.filter((cell) => cell.row >= 0 && cell.col >= 0 && cell.row < rows && cell.col < cols).length;
  const oneBasedHits = cells.filter((cell) => cell.row >= 1 && cell.col >= 1 && cell.row <= rows && cell.col <= cols).length;
  const useOneBased = oneBasedHits > zeroBasedHits * 1.2 && oneBasedHits >= Math.max(1, Math.floor(cells.length * 0.4));
  if (!useOneBased) {
    return {
      cells,
      addressBase: 0,
    };
  }
  return {
    addressBase: 1,
    cells: cells.map((cell) => ({
      ...cell,
      row: Math.max(0, cell.row - 1),
      col: Math.max(0, cell.col - 1),
    })),
  };
}

function attachTableCellTexts(tableInfo, descendants, decoded) {
  const paragraphs = extractTableParagraphTexts(descendants, decoded);
  const fallbackParagraphs = paragraphs.length ? paragraphs : extractControlParagraphTexts(descendants, decoded);
  const pool = fallbackParagraphs;
  let cursor = 0;
  const cells = tableInfo.cells.map((cell) => {
    const paragraphCount = Math.max(0, cell.paragraphCount);
    const selected = pool.slice(cursor, cursor + paragraphCount);
    cursor += paragraphCount;
    const text = selected
      .map((value) => normalizePreviewText(value))
      .filter((value) => value.length > 0)
      .join("\n");
    return {
      ...cell,
      text,
    };
  });

  return {
    ...tableInfo,
    cells,
    paragraphPool: pool.length,
    mappedParagraphs: Math.min(cursor, pool.length),
  };
}

function extractTableParagraphTexts(descendants, decoded) {
  const tableIndex = descendants.findIndex((record) => record.tag === 77);
  const records = tableIndex >= 0 ? descendants.slice(tableIndex + 1) : descendants;
  if (!records.length) {
    return [];
  }
  return extractControlParagraphTexts(records, decoded);
}

function extractControlParagraphBlocks(descendants, decoded, doc = state.doc) {
  if (!descendants?.length) {
    return [];
  }

  const paraHeaderLevels = descendants.filter((record) => record.tag === 66 && Number.isFinite(record.level)).map((record) => record.level);
  if (!paraHeaderLevels.length) {
    return [];
  }

  const baseLevel = Math.min(...paraHeaderLevels);
  const directLevel = baseLevel + 1;
  const paragraphs = [];
  let current = null;

  for (const record of descendants) {
    if (record.tag === 66 && record.level === baseLevel) {
      current = {
        headerRecord: record,
        textRecord: null,
        charShapeRecord: null,
        lineSegRecord: null,
      };
      paragraphs.push(current);
      continue;
    }

    if (!current || record.level !== directLevel) {
      continue;
    }

    if (record.tag === 67 && !current.textRecord) {
      current.textRecord = record;
    } else if (record.tag === 68 && !current.charShapeRecord) {
      current.charShapeRecord = record;
    } else if (record.tag === 69 && !current.lineSegRecord) {
      current.lineSegRecord = record;
    }
  }

  const blocks = [];
  const docInfo = doc?.docInfo ?? null;

  for (const paragraph of paragraphs) {
    if (!paragraph.textRecord) {
      continue;
    }
    const paragraphTextData = decodeParagraphText(getRecordPayload(decoded, paragraph.textRecord));
    const visibleTokens = (paragraphTextData.tokens ?? []).filter((token) => token.type !== "control");
    const text = tokensToPlainText(visibleTokens);
    if (!text.length) {
      continue;
    }

    const header = paragraph.headerRecord ? parseParaHeader(getRecordPayload(decoded, paragraph.headerRecord)) : null;
    const charRuns = paragraph.charShapeRecord ? parseParaCharShape(getRecordPayload(decoded, paragraph.charShapeRecord)) : [];
    const styleRef = header?.paraStyleId != null ? docInfo?.styleById?.get(header.paraStyleId) ?? null : null;
    let charShapeId = charRuns.length ? charRuns[0].charShapeId : null;
    if (charShapeId == null && styleRef) {
      charShapeId = styleRef.charShapeId ?? null;
    }
    const charShape = charShapeId != null ? docInfo?.charShapeById?.get(charShapeId) ?? null : null;
    const paraShape = header?.paraShapeId != null ? docInfo?.paraShapeById?.get(header.paraShapeId) ?? null : null;
    const style = resolvePreviewStyle(paraShape, charShape);
    const textRuns = buildStyledTextRuns(visibleTokens, charRuns, docInfo, style, charShapeId);
    const lineSegments = paragraph.lineSegRecord ? parseParaLineSeg(getRecordPayload(decoded, paragraph.lineSegRecord)) : [];
    const textLines =
      lineSegments.length > 0 ? buildStyledTextLines(visibleTokens, charRuns, lineSegments, docInfo, style, charShapeId) : [];

    blocks.push({
      text,
      style,
      textRuns,
      textLines,
      tabStopsPx: [],
    });
  }

  return blocks;
}

function extractControlParagraphTexts(descendants, decoded, doc = state.doc) {
  return extractControlParagraphBlocks(descendants, decoded, doc).map((paragraph) => paragraph.text ?? "");
}

function getPreviewSectionStreams(entries, fileHeader = null) {
  const useViewText = Boolean(fileHeader?.distributable);
  const pick = (pattern) =>
    entries
      .filter((entry) => entry.type === 2 && pattern.test(entry.fullPath))
      .sort((a, b) => getSectionNumber(a.fullPath) - getSectionNumber(b.fullPath));

  if (useViewText) {
    const view = pick(/\/ViewText\/Section\d+$/);
    if (view.length) {
      return view;
    }
  }
  return pick(/\/BodyText\/Section\d+$/);
}

function getSectionNumber(path) {
  const match = path.match(/Section(\d+)$/);
  if (!match) {
    return Number.MAX_SAFE_INTEGER;
  }
  return Number(match[1]);
}

function toSectionLabel(path) {
  const sectionNumber = getSectionNumber(path);
  if (!Number.isFinite(sectionNumber)) {
    return path;
  }
  return `Section${sectionNumber}`;
}

function resolvePreviewStyle(paraShape, charShape) {
  const alignValue = paraShape?.property1Bits?.align;
  const textAlign =
    alignValue === 2
      ? "right"
      : alignValue === 3
        ? "center"
        : alignValue === 1
          ? "left"
          : "justify";

  const fontSizePt =
    charShape?.baseSize != null && charShape.baseSize > 0
      ? clamp((charShape.baseSize / 100) * PREVIEW_FONT_SCALE, 6, 72)
      : 10;

  const lineHeightRatio = paraShape?.lineSpacing != null && paraShape.lineSpacing > 0
    ? clamp((paraShape.lineSpacing / 100) / fontSizePt, 1.05, 2.4)
    : 1.6;

  const marginLeftPx = paraShape?.leftMargin != null ? clamp(hwpToPx(paraShape.leftMargin), -120, 260) : 0;
  const marginRightPx = paraShape?.rightMargin != null ? clamp(hwpToPx(paraShape.rightMargin), -120, 260) : 0;
  const textIndentPx = paraShape?.indent != null ? clamp(hwpToPx(paraShape.indent), -180, 220) : 0;
  const spacingBeforePx = paraShape?.spacingBefore != null ? clamp(hwpToPx(paraShape.spacingBefore), 0, 160) : 0;
  const spacingAfterPx = paraShape?.spacingAfter != null ? clamp(hwpToPx(paraShape.spacingAfter), 0, 160) : 0;

  return {
    textAlign,
    fontFamily: toCssFontFamily(charShape?.primaryFont),
    fontSizePt,
    fontWeight: charShape?.attributeBits?.bold ? 600 : 400,
    fontStyle: charShape?.attributeBits?.italic ? "italic" : "normal",
    color: !charShape?.textColor || charShape.textColor.auto ? "#111111" : charShape.textColor.hex,
    lineHeight: lineHeightRatio.toFixed(2),
    marginLeftPx,
    marginRightPx,
    textIndentPx,
    spacingBeforePx,
    spacingAfterPx,
  };
}

function toCssFontFamily(name) {
  const fallback = '"Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif';
  if (!name) {
    return fallback;
  }
  const sanitized = name.replace(/["\\]/g, "").trim();
  if (!sanitized) {
    return fallback;
  }
  return `"${sanitized}", ${fallback}`;
}

function normalizePreviewText(value) {
  if (!value) {
    return "";
  }
  return normalizeDisplayText(value)
    .replace(/\[CTRL:[^\]]+\]/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeDisplayText(value) {
  if (!value) {
    return "";
  }
  return String(value).replaceAll("", "한").replaceAll("ᄒᆞᆫ", "한").normalize("NFC");
}

function textToHtml(value) {
  return escapeHtml(value).replaceAll("\n", "<br>");
}

function renderStreamList() {
  renderStreamListPanel({
    streamListEl,
    doc: state.doc,
    selectedEntryIndex: state.selectedEntryIndex,
    formatBytes,
    escapeHtml,
    onSelectEntry: (entryIndex) => {
      state.selectedEntryIndex = entryIndex;
      state.selectedRecordIndex = null;
      renderStreamList();
      renderRecordPanel();
    },
  });
}

function renderRecordPanel() {
  const stream = getSelectedStream();
  const analysis = stream ? getStreamAnalysis(stream) : null;

  renderRecordPanelView({
    recordPanelEl,
    detailPanelEl,
    stream,
    analysis,
    selectedRecordIndex: state.selectedRecordIndex,
    getRecordTagName,
    getParagraphIndexByRecord,
    formatBytes,
    formatHex,
    escapeHtml,
    onSelectRecord: (recordIndex) => {
      state.selectedRecordIndex = recordIndex;
      renderRecordPanel();
    },
    renderDetailPanel,
  });
}

function renderDetailPanel(stream, analysis) {
  const records = analysis.records;
  const selectedRecord =
    state.selectedRecordIndex != null && records[state.selectedRecordIndex]
      ? records[state.selectedRecordIndex]
      : records[0];
  const docInfo = state.doc?.docInfo ?? null;

  if (!selectedRecord) {
    detailPanelEl.innerHTML = `<div class="empty">No record selected.</div>`;
    return;
  }

  const payload = getRecordPayload(analysis.decoded, selectedRecord);
  const payloadHex = formatHex(payload, selectedRecord.payloadOffset, 1536);
  const tagName = getRecordTagName(selectedRecord.tag);

  const paraText = selectedRecord.tag === 67 ? decodeParagraphText(payload) : null;
  const paragraph = getParagraphByRecord(analysis, selectedRecord.index);
  const paragraphIndex = paragraph?.index ?? null;
  const paragraphHeader = paragraph?.headerRecord
    ? parseParaHeader(getRecordPayload(analysis.decoded, paragraph.headerRecord))
    : null;
  const paragraphText = paragraph?.textRecord ? decodeParagraphText(getRecordPayload(analysis.decoded, paragraph.textRecord)) : null;
  const charShapeRuns = paragraph?.charShapeRecord
    ? parseParaCharShape(getRecordPayload(analysis.decoded, paragraph.charShapeRecord))
    : [];
  const lineSegments = paragraph?.lineSegRecord
    ? parseParaLineSeg(getRecordPayload(analysis.decoded, paragraph.lineSegRecord))
    : [];
  const ctrlHeaders = paragraph?.ctrlHeaderRecords?.length
    ? paragraph.ctrlHeaderRecords.map((record) => parseCtrlHeader(getRecordPayload(analysis.decoded, record), record))
    : [];
  const streamText = extractParagraphText(analysis.decoded, records);
  const advancedTagInfo = decodeAdvancedTagInfo(selectedRecord.tag, payload);

  detailPanelEl.innerHTML = `
    <div class="record-head">
      <div><span>Record</span><strong>#${selectedRecord.index}</strong></div>
      <div><span>Tag</span><strong>${selectedRecord.tag} (${escapeHtml(tagName)})</strong></div>
      <div><span>Paragraph</span><strong>${paragraphIndex != null ? `#${paragraphIndex}` : "-"}</strong></div>
      <div><span>Level</span><strong>${selectedRecord.level}</strong></div>
      <div><span>Payload Size</span><strong>${selectedRecord.size}</strong></div>
    </div>

    ${
      paragraphHeader
        ? `<section class="text-block">
            <h3>Paragraph Header (66)</h3>
            <pre>${escapeHtml(renderParagraphHeader(paragraphHeader, docInfo))}</pre>
          </section>`
        : ""
    }

    ${
      paragraphText
        ? `<section class="text-block">
            <h3>Paragraph Text (paragraph context)</h3>
            <pre>${escapeHtml(paragraphText.text || "(empty)")}</pre>
          </section>`
        : ""
    }

    ${
      paraText
        ? `<section class="text-block">
            <h3>Paragraph Text (current record)</h3>
            <pre>${escapeHtml(paraText.text || "(empty)")}</pre>
          </section>`
        : ""
    }

    ${
      charShapeRuns.length
        ? `<section class="text-block">
            <h3>PARA_CHAR_SHAPE (68)</h3>
            <div class="record-table-wrap">
              <table class="record-table">
                <thead>
                  <tr>
                    <th>Run</th>
                    <th>StartPos</th>
                    <th>ShapeId</th>
                    <th>Resolved</th>
                  </tr>
                </thead>
                <tbody>
                  ${charShapeRuns
                    .map(
                      (run, idx) => `
                        <tr>
                          <td>${idx}</td>
                          <td>${run.startPos}</td>
                          <td>${run.charShapeId}</td>
                          <td>${escapeHtml(describeCharShape(run.charShapeId, docInfo))}</td>
                        </tr>
                      `
                    )
                    .join("")}
                </tbody>
              </table>
            </div>
          </section>`
        : ""
    }

    ${
      lineSegments.length
        ? `<section class="text-block">
            <h3>PARA_LINE_SEG (69)</h3>
            <div class="record-table-wrap">
              <table class="record-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>TextPos</th>
                    <th>VertPos</th>
                    <th>Height</th>
                    <th>Flags</th>
                  </tr>
                </thead>
                <tbody>
                  ${lineSegments
                    .map(
                      (seg, idx) => `
                        <tr>
                          <td>${idx}</td>
                          <td>${seg.textStartPos}</td>
                          <td>${seg.lineVerticalPos}</td>
                          <td>${seg.lineHeight}</td>
                          <td>${escapeHtml(seg.flagLabels.join("|") || "-")}</td>
                        </tr>
                      `
                    )
                    .join("")}
                </tbody>
              </table>
            </div>
          </section>`
        : ""
    }

    ${
      ctrlHeaders.length
        ? `<section class="text-block">
            <h3>CTRL_HEADER (71)</h3>
            <div class="record-table-wrap">
              <table class="record-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>CtrlID</th>
                    <th>Name</th>
                    <th>Level</th>
                    <th>Size</th>
                    <th>Words</th>
                  </tr>
                </thead>
                <tbody>
                  ${ctrlHeaders
                    .map(
                      (ctrl, idx) => `
                        <tr>
                          <td>${idx}</td>
                          <td>${escapeHtml(ctrl.ctrlId || "-")}</td>
                          <td>${escapeHtml(ctrl.name)}</td>
                          <td>${ctrl.level}</td>
                          <td>${ctrl.payloadSize}</td>
                          <td>${escapeHtml(ctrl.wordsText)}</td>
                        </tr>
                      `
                    )
                    .join("")}
                </tbody>
              </table>
            </div>
          </section>`
        : ""
    }

    ${
      advancedTagInfo
        ? `<section class="text-block">
            <h3>${escapeHtml(advancedTagInfo.title)}</h3>
            <pre>${escapeHtml(advancedTagInfo.text)}</pre>
          </section>`
        : ""
    }

    ${
      streamText.text
        ? `<section class="text-block">
            <h3>Paragraph Text (stream)</h3>
            <pre>${escapeHtml(streamText.text)}</pre>
          </section>`
        : ""
    }

    <section class="text-block">
      <h3>Payload Hex</h3>
      <pre class="hex">${escapeHtml(payloadHex)}</pre>
    </section>

    <section class="text-block">
      <h3>Stream Notes</h3>
      <pre>${escapeHtml(streamParseNote(stream, analysis))}</pre>
    </section>
  `;
}

function getSelectedStream() {
  const doc = state.doc;
  if (!doc || state.selectedEntryIndex == null) {
    return null;
  }
  return doc.entries.find((entry) => entry.index === state.selectedEntryIndex) ?? null;
}

function getStreamAnalysis(stream) {
  const cacheKey = stream.fullPath;
  const cached = state.doc.streamAnalysis.get(cacheKey);
  if (cached) {
    return cached;
  }

  const header = state.doc.fileHeader;
  const compressed = Boolean(header?.compressed);
  const distributable = Boolean(header?.distributable);
  const analysis = analyzeStream(stream, compressed, state.doc.formatInfo?.mode || "unknown", distributable);
  state.doc.streamAnalysis.set(cacheKey, analysis);
  return analysis;
}

function analyzeStream(stream, compressedFlag, formatMode = "unknown", distributableFlag = false) {
  const raw = stream.content;
  if (!raw || raw.length === 0) {
    return {
      mode: "raw",
      decoded: new Uint8Array(),
      parse: { records: [], consumed: 0, complete: true },
      records: [],
      paragraphs: [],
      recordToParagraph: new Map(),
    };
  }

  if (distributableFlag && isDistributableEncryptedStream(stream.fullPath)) {
    const distCandidates = decodeDistributableStreamCandidates(raw, compressedFlag);
    if (distCandidates.length) {
      return selectBestStreamAnalysis(distCandidates, stream.fullPath);
    }
  }

  if (formatMode === "legacy-3x-binary" && !isLikelyRecordStreamPath(stream.fullPath)) {
    return analyzeLegacy3Stream(raw);
  }

  const candidates = [{ mode: "raw", bytes: raw }];
  if (compressedFlag && mayBeCompressedRecordStreamPath(stream.fullPath)) {
    const rawInflated = safeInflate(raw, true);
    if (rawInflated) {
      candidates.push({ mode: "inflateRaw(-15)", bytes: rawInflated });
    }

    const zlibInflated = safeInflate(raw, false);
    if (zlibInflated) {
      candidates.push({ mode: "inflate(zlib)", bytes: zlibInflated });
    }
  }

  const best = selectBestStreamAnalysis(candidates, stream.fullPath);

  if (best && best.records.length === 0 && formatMode === "legacy-3x-binary") {
    return analyzeLegacy3Stream(raw);
  }

  return best;
}

function selectBestStreamAnalysis(candidates, streamPath) {
  let best = null;
  for (const candidate of candidates) {
    const parse = isLikelyRecordStreamPath(streamPath)
      ? parseRecords(candidate.bytes)
      : { records: [], consumed: 0, complete: false };
    const paragraphContext = parse.records.length
      ? buildParagraphContext(parse.records)
      : { paragraphs: [], recordToParagraph: new Map() };

    const score = parse.records.length + paragraphContext.paragraphs.length + (parse.complete ? 10 : 0);
    if (!best || score > best.score || (score === best.score && candidate.bytes.length > best.decoded.length)) {
      best = {
        mode: candidate.mode,
        decoded: candidate.bytes,
        parse,
        records: parse.records,
        paragraphs: paragraphContext.paragraphs,
        recordToParagraph: paragraphContext.recordToParagraph,
        score,
      };
    }
  }
  return best;
}

function isDistributableEncryptedStream(path) {
  return /\/ViewText\/Section\d+$/.test(path) || /\/Scripts\/[^/]+$/.test(path);
}

function decodeDistributableStreamCandidates(rawBytes, compressedFlag) {
  const head = readRecordAt(rawBytes, 0);
  if (!head || head.tag !== DISTRIBUTE_DOC_RECORD_TAG || head.size !== 256) {
    return [];
  }

  const key = decodeDistributeDocHeadToKey(rawBytes.subarray(head.payloadOffset, head.payloadOffset + head.size));
  if (!key || key.length !== 16) {
    return [];
  }

  const tail = rawBytes.subarray(head.nextOffset);
  if (!tail.length) {
    return [];
  }

  const alignedLength = tail.length - (tail.length % 16);
  if (alignedLength <= 0) {
    return [];
  }
  const encryptedTail = tail.subarray(0, alignedLength);
  const decrypted = decryptAes128Ecb(encryptedTail, key);

  const candidates = [{ mode: "distdoc+aes-ecb", bytes: decrypted }];
  const rawInflated = safeInflate(decrypted, true);
  if (rawInflated) {
    candidates.push({ mode: "distdoc+aes-ecb+inflateRaw(-15)", bytes: rawInflated });
  }
  const zlibInflated = safeInflate(decrypted, false);
  if (zlibInflated) {
    candidates.push({ mode: "distdoc+aes-ecb+inflate(zlib)", bytes: zlibInflated });
  }

  if (!compressedFlag) {
    return candidates;
  }
  return candidates;
}

function readRecordAt(bytes, offset) {
  if (!bytes || offset < 0 || offset + 4 > bytes.length) {
    return null;
  }
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const header = dv.getUint32(offset, true);
  const tag = header & 0x3ff;
  const level = (header >>> 10) & 0x3ff;
  let size = (header >>> 20) & 0xfff;
  let headerSize = 4;
  if (size === 0xfff) {
    if (offset + 8 > bytes.length) {
      return null;
    }
    size = dv.getUint32(offset + 4, true);
    headerSize = 8;
  }
  const payloadOffset = offset + headerSize;
  const nextOffset = payloadOffset + size;
  if (nextOffset > bytes.length) {
    return null;
  }
  return {
    offset,
    tag,
    level,
    size,
    headerSize,
    payloadOffset,
    nextOffset,
  };
}

function decodeDistributeDocHeadToKey(payload256) {
  if (!payload256 || payload256.length !== 256) {
    return null;
  }

  const decoded = new Uint8Array(payload256);
  const dv = new DataView(decoded.buffer, decoded.byteOffset, decoded.byteLength);
  const seed = dv.getUint32(0, true);
  const random = new MsvcRandom(seed);

  let n = 0;
  let keyByte = 0;
  for (let i = 0; i < 256; i += 1) {
    if (n === 0) {
      keyByte = random.rand() & 0xff;
      n = (random.rand() & 0x0f) + 1;
    }
    if (i >= 4) {
      decoded[i] ^= keyByte;
    }
    n -= 1;
  }

  const sha1Offset = 4 + (seed & 0x0f);
  return decoded.subarray(sha1Offset, sha1Offset + 16);
}

class MsvcRandom {
  constructor(seed) {
    this.seed = seed >>> 0;
  }

  rand() {
    this.seed = (this.seed * 214013 + 2531011) >>> 0;
    return (this.seed >>> 16) & 0x7fff;
  }
}

function decryptAes128Ecb(cipherBytes, keyBytes) {
  const decrypted = CryptoJS.AES.decrypt(
    { ciphertext: toCryptoWordArray(cipherBytes) },
    toCryptoWordArray(keyBytes),
    {
      mode: CryptoJS.mode.ECB,
      padding: CryptoJS.pad.NoPadding,
    }
  );
  return cryptoWordArrayToUint8Array(decrypted);
}

function toCryptoWordArray(bytes) {
  const words = [];
  for (let i = 0; i < bytes.length; i += 4) {
    words.push(
      ((bytes[i] || 0) << 24) |
        ((bytes[i + 1] || 0) << 16) |
        ((bytes[i + 2] || 0) << 8) |
        (bytes[i + 3] || 0)
    );
  }
  return CryptoJS.lib.WordArray.create(words, bytes.length);
}

function cryptoWordArrayToUint8Array(wordArray) {
  const out = new Uint8Array(wordArray.sigBytes);
  for (let i = 0; i < wordArray.sigBytes; i += 1) {
    out[i] = (wordArray.words[i >>> 2] >>> (24 - (i % 4) * 8)) & 0xff;
  }
  return out;
}

function analyzeLegacy3Stream(bytes) {
  const blocks = parseLegacy3TextBlocks(bytes, 320);
  const records = blocks.map((block, index) => ({
    index,
    offset: block.offset,
    headerSize: 0,
    tag: block.kind === "utf16" ? 897 : 896,
    level: 0,
    size: block.length,
    payloadOffset: block.offset,
    legacyText: block.text,
    legacyKind: block.kind,
  }));

  return {
    mode: "legacy3-heuristic",
    decoded: bytes,
    parse: { records, consumed: bytes.length, complete: false },
    records,
    paragraphs: [],
    recordToParagraph: new Map(),
  };
}

function parseLegacy3TextBlocks(bytes, maxBlocks = 240) {
  const blocks = [];

  const addBlock = (offset, length, kind, text) => {
    if (!Number.isFinite(offset) || !Number.isFinite(length) || length <= 0) {
      return;
    }
    blocks.push({ offset, length, kind, text: text || "" });
  };

  // ASCII scan (legacy headers and embedded labels)
  let startAscii = -1;
  for (let i = 0; i < bytes.length; i += 1) {
    const byte = bytes[i];
    const printable = byte === 0x09 || byte === 0x0a || byte === 0x0d || (byte >= 0x20 && byte <= 0x7e);
    if (printable) {
      if (startAscii < 0) {
        startAscii = i;
      }
      continue;
    }
    if (startAscii >= 0) {
      const len = i - startAscii;
      if (len >= 12) {
        const text = bytesToAscii(bytes.subarray(startAscii, i)).replace(/\s+/g, " ").trim();
        if (text.length >= 4) {
          addBlock(startAscii, len, "ascii", text);
        }
      }
      startAscii = -1;
    }
  }
  if (startAscii >= 0) {
    const len = bytes.length - startAscii;
    if (len >= 12) {
      const text = bytesToAscii(bytes.subarray(startAscii, bytes.length)).replace(/\s+/g, " ").trim();
      if (text.length >= 4) {
        addBlock(startAscii, len, "ascii", text);
      }
    }
  }

  // UTF-16LE scan (often used by legacy HWP text payloads)
  let i = 0;
  while (i + 3 < bytes.length) {
    const lo = bytes[i];
    const hi = bytes[i + 1];
    const likelyAsciiUtf16 = hi === 0 && lo >= 0x20 && lo <= 0x7e;
    const likelyHangulUtf16 = hi >= 0xac && hi <= 0xd7;
    if (!likelyAsciiUtf16 && !likelyHangulUtf16) {
      i += 1;
      continue;
    }

    const start = i;
    let end = i;
    while (end + 1 < bytes.length) {
      const x = bytes[end];
      const y = bytes[end + 1];
      const valid =
        y === 0 ||
        (y >= 0xac && y <= 0xd7) ||
        (x === 0x09 && y === 0x00) ||
        (x === 0x0a && y === 0x00) ||
        (x === 0x0d && y === 0x00);
      if (!valid) {
        break;
      }
      end += 2;
    }

    const len = end - start;
    if (len >= 16) {
      const text = UTF16LE_DECODER.decode(bytes.subarray(start, end)).replace(/\0/g, "").replace(/\s+/g, " ").trim();
      if (text.length >= 4 && /[A-Za-z0-9\u3131-\u318E\uAC00-\uD7A3]/.test(text)) {
        addBlock(start, len, "utf16", text);
      }
    }
    i = end > i ? end : i + 1;
  }

  blocks.sort((a, b) => a.offset - b.offset || b.length - a.length);
  const deduped = [];
  for (const block of blocks) {
    const overlaps = deduped.find(
      (item) =>
        block.offset >= item.offset &&
        block.offset + block.length <= item.offset + item.length &&
        block.kind === item.kind &&
        block.text === item.text
    );
    if (!overlaps) {
      deduped.push(block);
    }
    if (deduped.length >= maxBlocks) {
      break;
    }
  }
  return deduped;
}

function parseRecords(bytes, maxRecords = 12000) {
  const records = [];
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  let offset = 0;
  while (offset + 4 <= bytes.length && records.length < maxRecords) {
    const header = dv.getUint32(offset, true);
    const tag = header & 0x3ff;
    const level = (header >>> 10) & 0x3ff;
    let size = (header >>> 20) & 0xfff;
    let headerSize = 4;

    if (size === 0xfff) {
      if (offset + 8 > bytes.length) {
        break;
      }
      size = dv.getUint32(offset + 4, true);
      headerSize = 8;
    }

    const payloadOffset = offset + headerSize;
    const nextOffset = payloadOffset + size;
    if (nextOffset > bytes.length) {
      break;
    }

    records.push({
      index: records.length,
      offset,
      headerSize,
      tag,
      level,
      size,
      payloadOffset,
    });

    offset = nextOffset;
  }

  return {
    records,
    consumed: offset,
    complete: offset === bytes.length,
  };
}

function parseDocInfo(rawBytes, compressedFlag) {
  const decoded = decodeDocInfoStream(rawBytes, compressedFlag);
  const parsedRecords = parseRecords(decoded.bytes);
  const parsed = parseDocInfoRecords(decoded.bytes, parsedRecords.records);

  return {
    decodeMode: decoded.mode,
    decodedSize: decoded.bytes.length,
    records: parsedRecords.records,
    idMappings: parsed.idMappings,
    faceNameOffsets: parsed.faceNameOffsets,
    faceNames: parsed.faceNames,
    binDataRecords: parsed.binDataRecords,
    binDataById: parsed.binDataById,
    tabDefs: parsed.tabDefs,
    numberings: parsed.numberings,
    bullets: parsed.bullets,
    memoShapes: parsed.memoShapes,
    forbiddenChars: parsed.forbiddenChars,
    trackChanges: parsed.trackChanges,
    trackChangeAuthors: parsed.trackChangeAuthors,
    charShapes: parsed.charShapes,
    paraShapes: parsed.paraShapes,
    styles: parsed.styles,
    tabDefById: parsed.tabDefById,
    numberingById: parsed.numberingById,
    bulletById: parsed.bulletById,
    memoShapeById: parsed.memoShapeById,
    forbiddenCharById: parsed.forbiddenCharById,
    trackChangeById: parsed.trackChangeById,
    trackChangeAuthorById: parsed.trackChangeAuthorById,
    charShapeById: parsed.charShapeById,
    paraShapeById: parsed.paraShapeById,
    styleById: parsed.styleById,
  };
}

function buildBinDataCatalog(entries, docInfo) {
  const records = docInfo?.binDataRecords ?? [];
  const recordById = new Map();
  for (const record of records) {
    if (record.id != null) {
      recordById.set(record.id, record);
    }
  }

  const list = [];
  const byId = new Map();
  const binEntries = entries.filter(
    (entry) => entry.type === 2 && /\/BinData\/BIN[0-9A-Fa-f]{4}(?:\.[^/]+)?$/.test(entry.fullPath)
  );

  for (const entry of binEntries) {
    const parsed = parseBinDataEntryInfo(entry.fullPath);
    if (!parsed) {
      continue;
    }
    const record = recordById.get(parsed.id) ?? null;
    const extension = normalizeExtension(record?.extension || parsed.extension || "");
    const decoded = decodeBinDataPayload(entry.content, extension);
    const item = {
      id: parsed.id,
      extension,
      streamPath: entry.fullPath,
      storageType: record?.storageType ?? 1,
      compression: record?.compression ?? null,
      flags: record?.flags ?? null,
      rawSize: entry.content.length,
      decodedSize: decoded.bytes.length,
      decodeStrategy: decoded.decodeStrategy,
      format: decoded.format || "",
      mime: decoded.mime || mimeFromExtension(extension) || "application/octet-stream",
      bytes: decoded.bytes,
      objectUrl: null,
      isRenderableImage: isRenderableImageMime(decoded.mime),
    };
    list.push(item);
    byId.set(item.id, item);
  }

  for (const record of records) {
    if (record.id == null || byId.has(record.id)) {
      continue;
    }
    const extension = normalizeExtension(record.extension || "");
    const mime = mimeFromExtension(extension) || "application/octet-stream";
    const item = {
      id: record.id,
      extension,
      streamPath: "",
      storageType: record.storageType,
      compression: record.compression,
      flags: record.flags,
      rawSize: 0,
      decodedSize: 0,
      decodeStrategy: "no-stream",
      format: extension,
      mime,
      bytes: new Uint8Array(),
      objectUrl: null,
      isRenderableImage: isRenderableImageMime(mime),
    };
    list.push(item);
    byId.set(item.id, item);
  }

  list.sort((a, b) => a.id - b.id);
  return { list, byId };
}

function parseBinDataEntryInfo(fullPath) {
  const match = fullPath.match(/\/BinData\/BIN([0-9A-Fa-f]{4})(?:\.([^/]+))?$/);
  if (!match) {
    return null;
  }
  return {
    id: Number.parseInt(match[1], 16),
    extension: normalizeExtension(match[2] || ""),
  };
}

function decodeBinDataPayload(rawBytes, extensionHint = "") {
  const raw = toUint8Array(rawBytes);
  if (!raw.length) {
    return {
      bytes: raw,
      format: normalizeExtension(extensionHint),
      mime: mimeFromExtension(extensionHint) || "application/octet-stream",
      decodeStrategy: "empty",
    };
  }

  const attempts = [];
  const seen = new Set();
  const addAttempt = (label, bytes) => {
    const key = `${label}:${bytes.length}`;
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    attempts.push({ label, bytes });
  };

  addAttempt("raw", raw);
  for (let skip = 1; skip <= 8 && skip < raw.length; skip += 1) {
    addAttempt(`raw+${skip}`, raw.subarray(skip));
  }

  const inflatedRaw = safeInflate(raw, true);
  if (inflatedRaw) {
    addAttempt("inflateRaw", inflatedRaw);
  }
  const inflatedZlib = safeInflate(raw, false);
  if (inflatedZlib) {
    addAttempt("inflate", inflatedZlib);
  }

  for (let skip = 1; skip <= 8 && skip < raw.length; skip += 1) {
    const subset = raw.subarray(skip);
    const subInflatedRaw = safeInflate(subset, true);
    if (subInflatedRaw) {
      addAttempt(`inflateRaw+${skip}`, subInflatedRaw);
    }
    const subInflatedZlib = safeInflate(subset, false);
    if (subInflatedZlib) {
      addAttempt(`inflate+${skip}`, subInflatedZlib);
    }
  }

  for (const attempt of attempts) {
    const detected = detectImageSignature(attempt.bytes);
    if (!detected) {
      continue;
    }
    const normalized = detected.offset > 0 ? attempt.bytes.subarray(detected.offset) : attempt.bytes;
    return {
      bytes: normalized,
      format: detected.format,
      mime: detected.mime,
      decodeStrategy: attempt.label,
    };
  }

  return {
    bytes: raw,
    format: normalizeExtension(extensionHint),
    mime: mimeFromExtension(extensionHint) || "application/octet-stream",
    decodeStrategy: "raw-unknown",
  };
}

function detectImageSignature(bytes, scanLimit = 40) {
  if (!bytes || bytes.length < 2) {
    return null;
  }

  const signatures = [
    { format: "png", mime: "image/png", bytes: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a] },
    { format: "jpg", mime: "image/jpeg", bytes: [0xff, 0xd8, 0xff] },
    { format: "gif", mime: "image/gif", bytes: [0x47, 0x49, 0x46, 0x38] },
    { format: "bmp", mime: "image/bmp", bytes: [0x42, 0x4d] },
    { format: "webp", mime: "image/webp", bytes: [0x52, 0x49, 0x46, 0x46] },
    { format: "wmf", mime: "image/wmf", bytes: [0xd7, 0xcd, 0xc6, 0x9a] },
  ];

  const maxOffset = Math.max(0, Math.min(scanLimit, bytes.length - 2));
  for (let offset = 0; offset <= maxOffset; offset += 1) {
    for (const signature of signatures) {
      if (offset + signature.bytes.length > bytes.length) {
        continue;
      }
      let matched = true;
      for (let i = 0; i < signature.bytes.length; i += 1) {
        if (bytes[offset + i] !== signature.bytes[i]) {
          matched = false;
          break;
        }
      }
      if (!matched) {
        continue;
      }
      if (signature.format === "webp") {
        if (offset + 12 <= bytes.length && bytes[offset + 8] === 0x57 && bytes[offset + 9] === 0x45) {
          return { ...signature, offset };
        }
        continue;
      }
      return { ...signature, offset };
    }
  }
  return null;
}

function normalizeExtension(ext) {
  if (!ext) {
    return "";
  }
  return String(ext).replace(/^\./, "").trim().toLowerCase();
}

function mimeFromExtension(extension) {
  const ext = normalizeExtension(extension);
  return IMAGE_MIME_BY_EXT[ext] || "";
}

function isRenderableImageMime(mime) {
  if (!mime) {
    return false;
  }
  return (
    mime === "image/png" ||
    mime === "image/jpeg" ||
    mime === "image/gif" ||
    mime === "image/bmp" ||
    mime === "image/webp" ||
    mime === "image/svg+xml"
  );
}

function resolveBinDataImageRef(binDataId, doc = state.doc) {
  const ref = resolveBinDataRef(binDataId, doc);
  if (!ref || !ref.mime?.startsWith("image/")) {
    return null;
  }
  return ref;
}

function resolveBinDataRef(binDataId, doc = state.doc) {
  if (!doc?.binDataById || binDataId == null) {
    return null;
  }
  const item = doc.binDataById.get(binDataId) ?? null;
  if (!item || !item.bytes?.length) {
    return null;
  }
  const url = ensureBinDataObjectUrl(item);
  if (!url) {
    return null;
  }
  return {
    id: item.id,
    url,
    mime: item.mime,
    format: item.format || item.extension || "",
    decodeStrategy: item.decodeStrategy,
    rawSize: item.rawSize,
    decodedSize: item.decodedSize,
  };
}

function ensureBinDataObjectUrl(item) {
  if (item.objectUrl) {
    return item.objectUrl;
  }
  if (!item.bytes?.length) {
    return null;
  }
  try {
    item.objectUrl = URL.createObjectURL(new Blob([item.bytes], { type: item.mime || "application/octet-stream" }));
    return item.objectUrl;
  } catch {
    return null;
  }
}

function releaseDocResources(doc) {
  if (!doc?.binDataList?.length) {
    return;
  }
  for (const item of doc.binDataList) {
    if (!item?.objectUrl) {
      continue;
    }
    try {
      URL.revokeObjectURL(item.objectUrl);
    } catch {
      // ignore
    }
    item.objectUrl = null;
  }
}

function decodeDocInfoStream(rawBytes, compressedFlag) {
  const candidates = [{ mode: "raw", bytes: rawBytes }];
  if (compressedFlag) {
    const rawInflated = safeInflate(rawBytes, true);
    if (rawInflated) {
      candidates.push({ mode: "inflateRaw(-15)", bytes: rawInflated });
    }
    const zlibInflated = safeInflate(rawBytes, false);
    if (zlibInflated) {
      candidates.push({ mode: "inflate(zlib)", bytes: zlibInflated });
    }
  }

  let best = null;
  for (const candidate of candidates) {
    const parsed = parseRecords(candidate.bytes);
    const score = parsed.records.length + (parsed.complete ? 10 : 0);
    if (!best || score > best.score || (score === best.score && candidate.bytes.length > best.bytes.length)) {
      best = { ...candidate, score };
    }
  }

  return best ?? { mode: "raw", bytes: rawBytes };
}

function parseDocInfoRecords(decoded, records) {
  let idMappings = null;
  const faceNames = [];
  const binDataRecords = [];
  const tabDefs = [];
  const numberings = [];
  const bullets = [];
  const memoShapes = [];
  const forbiddenChars = [];
  const trackChanges = [];
  const trackChangeAuthors = [];
  const charShapes = [];
  const paraShapes = [];
  const styles = [];

  for (const record of records) {
    const payload = getRecordPayload(decoded, record);
    if (record.tag === 17) {
      idMappings = parseIdMappings(payload);
      continue;
    }
    if (record.tag === 19) {
      faceNames.push(parseFaceName(payload, faceNames.length));
      continue;
    }
    if (record.tag === 18) {
      binDataRecords.push(parseBinDataRecord(payload, binDataRecords.length));
      continue;
    }
    if (record.tag === 22) {
      tabDefs.push(parseTabDef(payload, tabDefs.length));
      continue;
    }
    if (record.tag === 23) {
      numberings.push(parseNumbering(payload, numberings.length));
      continue;
    }
    if (record.tag === 24) {
      bullets.push(parseBullet(payload, bullets.length));
      continue;
    }
    if (record.tag === 32 || record.tag === 96) {
      trackChanges.push(parseTrackChangeRecord(payload, trackChanges.length, record.tag));
      continue;
    }
    if (record.tag === 92) {
      memoShapes.push(parseMemoShapeRecord(payload, memoShapes.length));
      continue;
    }
    if (record.tag === 94) {
      forbiddenChars.push(parseForbiddenCharRecord(payload, forbiddenChars.length));
      continue;
    }
    if (record.tag === 97) {
      trackChangeAuthors.push(parseTrackChangeAuthorRecord(payload, trackChangeAuthors.length));
      continue;
    }
    if (record.tag === 21) {
      charShapes.push(parseDocCharShape(payload, charShapes.length));
      continue;
    }
    if (record.tag === 25) {
      paraShapes.push(parseDocParaShape(payload, paraShapes.length));
      continue;
    }
    if (record.tag === 26) {
      styles.push(parseStyle(payload, styles.length));
    }
  }

  const faceNameOffsets = buildFaceNameOffsets(idMappings);
  for (const shape of charShapes) {
    const fontNames = {};
    SCRIPT_LANGS.forEach((lang, langIndex) => {
      const faceId = shape.fontFaceIds[langIndex] ?? null;
      fontNames[lang] = resolveFaceName(faceNames, faceNameOffsets, langIndex, faceId);
    });
    shape.fontNames = fontNames;
    shape.primaryFont =
      fontNames.ko ||
      fontNames.en ||
      fontNames.hanja ||
      fontNames.jp ||
      fontNames.other ||
      fontNames.symbol ||
      fontNames.user ||
      null;
  }

  const binDataById = new Map();
  for (const item of binDataRecords) {
    if (item.id != null) {
      binDataById.set(item.id, item);
    }
  }

  const tabDefById = new Map();
  const numberingById = new Map();
  const bulletById = new Map();
  const memoShapeById = new Map();
  const forbiddenCharById = new Map();
  const trackChangeById = new Map();
  const trackChangeAuthorById = new Map();
  const charShapeById = new Map();
  const paraShapeById = new Map();
  const styleById = new Map();
  for (const item of tabDefs) {
    tabDefById.set(item.id, item);
  }
  for (const item of numberings) {
    numberingById.set(item.id, item);
  }
  for (const item of bullets) {
    bulletById.set(item.id, item);
  }
  for (const item of memoShapes) {
    memoShapeById.set(item.id, item);
  }
  for (const item of forbiddenChars) {
    forbiddenCharById.set(item.id, item);
  }
  for (const item of trackChanges) {
    trackChangeById.set(item.id, item);
  }
  for (const item of trackChangeAuthors) {
    trackChangeAuthorById.set(item.id, item);
  }
  for (const item of charShapes) {
    charShapeById.set(item.id, item);
  }
  for (const item of paraShapes) {
    paraShapeById.set(item.id, item);
  }
  for (const item of styles) {
    styleById.set(item.id, item);
  }

  return {
    idMappings,
    faceNameOffsets,
    faceNames,
    binDataRecords,
    binDataById,
    tabDefs,
    numberings,
    bullets,
    memoShapes,
    forbiddenChars,
    trackChanges,
    trackChangeAuthors,
    charShapes,
    paraShapes,
    styles,
    tabDefById,
    numberingById,
    bulletById,
    memoShapeById,
    forbiddenCharById,
    trackChangeById,
    trackChangeAuthorById,
    charShapeById,
    paraShapeById,
    styleById,
  };
}

function parseIdMappings(payload) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const counts = [];
  for (let offset = 0; offset + 3 < payload.length; offset += 4) {
    counts.push(dv.getUint32(offset, true));
  }

  const names = [
    "binData",
    "faceNameKo",
    "faceNameEn",
    "faceNameHanja",
    "faceNameJp",
    "faceNameOther",
    "faceNameSymbol",
    "faceNameUser",
    "borderFill",
    "charShape",
    "tabDef",
    "numbering",
    "bullet",
    "paraShape",
    "style",
    "memoShape",
    "trackChange",
    "trackChangeUser",
  ];

  const byName = {};
  counts.forEach((count, index) => {
    byName[names[index] ?? `field${index}`] = count;
  });

  return {
    counts,
    byName,
    faceNameCounts: counts.slice(1, 8),
  };
}

function parseFaceName(payload, id) {
  const attr = payload.length ? payload[0] : 0;
  const candidates = findLengthPrefixedUtf16Strings(payload, 1);
  const primaryName = candidates[0]?.text ?? "";
  const alternativeName = candidates.length > 1 ? candidates[candidates.length - 1].text : null;

  return {
    id,
    attr,
    primaryName,
    alternativeName,
    allNames: candidates.map((item) => item.text),
  };
}

function parseBinDataRecord(payload, index) {
  if (!payload || payload.length < 4) {
    return {
      id: index + 1,
      storageType: null,
      storageTypeName: null,
      compression: null,
      compressionName: null,
      extension: "",
      absolutePath: "",
      relativePath: "",
      flags: null,
    };
  }

  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const flags = dv.getUint16(0, true);
  const storageType = flags & 0x0f;
  const compression = (flags >>> 4) & 0x3;
  const accessState = (flags >>> 8) & 0x3;
  let id = index + 1;
  let extension = "";
  let absolutePath = "";
  let relativePath = "";
  let cursor = 2;

  if (storageType === 0) {
    const abs = readUtf16LengthPrefixed(payload, cursor);
    if (abs) {
      absolutePath = abs.text;
      cursor = abs.nextOffset;
    }
    const rel = readUtf16LengthPrefixed(payload, cursor);
    if (rel) {
      relativePath = rel.text;
      cursor = rel.nextOffset;
    }
  } else {
    if (cursor + 1 < payload.length) {
      id = dv.getUint16(cursor, true);
      cursor += 2;
    }
    const ext = readUtf16LengthPrefixed(payload, cursor);
    if (ext) {
      extension = ext.text;
      cursor = ext.nextOffset;
    }
  }

  return {
    id,
    storageType,
    storageTypeName: BIN_STORAGE_TYPE_NAMES[storageType] ?? `storage${storageType}`,
    compression,
    compressionName: BIN_COMPRESSION_NAMES[compression] ?? `compression${compression}`,
    accessState,
    extension: extension.trim(),
    absolutePath,
    relativePath,
    flags,
    payloadSize: payload.length,
  };
}

function parseStyle(payload, id) {
  let offset = 0;
  const localNameInfo = readUtf16LengthPrefixed(payload, offset);
  let localName = "";
  if (localNameInfo) {
    localName = localNameInfo.text;
    offset = localNameInfo.nextOffset;
  }

  const englishNameInfo = readUtf16LengthPrefixed(payload, offset);
  let englishName = "";
  if (englishNameInfo) {
    englishName = englishNameInfo.text;
    offset = englishNameInfo.nextOffset;
  }

  let styleType = null;
  let nextStyleId = null;
  let languageId = null;
  let paraShapeId = null;
  let charShapeId = null;
  let lockForm = null;

  if (offset + 10 <= payload.length) {
    const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
    styleType = payload[offset];
    nextStyleId = payload[offset + 1];
    languageId = dv.getUint16(offset + 2, true);
    paraShapeId = dv.getUint16(offset + 4, true);
    charShapeId = dv.getUint16(offset + 6, true);
    lockForm = dv.getUint16(offset + 8, true);
  }

  if (!localName && !englishName) {
    const candidates = findLengthPrefixedUtf16Strings(payload, 0);
    localName = candidates[0]?.text ?? "";
    englishName = candidates[1]?.text ?? "";
  }

  return {
    id,
    localName,
    englishName,
    styleType,
    nextStyleId,
    languageId,
    paraShapeId,
    charShapeId,
    lockForm,
  };
}

function parseTabDef(payload, id) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const property = payload.length >= 4 ? dv.getUint32(0, true) : 0;
  const countByHeader = payload.length >= 8 ? dv.getUint32(4, true) : 0;
  const tabStops = [];
  let cursor = 8;

  const maxCountByLength = Math.floor(Math.max(0, payload.length - cursor) / 8);
  const count = clamp(countByHeader, 0, maxCountByLength);
  for (let i = 0; i < count; i += 1) {
    const position = dv.getInt32(cursor, true);
    const align = payload[cursor + 4] ?? 0;
    const leader = payload[cursor + 5] ?? 0;
    const reserved = dv.getUint16(cursor + 6, true);
    tabStops.push({
      position,
      positionPx: hwpToPx(position),
      align,
      alignName: TAB_ALIGN_NAMES[align] ?? `align${align}`,
      leader,
      leaderName: TAB_LEADER_NAMES[leader] ?? `leader${leader}`,
      reserved,
    });
    cursor += 8;
  }

  return {
    id,
    property,
    count,
    tabStops,
  };
}

function parseNumbering(payload, id) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const levels = [];
  let cursor = payload.length >= 6 ? 6 : 0;

  for (let level = 0; level < 7; level += 1) {
    if (cursor + 8 > payload.length) {
      break;
    }

    const shapeRef = dv.getUint16(cursor, true);
    const widthAdjust = dv.getInt16(cursor + 2, true);
    const textOffset = dv.getInt16(cursor + 4, true);
    const formatLength = dv.getUint16(cursor + 6, true);
    const textStart = cursor + 8;
    const textEnd = textStart + formatLength * 2;
    if (formatLength > 120 || textEnd > payload.length) {
      break;
    }

    const format = UTF16LE_DECODER.decode(payload.subarray(textStart, textEnd)).replace(/\0/g, "");
    const tailA = textEnd + 1 < payload.length ? dv.getUint16(textEnd, true) : null;
    const tailB = textEnd + 3 < payload.length ? dv.getUint16(textEnd + 2, true) : null;
    const tailC = textEnd + 5 < payload.length ? dv.getUint16(textEnd + 4, true) : null;

    levels.push({
      level: level + 1,
      shapeRef,
      widthAdjust,
      textOffset,
      formatLength,
      format,
      charShapeId: tailA,
      unknownA: tailB,
      unknownB: tailC,
    });

    cursor = textEnd + 6;
  }

  if (!levels.length) {
    const candidates = findLengthPrefixedUtf16Strings(payload, 0)
      .map((item) => item.text)
      .slice(0, 7);
    candidates.forEach((format, index) => {
      levels.push({
        level: index + 1,
        shapeRef: null,
        widthAdjust: null,
        textOffset: null,
        formatLength: format.length,
        format,
        charShapeId: null,
        unknownA: null,
        unknownB: null,
      });
    });
  }

  return {
    id,
    property: payload.length >= 4 ? dv.getUint32(0, true) : null,
    levels,
  };
}

function parseBullet(payload, id) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const property = payload.length >= 4 ? dv.getUint32(0, true) : 0;
  const shapeRef = payload.length >= 8 ? dv.getUint16(6, true) : null;
  const charCode = payload.length >= 14 ? dv.getUint16(12, true) : 0;
  const marker = charCode ? String.fromCodePoint(charCode) : "•";

  return {
    id,
    property,
    shapeRef,
    charCode,
    marker,
  };
}

function parseMemoShapeRecord(payload, id) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const property = payload.length >= 4 ? dv.getUint32(0, true) : null;
  const width = payload.length >= 8 ? dv.getInt32(4, true) : null;
  const margin = payload.length >= 16 ? dv.getInt32(12, true) : null;
  const strings = findLengthPrefixedUtf16Strings(payload, 0).map((item) => item.text).slice(0, 3);
  return {
    id,
    property,
    width,
    margin,
    strings,
    payloadSize: payload.length,
  };
}

function parseForbiddenCharRecord(payload, id) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const countByHeader = payload.length >= 4 ? dv.getUint32(0, true) : null;
  const chars = [];
  for (let offset = 0; offset + 1 < payload.length && chars.length < 64; offset += 2) {
    const code = dv.getUint16(offset, true);
    if (code === 0 || code === 0xffff) {
      continue;
    }
    if (code >= 0xd800 && code <= 0xdfff) {
      continue;
    }
    const ch = String.fromCodePoint(code);
    if (!/\S/.test(ch)) {
      continue;
    }
    chars.push({
      code,
      char: ch,
    });
  }

  return {
    id,
    countByHeader,
    sample: chars.slice(0, 24),
    payloadSize: payload.length,
  };
}

function parseTrackChangeRecord(payload, id, tag) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const flags = payload.length >= 4 ? dv.getUint32(0, true) : null;
  const authorId = payload.length >= 8 ? dv.getUint32(4, true) : null;
  const timestampLow = payload.length >= 12 ? dv.getUint32(8, true) : null;
  const strings = findLengthPrefixedUtf16Strings(payload, 0).map((item) => item.text).slice(0, 4);
  const summary = strings[0] || extractReadableAscii(payload, 120) || payloadWordPreview(payload, 6);

  return {
    id,
    tag,
    flags,
    authorId,
    timestampLow,
    summary,
    strings,
    payloadSize: payload.length,
  };
}

function parseTrackChangeAuthorRecord(payload, id) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const authorIndex = payload.length >= 4 ? dv.getUint32(0, true) : null;
  const strings = findLengthPrefixedUtf16Strings(payload, 0).map((item) => item.text).slice(0, 6);
  const name = strings[0] || extractReadableAscii(payload, 80) || "";
  return {
    id,
    authorIndex,
    name,
    strings,
    payloadSize: payload.length,
  };
}

function parseDocCharShape(payload, id) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const fontFaceIds = [];
  for (let i = 0; i < 7; i += 1) {
    const offset = i * 2;
    if (offset + 1 >= payload.length) {
      fontFaceIds.push(null);
    } else {
      fontFaceIds.push(dv.getUint16(offset, true));
    }
  }

  const ratio = [];
  const spacing = [];
  const relativeSize = [];
  const positionOffset = [];
  for (let i = 0; i < 7; i += 1) {
    ratio.push(payload.length > 14 + i ? payload[14 + i] : null);
    spacing.push(payload.length > 21 + i ? toSignedByte(payload[21 + i]) : null);
    relativeSize.push(payload.length > 28 + i ? payload[28 + i] : null);
    positionOffset.push(payload.length > 35 + i ? toSignedByte(payload[35 + i]) : null);
  }

  const baseSize = payload.length >= 46 ? dv.getUint32(42, true) : null;
  const property = payload.length >= 50 ? dv.getUint32(46, true) : null;
  const shadowOffsetX = payload.length >= 51 ? toSignedByte(payload[50]) : null;
  const shadowOffsetY = payload.length >= 52 ? toSignedByte(payload[51]) : null;
  const textColor = payload.length >= 56 ? parseColorRef(dv.getUint32(52, true)) : null;
  const underlineColor = payload.length >= 60 ? parseColorRef(dv.getUint32(56, true)) : null;
  const shadeColor = payload.length >= 64 ? parseColorRef(dv.getUint32(60, true)) : null;
  const shadowColor = payload.length >= 68 ? parseColorRef(dv.getUint32(64, true)) : null;
  const borderFillId = payload.length >= 70 ? dv.getUint16(68, true) : null;
  const strikeoutColor = payload.length >= 74 ? parseColorRef(dv.getUint32(70, true)) : null;
  const attributeBits = parseCharShapeAttributes(property);

  return {
    id,
    fontFaceIds,
    ratio,
    spacing,
    relativeSize,
    positionOffset,
    baseSize,
    property,
    attributeBits,
    shadowOffsetX,
    shadowOffsetY,
    textColor,
    underlineColor,
    shadeColor,
    shadowColor,
    borderFillId,
    strikeoutColor,
    fontNames: {},
    primaryFont: null,
  };
}

function parseDocParaShape(payload, id) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);

  const property1 = payload.length >= 4 ? dv.getUint32(0, true) : null;
  const leftMargin = payload.length >= 8 ? dv.getInt32(4, true) : null;
  const rightMargin = payload.length >= 12 ? dv.getInt32(8, true) : null;
  const indent = payload.length >= 16 ? dv.getInt32(12, true) : null;
  const spacingBefore = payload.length >= 20 ? dv.getInt32(16, true) : null;
  const spacingAfter = payload.length >= 24 ? dv.getInt32(20, true) : null;
  const lineSpacing = payload.length >= 28 ? dv.getInt32(24, true) : null;
  const tabDefId = payload.length >= 30 ? dv.getUint16(28, true) : null;
  const numberingId = payload.length >= 32 ? dv.getUint16(30, true) : null;
  const borderFillId = payload.length >= 34 ? dv.getUint16(32, true) : null;
  const borderLeftSpacing = payload.length >= 36 ? dv.getInt16(34, true) : null;
  const borderRightSpacing = payload.length >= 38 ? dv.getInt16(36, true) : null;
  const borderTopSpacing = payload.length >= 40 ? dv.getInt16(38, true) : null;
  const borderBottomSpacing = payload.length >= 42 ? dv.getInt16(40, true) : null;
  const property2 = payload.length >= 46 ? dv.getUint32(42, true) : null;
  const property3 = payload.length >= 50 ? dv.getUint32(46, true) : null;
  const lineSpacingModern = payload.length >= 54 ? dv.getInt32(50, true) : null;
  const extension = payload.length >= 58 ? dv.getUint32(54, true) : null;
  const effectiveLineSpacing = lineSpacingModern != null && lineSpacingModern !== 0 ? lineSpacingModern : lineSpacing;
  const property1Bits = parseParaShapeProperty1(property1);
  const property2Bits = parseParaShapeProperty2(property2);
  const property3Bits = parseParaShapeProperty3(property3);

  return {
    id,
    property1,
    leftMargin,
    rightMargin,
    indent,
    spacingBefore,
    spacingAfter,
    lineSpacing: effectiveLineSpacing,
    lineSpacingLegacy: lineSpacing,
    lineSpacingModern,
    tabDefId,
    numberingId,
    borderFillId,
    borderLeftSpacing,
    borderRightSpacing,
    borderTopSpacing,
    borderBottomSpacing,
    property2,
    property3,
    extension,
    property1Bits,
    property2Bits,
    property3Bits,
  };
}

function parseColorRef(rawValue) {
  const value = rawValue >>> 0;
  if (value === 0xffffffff) {
    return {
      raw: value,
      auto: true,
      r: null,
      g: null,
      b: null,
      a: 255,
      hex: "AUTO",
    };
  }

  const r = value & 0xff;
  const g = (value >>> 8) & 0xff;
  const b = (value >>> 16) & 0xff;
  const a = (value >>> 24) & 0xff;

  return {
    raw: value,
    auto: false,
    r,
    g,
    b,
    a,
    hex: `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`,
  };
}

function bitsValue(value, start, length) {
  const mask = (1 << length) - 1;
  return (value >>> start) & mask;
}

function parseCharShapeAttributes(property) {
  if (property == null) {
    return null;
  }

  const underlineType = bitsValue(property, 2, 2);
  const underlineShape = bitsValue(property, 4, 4);
  const outlineType = bitsValue(property, 8, 3);
  const shadowType = bitsValue(property, 11, 2);
  const strikeoutType = bitsValue(property, 18, 3);
  const accentMark = bitsValue(property, 21, 4);
  const strikeoutShape = bitsValue(property, 26, 4);

  const result = {
    italic: Boolean(property & (1 << 0)),
    bold: Boolean(property & (1 << 1)),
    underlineType,
    underlineShape,
    outlineType,
    shadowType,
    embossed: Boolean(property & (1 << 13)),
    engraved: Boolean(property & (1 << 14)),
    superscript: Boolean(property & (1 << 15)),
    subscript: Boolean(property & (1 << 16)),
    strikeoutType,
    accentMark,
    useFontSpace: Boolean(property & (1 << 25)),
    strikeoutShape,
    kerning: Boolean(property & (1 << 30)),
  };

  const labels = [];
  if (result.italic) labels.push("italic");
  if (result.bold) labels.push("bold");
  if (result.underlineType !== 0) {
    labels.push(`underline:${UNDERLINE_TYPE_NAMES[result.underlineType] ?? result.underlineType}`);
  }
  if (result.outlineType !== 0) {
    labels.push(`outline:${OUTLINE_TYPE_NAMES[result.outlineType] ?? result.outlineType}`);
  }
  if (result.shadowType !== 0) {
    labels.push(`shadow:${SHADOW_TYPE_NAMES[result.shadowType] ?? result.shadowType}`);
  }
  if (result.embossed) labels.push("embossed");
  if (result.engraved) labels.push("engraved");
  if (result.superscript) labels.push("superscript");
  if (result.subscript) labels.push("subscript");
  if (result.strikeoutType !== 0) labels.push(`strike:${result.strikeoutType}`);
  if (result.accentMark !== 0) {
    labels.push(`accent:${ACCENT_MARK_NAMES[result.accentMark] ?? result.accentMark}`);
  }
  if (result.useFontSpace) labels.push("font-space");
  if (result.kerning) labels.push("kerning");

  result.labels = labels;
  return result;
}

function parseParaShapeProperty1(property) {
  if (property == null) {
    return null;
  }

  const result = {
    lineSpacingTypeLegacy: bitsValue(property, 0, 2),
    align: bitsValue(property, 2, 3),
    breakLatinWord: bitsValue(property, 5, 2),
    breakKoreanWord: bitsValue(property, 7, 1),
    useGrid: Boolean(property & (1 << 8)),
    minSpace: bitsValue(property, 9, 7),
    widowOrphanProtect: Boolean(property & (1 << 16)),
    keepWithNext: Boolean(property & (1 << 17)),
    paragraphProtect: Boolean(property & (1 << 18)),
    pageBreakBefore: Boolean(property & (1 << 19)),
    verticalAlign: bitsValue(property, 20, 2),
    fontLineHeight: Boolean(property & (1 << 22)),
    headingType: bitsValue(property, 23, 2),
    headingLevelRaw: bitsValue(property, 25, 3),
    connectBorder: Boolean(property & (1 << 28)),
    ignoreMargin: Boolean(property & (1 << 29)),
    tailShape: Boolean(property & (1 << 30)),
  };

  result.headingLevel = result.headingLevelRaw === 0 ? 0 : result.headingLevelRaw;
  result.labels = [];
  result.labels.push(`align:${PARA_ALIGN_NAMES[result.align] ?? result.align}`);
  if (result.useGrid) result.labels.push("grid");
  if (result.keepWithNext) result.labels.push("keep-next");
  if (result.pageBreakBefore) result.labels.push("page-break-before");
  if (result.headingType !== 0) {
    result.labels.push(`heading:${PARA_HEADING_TYPE_NAMES[result.headingType] ?? result.headingType}`);
  }
  if (result.headingLevel > 0) {
    result.labels.push(`level:${result.headingLevel}`);
  }

  return result;
}

function parseParaShapeProperty2(property) {
  if (property == null) {
    return null;
  }

  const result = {
    singleLineInput: bitsValue(property, 0, 2),
    autoSpaceHangulEnglish: Boolean(property & (1 << 4)),
    autoSpaceHangulNumber: Boolean(property & (1 << 5)),
  };

  result.labels = [];
  if (result.autoSpaceHangulEnglish) result.labels.push("auto-space-ko-en");
  if (result.autoSpaceHangulNumber) result.labels.push("auto-space-ko-num");
  return result;
}

function parseParaShapeProperty3(property) {
  if (property == null) {
    return null;
  }

  const result = {
    lineSpacingType: bitsValue(property, 0, 5),
  };
  result.lineSpacingTypeName = PARA_LINE_SPACING_TYPE_NAMES[result.lineSpacingType] ?? `type-${result.lineSpacingType}`;
  return result;
}

function buildFaceNameOffsets(idMappings) {
  const counts = idMappings?.faceNameCounts ?? [];
  const offsets = [];
  let start = 0;
  for (let i = 0; i < SCRIPT_LANGS.length; i += 1) {
    const count = counts[i] ?? 0;
    offsets.push({ start, count });
    start += count;
  }
  return offsets;
}

function resolveFaceName(faceNames, faceNameOffsets, langIndex, faceId) {
  if (faceId == null) {
    return null;
  }

  const mapping = faceNameOffsets[langIndex];
  if (!mapping) {
    return null;
  }

  if (faceId < 0 || faceId >= mapping.count) {
    return null;
  }

  const item = faceNames[mapping.start + faceId];
  if (!item) {
    return null;
  }

  return item.primaryName || item.alternativeName || null;
}

function readUtf16LengthPrefixed(payload, offset) {
  if (offset + 1 >= payload.length) {
    return null;
  }
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const charLength = dv.getUint16(offset, true);
  const byteLength = charLength * 2;
  const textStart = offset + 2;
  const textEnd = textStart + byteLength;
  if (textEnd > payload.length) {
    return null;
  }

  const text = UTF16LE_DECODER.decode(payload.subarray(textStart, textEnd)).replace(/\0/g, "");

  return {
    charLength,
    text,
    nextOffset: textEnd,
  };
}

function findLengthPrefixedUtf16Strings(payload, startOffset = 0) {
  const out = [];
  const seen = new Set();

  for (let offset = startOffset; offset + 1 < payload.length; offset += 1) {
    const info = readUtf16LengthPrefixed(payload, offset);
    if (!info || info.charLength <= 0 || info.charLength > 80) {
      continue;
    }
    if (!isPlausibleUtf16Text(info.text)) {
      continue;
    }
    if (seen.has(offset)) {
      continue;
    }
    seen.add(offset);
    out.push({ offset, text: info.text, charLength: info.charLength });
  }

  out.sort((a, b) => a.offset - b.offset);
  return out;
}

function isPlausibleUtf16Text(text) {
  if (!text) {
    return false;
  }
  if (!/[A-Za-z0-9\u3131-\u318E\uAC00-\uD7A3]/.test(text)) {
    return false;
  }

  for (const char of text) {
    const code = char.charCodeAt(0);
    if (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) {
      return false;
    }
  }
  return true;
}

function toSignedByte(value) {
  return value > 127 ? value - 256 : value;
}

function buildParagraphContext(records) {
  const paragraphs = [];
  const recordToParagraph = new Map();
  const paragraphLevels = records.filter((record) => record.tag === 66 && Number.isFinite(record.level)).map((record) => record.level);
  const baseParagraphLevel = paragraphLevels.length ? Math.min(...paragraphLevels) : 0;
  const directChildLevel = baseParagraphLevel + 1;
  const hasDirectChildRecords = records.some((record) => {
    if (!Number.isFinite(record.level)) {
      return false;
    }
    if (record.level !== directChildLevel) {
      return false;
    }
    return record.tag === 67 || record.tag === 68 || record.tag === 69 || record.tag === 71;
  });
  const contentRecordLevel = hasDirectChildRecords ? directChildLevel : baseParagraphLevel;
  let current = null;

  for (const record of records) {
    if (record.tag === 66 && record.level === baseParagraphLevel) {
      current = {
        index: paragraphs.length,
        headerRecord: record,
        textRecord: null,
        charShapeRecord: null,
        lineSegRecord: null,
        ctrlHeaderRecords: [],
        records: [],
      };
      paragraphs.push(current);
    }

    if (!current) {
      continue;
    }

    current.records.push(record);
    recordToParagraph.set(record.index, current.index);

    if (record.level !== contentRecordLevel) {
      continue;
    }

    if (record.tag === 67 && !current.textRecord) {
      current.textRecord = record;
    } else if (record.tag === 68 && !current.charShapeRecord) {
      current.charShapeRecord = record;
    } else if (record.tag === 69 && !current.lineSegRecord) {
      current.lineSegRecord = record;
    } else if (record.tag === 71) {
      current.ctrlHeaderRecords.push(record);
    }
  }

  return { paragraphs, recordToParagraph };
}

function getParagraphIndexByRecord(analysis, recordIndex) {
  if (!analysis.recordToParagraph.has(recordIndex)) {
    return null;
  }
  return analysis.recordToParagraph.get(recordIndex);
}

function getParagraphByRecord(analysis, recordIndex) {
  const paragraphIndex = getParagraphIndexByRecord(analysis, recordIndex);
  if (paragraphIndex == null) {
    return null;
  }
  return analysis.paragraphs[paragraphIndex] ?? null;
}

function getRecordPayload(decoded, record) {
  return decoded.subarray(record.payloadOffset, record.payloadOffset + record.size);
}

function parseParaHeader(payload) {
  if (payload.length < 22) {
    return null;
  }

  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const header = {
    textCharCount: dv.getUint32(0, true),
    controlMask: dv.getUint32(4, true),
    paraShapeId: dv.getUint16(8, true),
    paraStyleId: dv.getUint8(10),
    splitType: dv.getUint8(11),
    charShapeCount: dv.getUint16(12, true),
    rangeTagCount: dv.getUint16(14, true),
    lineSegCount: dv.getUint16(16, true),
    instanceId: dv.getUint32(18, true),
  };

  if (payload.length >= 24) {
    header.trackChangeMerge = dv.getUint16(22, true);
  }

  return header;
}

function parseParaCharShape(payload) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const runCount = Math.floor(payload.length / 8);
  const runs = [];

  for (let i = 0; i < runCount; i += 1) {
    const base = i * 8;
    runs.push({
      startPos: dv.getUint32(base, true),
      charShapeId: dv.getUint32(base + 4, true),
    });
  }

  return runs;
}

function parseParaLineSeg(payload) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const segSize = 36;
  const segmentCount = Math.floor(payload.length / segSize);
  const segments = [];

  for (let i = 0; i < segmentCount; i += 1) {
    const base = i * segSize;
    const flags = dv.getUint32(base + 32, true);
    segments.push({
      textStartPos: dv.getUint32(base, true),
      lineVerticalPos: dv.getInt32(base + 4, true),
      lineHeight: dv.getInt32(base + 8, true),
      textHeight: dv.getInt32(base + 12, true),
      lineDistanceToBase: dv.getInt32(base + 16, true),
      lineSpaceBelow: dv.getInt32(base + 20, true),
      horizStartPos: dv.getInt32(base + 24, true),
      lineWidthToColumn: dv.getInt32(base + 28, true),
      flags,
      flagLabels: decodeLineSegFlags(flags),
    });
  }

  return segments;
}

function parseCtrlHeader(payload, record = null) {
  const ctrlBytes = payload.subarray(0, 4);
  const rawId = ctrlBytes.length === 4 ? bytesToAscii(ctrlBytes) : "";
  const ctrlId = rawId ? rawId.split("").reverse().join("") : "";
  const name = CTRL_ID_NAMES[ctrlId] ?? "Unknown Control";
  const detailBytes = payload.subarray(4);
  const dv = new DataView(detailBytes.buffer, detailBytes.byteOffset, detailBytes.byteLength);
  const words = [];
  for (let offset = 0; offset + 3 < detailBytes.length && offset < 16; offset += 4) {
    words.push(`0x${dv.getUint32(offset, true).toString(16)}`);
  }

  return {
    ctrlId,
    name,
    recordIndex: record?.index ?? null,
    level: record?.level ?? "-",
    payloadSize: payload.length,
    words,
    wordsText: words.join(", "),
  };
}

function decodeAdvancedTagInfo(tag, payload) {
  if (!payload?.length) {
    return null;
  }

  if (tag === 32 || tag === 96) {
    const parsed = parseTrackChangeRecord(payload, 0, tag);
    return {
      title: `${RECORD_TAGS[tag] || "TRACK_CHANGE"} (${tag})`,
      text: [
        `flags=${parsed.flags != null ? `0x${parsed.flags.toString(16)}` : "-"}`,
        `authorId=${parsed.authorId ?? "-"}`,
        `timestampLow=${parsed.timestampLow ?? "-"}`,
        `summary=${parsed.summary || "-"}`,
        `strings=${parsed.strings.length ? parsed.strings.join(" | ") : "-"}`,
      ].join("\n"),
    };
  }

  if (tag === 97) {
    const parsed = parseTrackChangeAuthorRecord(payload, 0);
    return {
      title: `${RECORD_TAGS[tag] || "TRACK_CHANGE_AUTHOR"} (${tag})`,
      text: [
        `authorIndex=${parsed.authorIndex ?? "-"}`,
        `name=${parsed.name || "-"}`,
        `strings=${parsed.strings.length ? parsed.strings.join(" | ") : "-"}`,
      ].join("\n"),
    };
  }

  if (tag === 92) {
    const parsed = parseMemoShapeRecord(payload, 0);
    return {
      title: `${RECORD_TAGS[tag] || "MEMO_SHAPE"} (${tag})`,
      text: [
        `property=${parsed.property != null ? `0x${parsed.property.toString(16)}` : "-"}`,
        `width=${parsed.width ?? "-"}`,
        `margin=${parsed.margin ?? "-"}`,
        `strings=${parsed.strings.length ? parsed.strings.join(" | ") : "-"}`,
      ].join("\n"),
    };
  }

  if (tag === 94) {
    const parsed = parseForbiddenCharRecord(payload, 0);
    const sample = parsed.sample?.length
      ? parsed.sample.map((item) => `${item.char}(U+${item.code.toString(16).toUpperCase().padStart(4, "0")})`).join(", ")
      : "-";
    return {
      title: `${RECORD_TAGS[tag] || "FORBIDDEN_CHAR"} (${tag})`,
      text: [`count=${parsed.countByHeader ?? "-"}`, `sample=${sample}`].join("\n"),
    };
  }

  if (tag === 95) {
    const parsed = parseChartData(payload);
    return {
      title: `${RECORD_TAGS[tag] || "CHART_DATA"} (${tag})`,
      text: [
        `chartType=${parsed?.chartType ?? "-"}`,
        `chartId=${parsed?.chartId ?? "-"}`,
        `summary=${parsed?.summary || "-"}`,
        `strings=${parsed?.strings?.length ? parsed.strings.join(" | ") : "-"}`,
      ].join("\n"),
    };
  }

  if (tag === 98) {
    const parsed = parseVideoData(payload, state.doc?.binDataById);
    return {
      title: `${RECORD_TAGS[tag] || "VIDEO_DATA"} (${tag})`,
      text: [
        `flags=${parsed?.flags != null ? `0x${parsed.flags.toString(16)}` : "-"}`,
        `binDataId=${parsed?.binDataId ?? "-"}`,
        `mediaPath=${parsed?.mediaPath || "-"}`,
        `strings=${parsed?.strings?.length ? parsed.strings.join(" | ") : "-"}`,
      ].join("\n"),
    };
  }

  if (tag === 90) {
    const parsed = parseTextArtData(payload);
    return {
      title: `${RECORD_TAGS[tag] || "SHAPE_COMPONENT_TEXTART"} (${tag})`,
      text: [`text=${parsed?.text || "-"}`, `strings=${parsed?.strings?.length ? parsed.strings.join(" | ") : "-"}`].join("\n"),
    };
  }

  if (tag === 91) {
    const parsed = parseFormObjectData(payload);
    return {
      title: `${RECORD_TAGS[tag] || "FORM_OBJECT"} (${tag})`,
      text: [
        `name=${parsed?.name || "-"}`,
        `help=${parsed?.help || "-"}`,
        `strings=${parsed?.strings?.length ? parsed.strings.join(" | ") : "-"}`,
      ].join("\n"),
    };
  }

  if (tag === 93) {
    const parsed = parseListHeader(payload);
    return {
      title: `${RECORD_TAGS[tag] || "MEMO_LIST"} (${tag})`,
      text: parsed
        ? [
            `paragraphCount=${parsed.paragraphCount}`,
            `direction=${parsed.textDirectionName}`,
            `lineBreak=${parsed.lineBreakName}`,
            `verticalAlign=${parsed.verticalAlignName}`,
          ].join("\n")
        : "unable to parse list header payload",
    };
  }

  if (tag === 115) {
    return {
      title: `${RECORD_TAGS[tag] || "SHAPE_COMPONENT_UNKNOWN"} (${tag})`,
      text: [
        `payload=${payload.length} bytes`,
        `ascii=${extractReadableAscii(payload, 200) || "-"}`,
      ].join("\n"),
    };
  }

  if (tag === 896 || tag === 897) {
    const text = tag === 897 ? UTF16LE_DECODER.decode(payload).replace(/\0/g, "") : bytesToAscii(payload);
    return {
      title: `${RECORD_TAGS[tag] || "LEGACY3_BLOCK"} (${tag})`,
      text: normalizePreviewText(text).slice(0, 1500) || "(empty)",
    };
  }

  return null;
}

function decodeLineSegFlags(flags) {
  const names = [];
  const flagValue = flags >>> 0;
  for (const [bitStr, name] of Object.entries(LINE_SEG_FLAGS)) {
    const bit = Number(bitStr) >>> 0;
    if (((flagValue & bit) >>> 0) === bit) {
      names.push(name);
    }
  }
  return names;
}

function renderParagraphHeader(header, docInfo = null) {
  const style = docInfo?.styleById?.get(header.paraStyleId) ?? null;
  const paraShape = docInfo?.paraShapeById?.get(header.paraShapeId) ?? null;
  const styleCharShape = style ? docInfo?.charShapeById?.get(style.charShapeId) ?? null : null;

  const lines = [
    `textCharCount=${header.textCharCount}`,
    `controlMask=0x${header.controlMask.toString(16)}`,
    `paraShapeId=${header.paraShapeId}${paraShape ? ` (${formatParaShapeSummary(paraShape)})` : ""}`,
    `paraStyleId=${header.paraStyleId}${style ? ` (${formatStyleName(style)})` : ""}`,
    `splitType=0x${header.splitType.toString(16)}`,
    `charShapeCount=${header.charShapeCount}`,
    `rangeTagCount=${header.rangeTagCount}`,
    `lineSegCount=${header.lineSegCount}`,
    `instanceId=${header.instanceId}`,
  ];

  if (header.trackChangeMerge != null) {
    lines.push(`trackChangeMerge=${header.trackChangeMerge}`);
  }

  if (style) {
    const styleLang = style.languageId != null ? `0x${style.languageId.toString(16)}` : "?";
    lines.push(`style.type=${style.styleType}, style.next=${style.nextStyleId}, style.lang=${styleLang}`);
    lines.push(
      `style.refs paraShape=${style.paraShapeId}, charShape=${style.charShapeId}${styleCharShape ? ` (${describeDocCharShape(styleCharShape)})` : ""}`
    );
  }

  if (paraShape?.property1Bits) {
    const p1 = paraShape.property1Bits;
    lines.push(
      `para.attr1 align=${PARA_ALIGN_NAMES[p1.align] ?? p1.align}, latinBreak=${PARA_LATIN_BREAK_NAMES[p1.breakLatinWord] ?? p1.breakLatinWord}, koBreak=${PARA_KOREAN_BREAK_NAMES[p1.breakKoreanWord] ?? p1.breakKoreanWord}`
    );
    if (p1.labels?.length) {
      lines.push(`para.attr1 flags=${p1.labels.join(", ")}`);
    }
  }

  if (paraShape?.property3Bits) {
    lines.push(`para.attr3 lineSpacingType=${paraShape.property3Bits.lineSpacingTypeName}`);
  }

  return lines.join("\n");
}

function formatStyleName(style) {
  const local = style.localName?.trim() ?? "";
  const english = style.englishName?.trim() ?? "";
  if (local && english && local !== english) {
    return `${local} / ${english}`;
  }
  return local || english || `STYLE#${style.id}`;
}

function formatParaShapeSummary(paraShape) {
  const pieces = [];
  const align = paraShape.property1Bits?.align;
  if (align != null) {
    pieces.push(`align=${PARA_ALIGN_NAMES[align] ?? align}`);
  }
  if (paraShape.leftMargin != null) {
    pieces.push(`left=${paraShape.leftMargin}`);
  }
  if (paraShape.rightMargin != null) {
    pieces.push(`right=${paraShape.rightMargin}`);
  }
  if (paraShape.indent != null) {
    pieces.push(`indent=${paraShape.indent}`);
  }
  if (paraShape.lineSpacing != null) {
    pieces.push(`line=${paraShape.lineSpacing}`);
  }
  if (paraShape.property3Bits?.lineSpacingTypeName) {
    pieces.push(`lineType=${paraShape.property3Bits.lineSpacingTypeName}`);
  }
  if (paraShape.property1Bits?.headingType) {
    pieces.push(`heading=${PARA_HEADING_TYPE_NAMES[paraShape.property1Bits.headingType] ?? paraShape.property1Bits.headingType}`);
  }
  return pieces.join(", ") || `PARA_SHAPE#${paraShape.id}`;
}

function describeCharShape(charShapeId, docInfo) {
  const shape = docInfo?.charShapeById?.get(charShapeId) ?? null;
  if (!shape) {
    return `CHAR_SHAPE#${charShapeId}`;
  }
  return describeDocCharShape(shape);
}

function describeDocCharShape(shape) {
  const size = shape.baseSize != null ? `${(shape.baseSize / 100).toFixed(1)}pt` : "size=?";
  const font = shape.primaryFont || "font=?";
  const attrs = shape.attributeBits?.labels?.length ? shape.attributeBits.labels.join(",") : "";
  const color = colorRefToText(shape.textColor);

  const parts = [`${font}`, size];
  if (attrs) {
    parts.push(attrs);
  }
  if (shape.textColor) {
    parts.push(`color=${color}`);
  }
  return parts.join(" | ");
}

function colorRefToText(color) {
  if (!color) {
    return "-";
  }
  if (color.auto) {
    return "AUTO";
  }
  return `${color.hex} (${color.r},${color.g},${color.b})`;
}

function decodeParagraphText(payload) {
  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const unitCount = Math.floor(payload.length / 2);
  const chunks = [];
  const extensions = [];
  const tokens = [];
  let logicalPos = 0;

  let i = 0;
  while (i < unitCount) {
    const code = dv.getUint16(i * 2, true);

    if (code === 10 || code === 13) {
      chunks.push("\n");
      tokens.push({
        type: "newline",
        text: "\n",
        logicalPos,
      });
      logicalPos += 1;
      i += 1;
      continue;
    }

    if (code <= 31) {
      const controlSize = getControlCharSize(code);
      if (controlSize === 8 && i + 8 <= unitCount) {
        if (code === 9) {
          chunks.push("\t");
          tokens.push({
            type: "tab",
            text: "\t",
            logicalPos,
          });
          logicalPos += 1;
          i += 8;
          continue;
        }

        const extension = readExtensionControl(payload, i, code);
        const ctrlId = extension?.ctrlId || "";
        const text = `[CTRL:${ctrlId || `U+${code.toString(16).padStart(4, "0")}`}]`;
        chunks.push(text);
        tokens.push({
          type: "control",
          text,
          ctrlId,
          logicalPos,
        });
        extensions.push({
          ctrlId,
          charOffset: logicalPos,
        });
        logicalPos += 1;
        i += 8;
        continue;
      }

      const text = controlCharToText(code);
      chunks.push(text);
      tokens.push({
        type: text === " " ? "space" : "control",
        text,
        ctrlId: "",
        logicalPos,
      });
      logicalPos += 1;
      i += 1;
      continue;
    }

    const text = String.fromCharCode(code);
    chunks.push(text);
    tokens.push({
      type: "char",
      text,
      logicalPos,
    });
    logicalPos += 1;
    i += 1;
  }

  return {
    text: chunks.join(""),
    extensions,
    tokens,
    logicalLength: logicalPos,
  };
}

function getControlCharSize(code) {
  return CONTROL_CHAR_SIZE[code] ?? 1;
}

function controlCharToText(code) {
  if (code === 0x18) {
    return "-";
  }
  if (code === 0x1e || code === 0x1f || code === 0x00) {
    return " ";
  }
  return `[C${code}]`;
}

function readExtensionControl(payload, charOffset, code) {
  const byteOffset = charOffset * 2;
  if (byteOffset + 16 > payload.length) {
    return null;
  }

  const dv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const tailCode = dv.getUint16(byteOffset + 14, true);
  if (tailCode !== code) {
    return null;
  }

  const idBytes = payload.subarray(byteOffset + 2, byteOffset + 6);
  const rawId = bytesToAscii(idBytes).replace(/[^\x20-\x7e]/g, "");
  if (!rawId) {
    return { ctrlId: "" };
  }

  return {
    ctrlId: rawId.split("").reverse().join(""),
  };
}

function extractParagraphText(decoded, records) {
  const lines = [];
  for (const record of records) {
    if (record.tag !== 67) {
      continue;
    }

    const payload = decoded.subarray(record.payloadOffset, record.payloadOffset + record.size);
    const text = decodeParagraphText(payload).text;
    if (text.trim()) {
      lines.push(text);
    }
  }

  return {
    text: lines.join("\n"),
  };
}

function streamParseNote(stream, analysis) {
  const notes = [];
  const formatInfo = state.doc?.formatInfo;
  notes.push(`path=${stream.fullPath}`);
  if (formatInfo) {
    notes.push(`format=${formatInfo.label} (${Math.round(formatInfo.confidence * 100)}%)`);
  }
  notes.push(`record_candidate=${isLikelyRecordStreamPath(stream.fullPath) ? "yes" : "no"}`);
  notes.push(`distribute_stream=${isDistributableEncryptedStream(stream.fullPath) ? "yes" : "no"}`);
  notes.push(`decode_mode=${analysis.mode}`);

  if (analysis.mode !== "raw") {
    notes.push(
      "compressed document was detected from FileHeader bit0, and this stream was decoded using deflate before record parsing"
    );
  }

  if (analysis.records.length) {
    notes.push(
      `record_header=4bytes(10bit tag, 10bit level, 12bit size, with optional 8byte extended header when size==0xFFF)`
    );
    notes.push(`paragraphs=${analysis.paragraphs.length} (linked by HWPTAG_PARA_HEADER=66)`);
  } else if (analysis.mode === "legacy3-heuristic") {
    notes.push("legacy3 heuristic parser extracted text blocks (ASCII/UTF-16) because record parsing was unavailable");
  }

  const docInfo = state.doc?.docInfo;
  if (docInfo) {
    notes.push(
      `docinfo=face:${docInfo.faceNames.length}, charShape:${docInfo.charShapes.length}, paraShape:${docInfo.paraShapes.length}, style:${docInfo.styles.length}`
    );
  }

  return notes.join("\n");
}
