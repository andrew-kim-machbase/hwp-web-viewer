import { toFlagBits } from "../utils/numeric.js";
import { escapeHtml, escapeHtmlAttr } from "../utils/html.js";

export function renderSummaryPanel(summaryEl, doc) {
  if (!summaryEl) {
    return;
  }

  if (!doc) {
    summaryEl.innerHTML = `<div class="summary-empty">Load a .hwp file to inspect streams and records.</div>`;
    return;
  }

  const { fileHeader, formatInfo } = doc;
  const flagBits = fileHeader ? toFlagBits(fileHeader.flags) : "";
  const docInfo = doc.docInfo;
  const formatBanner = formatInfo
    ? `
      <div class="format-banner mode-${escapeHtmlAttr(formatInfo.modeClass)}">
        <div class="format-banner-head">
          <strong>${escapeHtml(formatInfo.label)}</strong>
          <span>${escapeHtml(`${Math.round(formatInfo.confidence * 100)}%`)}</span>
        </div>
        <div class="format-banner-note">${escapeHtml(formatInfo.note)}</div>
      </div>
    `
    : "";
  summaryEl.innerHTML = `
    ${formatBanner}
    <div class="summary-grid">
      <div class="kv"><span>File</span><strong>${escapeHtml(doc.fileName)}</strong></div>
      <div class="kv"><span>Container</span><strong>${doc.formatInfo?.containerLabel || "-"}</strong></div>
      <div class="kv"><span>Signature</span><strong>${escapeHtml(fileHeader?.signature || "(none)")}</strong></div>
      <div class="kv"><span>Version</span><strong>${escapeHtml(fileHeader?.versionDisplay || "(unknown)")}</strong></div>
      <div class="kv"><span>Flags</span><strong>${fileHeader ? `0x${fileHeader.flags.toString(16)}` : "-"}</strong></div>
      <div class="kv"><span>Compressed</span><strong>${fileHeader ? (fileHeader.compressed ? "yes" : "no") : "-"}</strong></div>
      <div class="kv"><span>Password</span><strong>${fileHeader ? (fileHeader.passwordProtected ? "yes" : "no") : "-"}</strong></div>
      <div class="kv"><span>Distributable</span><strong>${fileHeader ? (fileHeader.distributable ? "yes" : "no") : "-"}</strong></div>
      <div class="kv"><span>Set Bits</span><strong>${fileHeader ? flagBits || "none" : "-"}</strong></div>
      <div class="kv"><span>Streams</span><strong>${doc.entries.length}</strong></div>
      ${
        docInfo
          ? `<div class="kv"><span>DocInfo</span><strong>${escapeHtml(
              `${docInfo.decodeMode}, rec=${docInfo.records.length}`
            )}</strong></div>`
          : ""
      }
      ${
        docInfo
          ? `<div class="kv"><span>FaceNames</span><strong>${docInfo.faceNames.length}</strong></div>`
          : ""
      }
      ${
        docInfo
          ? `<div class="kv"><span>CharShapes</span><strong>${docInfo.charShapes.length}</strong></div>`
          : ""
      }
      ${
        docInfo
          ? `<div class="kv"><span>ParaShapes</span><strong>${docInfo.paraShapes.length}</strong></div>`
          : ""
      }
      ${
        docInfo
          ? `<div class="kv"><span>Styles</span><strong>${docInfo.styles.length}</strong></div>`
          : ""
      }
      ${
        docInfo
          ? `<div class="kv"><span>TabDef</span><strong>${docInfo.tabDefs.length}</strong></div>`
          : ""
      }
      ${
        docInfo
          ? `<div class="kv"><span>Numbering</span><strong>${docInfo.numberings.length}</strong></div>`
          : ""
      }
      ${
        docInfo
          ? `<div class="kv"><span>Bullet</span><strong>${docInfo.bullets.length}</strong></div>`
          : ""
      }
      ${
        docInfo
          ? `<div class="kv"><span>BinData Rec</span><strong>${docInfo.binDataRecords.length}</strong></div>`
          : ""
      }
      ${
        doc.binDataList?.length
          ? `<div class="kv"><span>BinData Streams</span><strong>${doc.binDataList.length}</strong></div>`
          : ""
      }
    </div>
  `;
}
