export function renderDetailPanelView({
  detailPanelEl,
  stream,
  analysis,
  selectedRecordIndex,
  docInfo,
  getRecordPayload,
  formatHex,
  getRecordTagName,
  decodeParagraphText,
  getParagraphByRecord,
  parseParaHeader,
  parseParaCharShape,
  parseParaLineSeg,
  parseCtrlHeader,
  extractParagraphText,
  decodeAdvancedTagInfo,
  escapeHtml,
  renderParagraphHeader,
  describeCharShape,
  streamParseNote,
}) {
  const records = analysis.records;
  const selectedRecord =
    selectedRecordIndex != null && records[selectedRecordIndex] ? records[selectedRecordIndex] : records[0];

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
  const paragraphHeader = paragraph?.headerRecord ? parseParaHeader(getRecordPayload(analysis.decoded, paragraph.headerRecord)) : null;
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
