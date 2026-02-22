export function renderStreamListPanel({
  streamListEl,
  doc,
  selectedEntryIndex,
  formatBytes,
  escapeHtml,
  onSelectEntry,
}) {
  if (!doc) {
    streamListEl.innerHTML = `<div class="empty">No stream.</div>`;
    return;
  }

  const streamEntries = doc.entries.filter((entry) => entry.fullPath && entry.fullPath !== "Root Entry");
  streamListEl.innerHTML = streamEntries
    .map((entry) => {
      const depth = entry.fullPath.split("/").length - 2;
      const selectedClass = entry.index === selectedEntryIndex ? "selected" : "";
      const kind = entry.type === 1 ? "DIR" : entry.type === 2 ? "STR" : "OBJ";
      const pathLabel = escapeHtml(entry.fullPath.replace(/^Root Entry\/?/, ""));
      return `
        <button class="stream-item ${selectedClass}" data-entry-index="${entry.index}" style="--indent:${depth}">
          <span class="kind">${kind}</span>
          <span class="path">${pathLabel || "/"}</span>
          <span class="size">${entry.type === 2 ? formatBytes(entry.content.length) : ""}</span>
        </button>
      `;
    })
    .join("");

  streamListEl.querySelectorAll(".stream-item").forEach((button) => {
    button.addEventListener("click", () => {
      onSelectEntry(Number(button.dataset.entryIndex));
    });
  });
}

export function renderRecordPanelView({
  recordPanelEl,
  detailPanelEl,
  stream,
  analysis,
  selectedRecordIndex,
  getRecordTagName,
  getParagraphIndexByRecord,
  formatBytes,
  formatHex,
  escapeHtml,
  onSelectRecord,
  renderDetailPanel,
}) {
  if (!stream) {
    recordPanelEl.innerHTML = `<div class="empty">Select a stream.</div>`;
    detailPanelEl.innerHTML = `<div class="empty">Record payload / decoded text will appear here.</div>`;
    return;
  }

  const records = analysis.records;
  const streamMetaHtml = `
    <div class="stream-meta">
      <div><span>Path</span><strong>${escapeHtml(stream.fullPath)}</strong></div>
      <div><span>Raw Size</span><strong>${formatBytes(stream.content.length)}</strong></div>
      <div><span>Decoded Mode</span><strong>${escapeHtml(analysis.mode)}</strong></div>
      <div><span>Decoded Size</span><strong>${formatBytes(analysis.decoded.length)}</strong></div>
    </div>
  `;

  if (!records.length) {
    recordPanelEl.innerHTML = `${streamMetaHtml}<pre class="hex">${escapeHtml(formatHex(analysis.decoded, 0, 1024))}</pre>`;
    detailPanelEl.innerHTML = `<div class="empty">No record structure detected for this stream.</div>`;
    return;
  }

  const table = records
    .map((record, index) => {
      const selectedClass = selectedRecordIndex === index ? "selected" : "";
      const tagName = getRecordTagName(record.tag);
      const paragraphIndex = getParagraphIndexByRecord(analysis, record.index);
      return `
        <tr class="record-row ${selectedClass}" data-record-index="${index}">
          <td>${index}</td>
          <td>${paragraphIndex != null ? paragraphIndex : "-"}</td>
          <td>0x${record.offset.toString(16)}</td>
          <td>${record.tag}</td>
          <td>${escapeHtml(tagName)}</td>
          <td>${record.level}</td>
          <td>${record.size}</td>
        </tr>
      `;
    })
    .join("");

  const parseState = analysis.parse.complete
    ? "complete"
    : `partial (${analysis.parse.consumed}/${analysis.decoded.length} bytes)`;

  recordPanelEl.innerHTML = `
    ${streamMetaHtml}
    <div class="record-summary">
      <span>Records: <strong>${records.length}</strong></span>
      <span>Paragraphs: <strong>${analysis.paragraphs.length}</strong></span>
      <span>Parse: <strong>${parseState}</strong></span>
    </div>
    <div class="record-table-wrap">
      <table class="record-table">
        <thead>
          <tr>
            <th>#</th>
            <th>Para</th>
            <th>Offset</th>
            <th>Tag</th>
            <th>Name</th>
            <th>Level</th>
            <th>Size</th>
          </tr>
        </thead>
        <tbody>${table}</tbody>
      </table>
    </div>
  `;

  recordPanelEl.querySelectorAll(".record-row").forEach((row) => {
    row.addEventListener("click", () => {
      onSelectRecord(Number(row.dataset.recordIndex));
    });
  });

  renderDetailPanel(stream, analysis);
}
