const APP_SHELL_HTML = `
  <div class="bg-orb orb-1"></div>
  <div class="bg-orb orb-2"></div>
  <div class="bg-grid"></div>

  <header class="topbar panel reveal-1">
    <div class="title-wrap">
      <h1>HWP Raw Viewer</h1>
      <p>CFB stream explorer + record parser + paragraph text extraction</p>
    </div>
    <label class="file-picker">
      <input id="fileInput" type="file" accept=".hwp" />
      <span>Select .hwp</span>
    </label>
  </header>

  <section class="summary panel reveal-2" id="summary">
    <div class="summary-empty">Load a .hwp file to inspect streams and records.</div>
  </section>

  <section class="panel reveal-2" id="docInfoPanel">
    <div class="pane-title">DocInfo Catalog</div>
    <div class="pane-body">
      <div class="summary-empty">DocInfo tables will appear after loading a file.</div>
    </div>
  </section>

  <main class="layout reveal-3">
    <section class="panel pane" id="streamPane">
      <div class="pane-title">Streams</div>
      <div class="pane-body" id="streamList"></div>
    </section>

    <section class="panel pane" id="recordPane">
      <div class="pane-title">Records</div>
      <div class="pane-body" id="recordPanel">
        <div class="empty">Select a stream.</div>
      </div>
    </section>

    <section class="panel pane" id="detailPane">
      <div class="pane-title">Detail</div>
      <div class="pane-body" id="detailPanel">
        <div class="empty">Record payload / decoded text will appear here.</div>
      </div>
    </section>
  </main>

  <section class="panel reveal-3" id="previewPanel">
    <div class="pane-title">Document Preview (Beta)</div>
    <div class="pane-body">
      <div class="summary-empty">Preview will appear after parsing BodyText sections.</div>
    </div>
  </section>
`;

export function renderAppShell(appEl) {
  if (!appEl) {
    throw new Error("App root element `#app` was not found.");
  }

  appEl.innerHTML = APP_SHELL_HTML;

  return {
    fileInput: appEl.querySelector("#fileInput"),
    summaryEl: appEl.querySelector("#summary"),
    docInfoPanelEl: appEl.querySelector("#docInfoPanel .pane-body"),
    previewPanelEl: appEl.querySelector("#previewPanel .pane-body"),
    streamListEl: appEl.querySelector("#streamList"),
    recordPanelEl: appEl.querySelector("#recordPanel"),
    detailPanelEl: appEl.querySelector("#detailPanel"),
  };
}
