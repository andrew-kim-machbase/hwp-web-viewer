#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { spawn, spawnSync } from "node:child_process";
import { chromium } from "playwright";

const DEFAULT_CASES = [
  { name: "3.0", hwp: "3.0.hwp", pdf: "3.0.pdf" },
  { name: "5.0", hwp: "5.0.hwp", pdf: "5.0.pdf" },
];

function parseCaseToken(rawValue) {
  const value = String(rawValue ?? "");
  const first = value.indexOf(":");
  if (first <= 0) {
    return null;
  }
  const second = value.indexOf(":", first + 1);
  if (second <= first + 1 || second >= value.length - 1) {
    return null;
  }
  const name = value.slice(0, first).trim();
  const hwp = value.slice(first + 1, second).trim();
  const pdf = value.slice(second + 1).trim();
  if (!name || !hwp || !pdf) {
    return null;
  }
  return { name, hwp, pdf };
}

function discoverCasesFromDir(dirPath, includeHwpx = false) {
  const targetDir = path.resolve(dirPath);
  const entries = fs.readdirSync(targetDir, { withFileTypes: true });
  const pairMap = new Map();
  for (const entry of entries) {
    if (!entry.isFile()) {
      continue;
    }
    const match = entry.name.match(/^(.*)\.(hwp|hwpx|pdf)$/i);
    if (!match) {
      continue;
    }
    const base = match[1];
    const ext = match[2].toLowerCase();
    if (!pairMap.has(base)) {
      pairMap.set(base, {});
    }
    pairMap.get(base)[ext] = entry.name;
  }

  const cases = [];
  for (const [base, value] of pairMap.entries()) {
    if (!value.pdf) {
      continue;
    }
    const source = value.hwp || (includeHwpx ? value.hwpx : null);
    if (!source) {
      continue;
    }
    cases.push({
      name: base,
      hwp: path.join(targetDir, source),
      pdf: path.join(targetDir, value.pdf),
    });
  }
  cases.sort((a, b) => a.name.localeCompare(b.name, "ko"));
  return cases;
}

function sanitizeCaseDirName(name, index) {
  const fallback = `case-${String(index + 1).padStart(2, "0")}`;
  const value = String(name ?? "").trim();
  if (!value) {
    return fallback;
  }
  return value.replace(/[\\/]/g, "_") || fallback;
}

function parseArgs(argv) {
  const args = {
    url: "http://127.0.0.1:4173",
    maxPages: 20,
    artifactsDir: "artifacts/diff",
    python: ".venv/bin/python",
    cases: DEFAULT_CASES,
    docsDir: null,
    caseLimit: null,
    includeHwpx: false,
    saveRendered: false,
    reuseServer: false,
  };
  const explicitCases = [];

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === "--url" && argv[i + 1]) {
      args.url = argv[++i];
    } else if (token === "--max-pages" && argv[i + 1]) {
      args.maxPages = Number.parseInt(argv[++i], 10) || args.maxPages;
    } else if (token === "--artifacts" && argv[i + 1]) {
      args.artifactsDir = argv[++i];
    } else if (token === "--python" && argv[i + 1]) {
      args.python = argv[++i];
    } else if (token === "--docs-dir" && argv[i + 1]) {
      args.docsDir = argv[++i];
    } else if (token === "--case-limit" && argv[i + 1]) {
      const parsed = Number.parseInt(argv[++i], 10);
      if (Number.isFinite(parsed) && parsed > 0) {
        args.caseLimit = parsed;
      }
    } else if (token === "--case" && argv[i + 1]) {
      const caseItem = parseCaseToken(argv[++i]);
      if (caseItem) {
        explicitCases.push(caseItem);
      }
    } else if (token === "--include-hwpx") {
      args.includeHwpx = true;
    } else if (token === "--save-rendered") {
      args.saveRendered = true;
    } else if (token === "--reuse-server") {
      args.reuseServer = true;
    }
  }

  if (explicitCases.length) {
    args.cases = explicitCases;
  } else if (args.docsDir) {
    args.cases = discoverCasesFromDir(args.docsDir, args.includeHwpx);
  }
  if (args.caseLimit != null) {
    args.cases = args.cases.slice(0, args.caseLimit);
  }
  if (!args.cases.length) {
    throw new Error("No valid comparison cases were resolved.");
  }

  return args;
}

async function isServerReady(url) {
  try {
    const res = await fetch(url);
    return res.ok;
  } catch {
    return false;
  }
}

async function waitForServer(url, timeoutMs = 30000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (await isServerReady(url)) {
      return true;
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  return false;
}

function parseHostPort(url) {
  const parsed = new URL(url);
  return {
    host: parsed.hostname,
    port: parsed.port || (parsed.protocol === "https:" ? "443" : "80"),
  };
}

async function ensureDevServer(url, reuseServer = false) {
  if (await isServerReady(url)) {
    if (!reuseServer) {
      throw new Error(
        `Target URL is already serving content: ${url}. Use --reuse-server or pass a different --url to avoid stale comparisons.`
      );
    }
    return { proc: null, owned: false };
  }
  const { host, port } = parseHostPort(url);
  const proc = spawn("npm", ["run", "dev", "--", "--host", host, "--port", port, "--strictPort"], {
    stdio: ["ignore", "pipe", "pipe"],
    env: process.env,
    detached: true,
  });
  proc.stdout.on("data", (chunk) => process.stdout.write(`[vite] ${chunk}`));
  proc.stderr.on("data", (chunk) => process.stderr.write(`[vite] ${chunk}`));

  const ready = await waitForServer(url, 45000);
  if (!ready) {
    proc.kill("SIGTERM");
    throw new Error(`Dev server did not become ready: ${url}`);
  }
  return { proc, owned: true };
}

function stopDevServer(server) {
  if (!server?.owned || !server.proc) {
    return;
  }
  try {
    process.kill(-server.proc.pid, "SIGTERM");
  } catch {
    try {
      server.proc.kill("SIGTERM");
    } catch {
      // ignore
    }
  }
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function resetDir(dirPath) {
  fs.rmSync(dirPath, { recursive: true, force: true });
  fs.mkdirSync(dirPath, { recursive: true });
}

async function waitForStablePreview({ page, selector, pollMs = 300, stablePasses = 8, timeoutMs = 120000 }) {
  const start = Date.now();
  let lastCount = -1;
  let stableCount = 0;

  while (Date.now() - start < timeoutMs) {
    const count = await page.locator(selector).count();
    if (count > 0 && count === lastCount) {
      stableCount += 1;
      if (stableCount >= stablePasses) {
        return count;
      }
    } else {
      stableCount = 0;
      lastCount = count;
    }
    await page.waitForTimeout(pollMs);
  }

  throw new Error(`Preview pages did not stabilize for selector: ${selector}`);
}

async function waitForLocatorImagesReady(locator, timeoutMs = 8000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    let ready = false;
    try {
      ready = await locator.evaluate((element) => {
        const images = Array.from(element.querySelectorAll("img"));
        if (!images.length) {
          return true;
        }
        return images.every((image) => image.complete);
      });
    } catch {
      ready = false;
    }
    if (ready) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 120));
  }
}

async function capturePreviewPages({ page, caseItem, outDir, maxPages }) {
  const hwpPath = path.resolve(caseItem.hwp);
  resetDir(outDir);
  await page.locator("#fileInput").setInputFiles(hwpPath);
  await page.waitForSelector(".preview-pages .preview-page", { timeout: 120000 });
  await waitForStablePreview({
    page,
    selector: ".preview-pages .preview-page",
    pollMs: 320,
    stablePasses: 8,
    timeoutMs: 120000,
  });
  await page.waitForTimeout(250);

  const pages = page.locator(".preview-pages .preview-page");
  const count = await pages.count();
  const captureCount = Math.min(count, Math.max(1, maxPages));
  for (let i = 0; i < captureCount; i += 1) {
    const target = pages.nth(i);
    await target.scrollIntoViewIfNeeded();
    await page.waitForTimeout(120);
    await waitForLocatorImagesReady(target, 10000);
    await target.screenshot({
      path: path.join(outDir, `preview-page-${String(i + 1).padStart(3, "0")}.png`),
    });
  }

  return { renderedPages: count, capturedPages: captureCount };
}

function runPdfComparison({ python, previewDir, pdfPath, outJson, maxPages, saveRendered }) {
  const scriptPath = path.resolve("scripts/pdf_compare.py");
  const args = [
    scriptPath,
    "--preview-dir",
    previewDir,
    "--pdf",
    path.resolve(pdfPath),
    "--out-json",
    outJson,
    "--max-pages",
    String(maxPages),
  ];
  if (saveRendered) {
    args.push("--save-rendered");
  }
  const result = spawnSync(python, args, { stdio: "pipe", encoding: "utf-8" });
  if (result.status !== 0) {
    throw new Error(`pdf_compare.py failed (${result.status}): ${result.stderr || result.stdout}`);
  }
  return JSON.parse(fs.readFileSync(outJson, "utf-8"));
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  ensureDir(args.artifactsDir);

  const server = await ensureDevServer(args.url, args.reuseServer);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 2400 } });

  const summary = [];
  try {
    await page.goto(args.url, { waitUntil: "networkidle" });
    for (let index = 0; index < args.cases.length; index += 1) {
      const caseItem = args.cases[index];
      const hwpPath = path.resolve(caseItem.hwp);
      const pdfPath = path.resolve(caseItem.pdf);
      if (!fs.existsSync(hwpPath)) {
        throw new Error(`Case "${caseItem.name}" source file not found: ${hwpPath}`);
      }
      if (!fs.existsSync(pdfPath)) {
        throw new Error(`Case "${caseItem.name}" reference PDF not found: ${pdfPath}`);
      }

      const caseDir = path.join(args.artifactsDir, sanitizeCaseDirName(caseItem.name, index));
      const previewDir = path.join(caseDir, "preview");
      const reportPath = path.join(caseDir, "report.json");
      ensureDir(caseDir);

      const capture = await capturePreviewPages({
        page,
        caseItem,
        outDir: previewDir,
        maxPages: args.maxPages,
      });

      const report = runPdfComparison({
        python: args.python,
        previewDir,
        pdfPath,
        outJson: reportPath,
        maxPages: args.maxPages,
        saveRendered: args.saveRendered,
      });

      summary.push({
        name: caseItem.name,
        hwp: hwpPath,
        pdf: pdfPath,
        renderedPages: capture.renderedPages,
        capturedPages: capture.capturedPages,
        comparedPages: report.pages_compared,
        avgRms: report.avg_rms,
        avgMae: report.avg_mae,
      });
    }

    const summaryPath = path.join(args.artifactsDir, "summary.json");
    fs.writeFileSync(summaryPath, JSON.stringify(summary, null, 2), "utf-8");
    console.log(`Saved: ${summaryPath}`);
    console.table(summary);
  } finally {
    await browser.close();
    stopDevServer(server);
  }
}

main().catch((error) => {
  console.error(error.stack || String(error));
  process.exit(1);
});
