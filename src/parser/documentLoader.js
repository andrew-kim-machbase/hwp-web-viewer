import * as CFB from "cfb";
import { toUint8Array, decodeUtf8Sample, bytesToAscii } from "../utils/bytes.js";
import { isLikelyRecordStreamPath } from "./streamPath.js";

export function parseFileHeader(bytes) {
  if (!bytes || bytes.length < 40) {
    return null;
  }

  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const signature = bytesToAscii(bytes.subarray(0, 32)).replace(/\0+$/g, "").trim();
  const v0 = dv.getUint8(32);
  const v1 = dv.getUint8(33);
  const v2 = dv.getUint8(34);
  const v3 = dv.getUint8(35);
  const flags = dv.getUint32(36, true);

  return {
    signature,
    versionBytes: [v0, v1, v2, v3],
    versionDisplay: `${v3}.${v2}.${v1}.${v0}`,
    flags,
    compressed: Boolean(flags & 0x00000001),
    passwordProtected: Boolean(flags & 0x00000002),
    distributable: Boolean(flags & 0x00000004),
  };
}

export function tryReadCfbEntries(bytes) {
  try {
    const cfb = CFB.read(bytes, { type: "array" });
    const entries = cfb.FullPaths.map((fullPath, index) => {
      const fileIndex = cfb.FileIndex[index];
      const normalizedPath = fullPath.endsWith("/") ? fullPath.slice(0, -1) : fullPath;
      return {
        index,
        fullPath: normalizedPath,
        name: fileIndex?.name ?? normalizedPath.split("/").at(-1),
        type: fileIndex?.type ?? 0,
        size: fileIndex?.size ?? 0,
        content: toUint8Array(fileIndex?.content),
      };
    });
    return {
      success: true,
      error: null,
      entries,
    };
  } catch (error) {
    return {
      success: false,
      error,
      entries: [],
    };
  }
}

export function detectDocumentFormat({ bytes, cfbReadSucceeded, cfbError, entries, fileHeader, docInfo }) {
  const sampleAscii = bytesToAscii(bytes.subarray(0, Math.min(bytes.length, 65536)));
  const sampleUtf8 = decodeUtf8Sample(bytes, 262144).toLowerCase();
  const hasBodySections = entries.some((entry) => /\/BodyText\/Section\d+$/.test(entry.fullPath));
  const hasRecordDocInfo = Boolean(docInfo?.records?.length);
  const hasHwpFileHeader = Boolean(fileHeader?.signature?.includes("HWP Document File"));
  const legacy3Signature = /HWP\s+Document\s+File\s+V3\./i.test(sampleAscii);
  const hwpmlSignature = sampleUtf8.includes("<hwpml") || sampleUtf8.includes("owpml");

  let mode = "unknown";
  let label = "Unknown / Needs Manual Analysis";
  let confidence = 0.35;
  let note = "Container and stream heuristics did not match a single known profile.";

  if (hwpmlSignature && !cfbReadSucceeded) {
    mode = "hwpml-xml";
    label = "HWPML XML";
    confidence = 0.95;
    note = "Raw XML signature detected (`<HWPML>` or OWPML reference).";
  } else if (legacy3Signature && !cfbReadSucceeded) {
    mode = "legacy-3x-binary";
    label = "Legacy 3.x Binary";
    confidence = 0.9;
    note = "Legacy signature (`HWP Document File V3.x`) detected in raw bytes.";
  } else if (hasHwpFileHeader && (hasRecordDocInfo || hasBodySections)) {
    mode = "hwp5-record";
    label = fileHeader?.distributable ? "HWP 5.x Distributable Record Stream" : "HWP 5.x Record Stream";
    confidence = hasRecordDocInfo ? 0.96 : 0.82;
    note = fileHeader?.distributable
      ? "Distributable flag is set; ViewText streams require distribute-document decode before record parsing."
      : hasRecordDocInfo
        ? "DocInfo/BodyText streams parse as TagID/Level/Size record structure."
        : "CFB container with HWP header and section streams detected.";
  } else if (legacy3Signature) {
    mode = "legacy-3x-binary";
    label = "Legacy 3.x Binary (Inside Container)";
    confidence = 0.62;
    note = "Legacy 3.x signature detected, but CFB-level parsing is also present.";
  } else if (hwpmlSignature) {
    mode = "hwpml-xml";
    label = "HWPML XML (Embedded)";
    confidence = 0.7;
    note = "HWPML markers detected from text sample.";
  } else if (cfbReadSucceeded) {
    mode = "hwp-container-unknown";
    label = "HWP Container (Unknown Variant)";
    confidence = 0.52;
    note = "CFB/OLE container opened, but record-level profile is unclear.";
  } else if (cfbError) {
    mode = "raw-unknown";
    label = "Raw Binary (Non-CFB)";
    confidence = 0.45;
    note = "CFB/OLE parsing failed; using raw fallback mode.";
  }

  const modeClass = mode.replace(/[^a-z0-9_-]/gi, "-");
  const containerLabel = cfbReadSucceeded ? "CFB/OLE" : "Raw";
  return {
    mode,
    modeClass,
    label,
    confidence,
    note,
    containerLabel,
  };
}

export function buildDocumentFromBytes({ fileName, bytes, parseDocInfo, buildBinDataCatalog }) {
  const readResult = tryReadCfbEntries(bytes);
  const entries = readResult.entries.length
    ? readResult.entries
    : [
        {
          index: 0,
          fullPath: "Legacy3/RawData",
          name: "RawData",
          type: 2,
          size: bytes.length,
          content: bytes,
        },
      ];

  const fileHeaderEntry = entries.find((entry) => entry.fullPath === "Root Entry/FileHeader");
  const fileHeader = fileHeaderEntry ? parseFileHeader(fileHeaderEntry.content) : null;
  const docInfoEntry = entries.find((entry) => entry.fullPath === "Root Entry/DocInfo");
  const docInfo = docInfoEntry ? parseDocInfo(docInfoEntry.content, Boolean(fileHeader?.compressed)) : null;
  const formatInfo = detectDocumentFormat({
    bytes,
    cfbReadSucceeded: readResult.success,
    cfbError: readResult.error,
    entries,
    fileHeader,
    docInfo,
  });
  const binDataCatalog = buildBinDataCatalog(entries, docInfo);

  const defaultStream = entries.find((entry) => entry.type === 2 && isLikelyRecordStreamPath(entry.fullPath));
  const selectedEntryIndex = defaultStream?.index ?? entries.find((entry) => entry.type === 2)?.index ?? null;

  return {
    doc: {
      fileName,
      rawBytes: bytes,
      entries,
      fileHeader,
      docInfo,
      formatInfo,
      binDataList: binDataCatalog.list,
      binDataById: binDataCatalog.byId,
      streamAnalysis: new Map(),
    },
    selectedEntryIndex,
  };
}
