export function selectBestStreamAnalysisCore(candidates, streamPath, deps, options = {}) {
  const { isLikelyRecordStreamPath, parseRecords, buildParagraphContext } = deps;
  const forceRecordParse = Boolean(options.forceRecordParse);

  let best = null;
  for (const candidate of candidates) {
    const parse = (forceRecordParse || isLikelyRecordStreamPath(streamPath))
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

function looksPlausibleLegacyRecordParse(result) {
  if (!result || !result.records?.length) {
    return false;
  }
  if (result.parse?.complete) {
    return true;
  }

  const decodedLength = result.decoded?.length ?? 0;
  const consumed = result.parse?.consumed ?? 0;
  if (!decodedLength || consumed <= 0) {
    return false;
  }
  const ratio = consumed / decodedLength;
  if (result.paragraphs?.length && ratio >= 0.18) {
    return true;
  }
  if (result.records.length >= 8 && ratio >= 0.7) {
    return true;
  }
  if (result.records.length >= 32 && ratio >= 0.35) {
    return true;
  }
  return false;
}

export function analyzeStreamCore(stream, compressedFlag, formatMode = "unknown", distributableFlag = false, deps) {
  const {
    isLikelyRecordStreamPath,
    mayBeCompressedRecordStreamPath,
    safeInflate,
    isDistributableEncryptedStream,
    decodeDistributableStreamCandidates,
    analyzeLegacy3Stream,
    parseRecords,
    buildParagraphContext,
  } = deps;

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
      return selectBestStreamAnalysisCore(distCandidates, stream.fullPath, {
        isLikelyRecordStreamPath,
        parseRecords,
        buildParagraphContext,
      });
    }
  }

  if (formatMode === "legacy-3x-binary" && !isLikelyRecordStreamPath(stream.fullPath)) {
    const legacyCandidates = [{ mode: "raw", bytes: raw }];
    if (compressedFlag) {
      const rawInflated = safeInflate(raw, true);
      if (rawInflated) {
        legacyCandidates.push({ mode: "inflateRaw(-15)", bytes: rawInflated });
      }
      const zlibInflated = safeInflate(raw, false);
      if (zlibInflated) {
        legacyCandidates.push({ mode: "inflate(zlib)", bytes: zlibInflated });
      }
    }

    const legacyParsed = selectBestStreamAnalysisCore(
      legacyCandidates,
      stream.fullPath,
      {
        isLikelyRecordStreamPath,
        parseRecords,
        buildParagraphContext,
      },
      { forceRecordParse: true }
    );
    if (looksPlausibleLegacyRecordParse(legacyParsed)) {
      return legacyParsed;
    }
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

  const best = selectBestStreamAnalysisCore(candidates, stream.fullPath, {
    isLikelyRecordStreamPath,
    parseRecords,
    buildParagraphContext,
  });

  if (best && best.records.length === 0 && formatMode === "legacy-3x-binary") {
    return analyzeLegacy3Stream(raw);
  }

  return best;
}
