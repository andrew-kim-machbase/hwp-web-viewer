export function selectBestStreamAnalysisCore(candidates, streamPath, deps) {
  const { isLikelyRecordStreamPath, parseRecords, buildParagraphContext } = deps;

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
