export function isLikelyRecordStreamPath(path) {
  return /\/DocInfo$|\/BodyText\/Section\d+$|\/ViewText\/Section\d+$/.test(path);
}

export function mayBeCompressedRecordStreamPath(path) {
  return isLikelyRecordStreamPath(path);
}
