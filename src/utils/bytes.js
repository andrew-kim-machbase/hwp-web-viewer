export function toUint8Array(value) {
  if (!value) {
    return new Uint8Array();
  }
  if (value instanceof Uint8Array) {
    return value;
  }
  return new Uint8Array(value);
}

export function decodeUtf8Sample(bytes, maxBytes = 262144) {
  try {
    const decoder = new TextDecoder("utf-8", { fatal: false });
    return decoder.decode(bytes.subarray(0, Math.min(bytes.length, maxBytes)));
  } catch {
    return "";
  }
}

export function extractReadableAscii(bytes, maxChars = 120) {
  if (!bytes?.length) {
    return "";
  }
  const chars = [];
  for (let i = 0; i < bytes.length && chars.length < maxChars; i += 1) {
    const value = bytes[i];
    if (value === 0x09 || value === 0x0a || value === 0x0d || (value >= 0x20 && value <= 0x7e)) {
      chars.push(String.fromCharCode(value));
    }
  }
  return chars.join("").replace(/\s+/g, " ").trim();
}

export function bytesToAscii(bytes) {
  return Array.from(bytes)
    .map((value) => String.fromCharCode(value))
    .join("");
}
