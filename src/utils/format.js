export function formatHex(bytes, base = 0, limit = 1024) {
  const len = Math.min(bytes.length, limit);
  const rows = [];

  for (let i = 0; i < len; i += 16) {
    const slice = bytes.subarray(i, i + 16);
    const hexPart = Array.from(slice)
      .map((byte) => byte.toString(16).padStart(2, "0"))
      .join(" ")
      .padEnd(16 * 3 - 1, " ");

    const asciiPart = Array.from(slice)
      .map((byte) => (byte >= 32 && byte <= 126 ? String.fromCharCode(byte) : "."))
      .join("");

    rows.push(`${(base + i).toString(16).padStart(8, "0")}  ${hexPart}  ${asciiPart}`);
  }

  if (bytes.length > limit) {
    rows.push(`... ${bytes.length - limit} more bytes omitted ...`);
  }

  return rows.join("\n");
}

export function formatBytes(size) {
  if (size < 1024) {
    return `${size} B`;
  }
  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`;
  }
  return `${(size / (1024 * 1024)).toFixed(2)} MB`;
}
