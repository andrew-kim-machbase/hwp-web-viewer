export function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

export function hwpToPx(value) {
  if (!Number.isFinite(value)) {
    return 0;
  }
  return value / 80;
}

export function hwpToMm(value) {
  if (!Number.isFinite(value)) {
    return 0;
  }
  return value * (25.4 / 7200);
}

export function toFlagBits(flags) {
  const bits = [];
  for (let i = 0; i < 32; i += 1) {
    if (flags & (1 << i)) {
      bits.push(i);
    }
  }
  return bits.join(", ");
}
