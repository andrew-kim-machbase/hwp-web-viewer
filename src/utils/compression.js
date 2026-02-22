import { inflate, inflateRaw } from "pako";

export function safeInflate(bytes, rawMode) {
  try {
    return rawMode ? inflateRaw(bytes) : inflate(bytes);
  } catch {
    return null;
  }
}
