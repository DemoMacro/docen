import { decodeBase64, encodeBase64 } from "@office-open/core/util";

const BINARY_TAG = "$officeOpenBinary";
const BINARY_MIME = "application/octet-stream";

function bytesOf(value: unknown): Uint8Array | undefined {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  return undefined;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isBinaryTag(value: unknown): value is { [BINARY_TAG]: string } {
  return isPlainObject(value) && typeof value[BINARY_TAG] === "string";
}

function dataUrlOf(bytes: Uint8Array, contentType?: string): string {
  return `data:${contentType || BINARY_MIME};base64,${encodeBase64(bytes)}`;
}

function assertJsonSafeBinary(value: unknown): void {
  if (value instanceof Blob || value instanceof ReadableStream) {
    throw new TypeError("Binary input must be a Uint8Array before document JSON serialization.");
  }
}

/** Serialize binary values with a JSON-safe binary tag. */
export function serializeJsonBinary(value: unknown): unknown {
  assertJsonSafeBinary(value);
  const bytes = bytesOf(value);
  if (bytes) return { [BINARY_TAG]: encodeBase64(bytes) };
  if (Array.isArray(value)) return value.map((item) => serializeJsonBinary(item));
  if (!isPlainObject(value)) return value;
  return Object.fromEntries(
    Object.entries(value).map(([key, item]) => [key, serializeJsonBinary(item)]),
  );
}

/** Make raw OPC parts JSON-safe without changing their package path/content type. */
export function serializeRawParts(value: unknown): unknown {
  if (!Array.isArray(value)) return value;
  return value.map((part) => {
    if (!isPlainObject(part)) return part;
    const contentType = typeof part.contentType === "string" ? part.contentType : undefined;
    assertJsonSafeBinary(part.data);
    const bytes = bytesOf(part.data);
    return { ...part, data: bytes ? dataUrlOf(bytes, contentType) : part.data };
  });
}

/** Restore values produced by serializeJsonBinary. */
export function parseJsonBinary(value: unknown): unknown {
  if (isBinaryTag(value)) return decodeBase64(value[BINARY_TAG]);
  if (Array.isArray(value)) return value.map(parseJsonBinary);
  if (!isPlainObject(value)) return value;
  return Object.fromEntries(
    Object.entries(value).map(([key, item]) => [key, parseJsonBinary(item)]),
  );
}
