/**
 * Base64 encoding utilities
 */

/**
 * Base64 lookup table for fast encoding
 */
const BASE64_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/**
 * Convert Uint8Array to base64 string using lookup table and bitwise operations
 * Similar to base64-arraybuffer implementation but without external dependencies
 * Performance: O(n) time complexity, no stack overflow risk
 *
 * @param bytes - Uint8Array to encode
 * @returns Base64 encoded string
 */
export function uint8ArrayToBase64(bytes: Uint8Array): string {
  const len = bytes.length;
  const resultLen = Math.ceil(len / 3) * 4;
  const result = Array.from<string>({ length: resultLen });
  let resultIndex = 0;

  // Process 3 bytes at a time (24 bits -> 4 base64 chars)
  for (let i = 0; i < len; i += 3) {
    // Read 3 bytes (24 bits)
    const byte1 = bytes[i];
    const byte2 = i + 1 < len ? bytes[i + 1] : 0;
    const byte3 = i + 2 < len ? bytes[i + 2] : 0;

    // Extract 4 x 6-bit values using bitwise operations
    const index0 = byte1 >> 2;
    const index1 = ((byte1 & 0x03) << 4) | (byte2 >> 4);
    const index2 = ((byte2 & 0x0f) << 2) | (byte3 >> 6);
    const index3 = byte3 & 0x3f;

    // Encode to base64 characters using lookup table
    result[resultIndex++] = BASE64_CHARS[index0];
    result[resultIndex++] = BASE64_CHARS[index1];
    result[resultIndex++] = i + 1 < len ? BASE64_CHARS[index2] : "=";
    result[resultIndex++] = i + 2 < len ? BASE64_CHARS[index3] : "=";
  }

  return result.join("");
}

/**
 * Convert base64 string to Uint8Array
 *
 * @param base64 - Base64 encoded string
 * @returns Decoded data as Uint8Array
 */
export function base64ToUint8Array(base64: string): Uint8Array {
  const binaryString = atob(base64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}
