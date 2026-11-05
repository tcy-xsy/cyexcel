/** Reads the exporter's uncompressed ZIP entries for semantic OOXML assertions. */
export function readXlsxEntries(bytes: Uint8Array): Map<string, string> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder();
  const files = new Map<string, string>();
  let offset = 0;
  while (view.getUint32(offset, true) === 0x04034b50) {
    if (view.getUint16(offset + 8, true) !== 0) throw new Error("Expected stored ZIP entries");
    const size = view.getUint32(offset + 18, true);
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const dataStart = offset + 30 + nameLength + extraLength;
    files.set(decoder.decode(bytes.subarray(offset + 30, offset + 30 + nameLength)),
      decoder.decode(bytes.subarray(dataStart, dataStart + size)));
    offset = dataStart + size;
  }
  if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("Missing ZIP central directory");
  return files;
}
