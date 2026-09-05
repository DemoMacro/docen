// PDF export — the paginated canvas snapshots (printSnapshots) flatten into a
// PDF 1.4 file: one opaque JPEG image XObject per page, drawn at the page's
// full MediaBox. Hand-rolled object/xref layout, zero dependencies — the
// browser canvas produces the JPEG, the writer only assembles bytes.

/** One page snapshot from the stage's printSnapshots(): the page's CSS-pixel
 *  paper size and its PNG data URL. */
export interface PdfPageShot {
  width: number;
  height: number;
  url: string;
}

/** Decode a snapshot PNG and flatten it onto white — the print canvases are
 *  transparent, JPEG has no alpha, and PDF viewers composite images on black,
 *  so the alpha must become paper white before encoding. Returns the JPEG
 *  bytes. */
async function jpegOf(
  shot: PdfPageShot,
): Promise<{ jpeg: Uint8Array<ArrayBuffer>; width: number; height: number }> {
  const img = new Image();
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("page snapshot failed to decode"));
    img.src = shot.url;
  });
  const canvas = document.createElement("canvas");
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("canvas 2d context unavailable");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0);
  const dataUrl = canvas.toDataURL("image/jpeg", 0.92);
  const bin = atob(dataUrl.slice(dataUrl.indexOf(",") + 1));
  const jpeg = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) jpeg[i] = bin.charCodeAt(i);
  return { jpeg, width: img.naturalWidth, height: img.naturalHeight };
}

/** Serialize the whole file: string parts stream through TextEncoder, binary
 *  parts (the JPEG streams) pass through as-is; every object's byte offset is
 *  recorded for the xref table. */
function pdfWriter(): {
  push: (part: string | Uint8Array) => void;
  offset: () => number;
  bytes: () => Uint8Array<ArrayBuffer>;
} {
  const encoder = new TextEncoder();
  // The intermediate chunks stay ArrayBufferLike (TextEncoder's output); only
  // the joined output is a fresh, exactly-typed buffer for the Blob.
  const chunks: Uint8Array[] = [];
  let length = 0;
  const push = (part: string | Uint8Array): void => {
    const bytes = typeof part === "string" ? encoder.encode(part) : part;
    chunks.push(bytes);
    length += bytes.length;
  };
  const bytes = (): Uint8Array<ArrayBuffer> => {
    const out = new Uint8Array(new ArrayBuffer(length));
    let at = 0;
    for (const c of chunks) {
      out.set(c, at);
      at += c.length;
    }
    return out;
  };
  return { push, offset: () => length, bytes };
}

/** The 10-digit zero-padded xref entry offset. */
const xrefAt = (offset: number): string => String(offset).padStart(10, "0");

/** Build the PDF blob from page snapshots. Pages keep their paper size
 *  (CSS px → pt at 72/96); the image fills the page exactly. */
export async function pagesToPdf(shots: readonly PdfPageShot[]): Promise<Blob> {
  const { push, offset, bytes } = pdfWriter();
  // %PDF-1.4 plus the binary-marker comment line (raw bytes — TextEncoder
  // would read the high bytes as U+00E2-style code points and re-encode them).
  push("%PDF-1.4\n");
  push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));

  const offsets: number[] = [0]; // object 0 is free
  const record = (): void => {
    offsets.push(offset());
  };

  const jpegs = await Promise.all(shots.map((s) => jpegOf(s)));
  const pt = (px: number): number => (px * 72) / 96;
  const count = jpegs.length;

  record();
  push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
  record();
  const kids = jpegs.map((_, i) => `${3 + i * 3} 0 R`).join(" ");
  push(`2 0 obj\n<< /Type /Pages /Kids [ ${kids} ] /Count ${count} >>\nendobj\n`);

  for (const [i, page] of jpegs.entries()) {
    const shot = shots[i]!;
    const pageNo = 3 + i * 3;
    const contentNo = pageNo + 1;
    const imageNo = pageNo + 2;
    // The MediaBox is the paper size (CSS px → pt); the image's own pixel
    // size only feeds /Width /Height — the content stream stretches it to
    // fill the page.
    const w = pt(shot.width).toFixed(2);
    const h = pt(shot.height).toFixed(2);
    record();
    push(
      `${pageNo} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [ 0 0 ${w} ${h} ] ` +
        `/Resources << /XObject << /Im0 ${imageNo} 0 R >> >> /Contents ${contentNo} 0 R >>\nendobj\n`,
    );
    record();
    const content = `q ${w} 0 0 ${h} 0 0 cm /Im0 Do Q`;
    push(
      `${contentNo} 0 obj\n<< /Length ${content.length} >>\nstream\n${content}\nendstream\nendobj\n`,
    );
    record();
    push(
      `${imageNo} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${page.width} /Height ${page.height} ` +
        `/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.jpeg.length} >>\nstream\n`,
    );
    push(page.jpeg);
    push("\nendstream\nendobj\n");
  }

  const xrefStart = offset();
  const size = offsets.length;
  let xref = `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let i = 1; i < size; i++) xref += `${xrefAt(offsets[i]!)} 00000 n \n`;
  xref += `trailer\n<< /Size ${size} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`;
  push(xref);

  return new Blob([bytes()], { type: "application/pdf" });
}
