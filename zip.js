/* global window */
// ============================================================
// LexRef — lecture / écriture d'archives ZIP (un .docx en est une), sans dépendance.
// Décompression et compression par les flux natifs du navigateur (« deflate-raw » :
// WebView2 / Chromium, Safari 16.4+). Expose window.WPZip.
//   read(bytes)    → Promise<[{ name, data: Uint8Array }]>, dans l'ordre de l'archive
//   write(entries) → Promise<Uint8Array>, mêmes entrées, même ordre (Word tient à ce que
//                    [Content_Types].xml reste en tête)
// Pas de ZIP64 : un acte Word ne s'en approche pas (4 Go).
// ============================================================
(function (root) {
  const SIG_LOCAL = 0x04034b50, SIG_CENTRAL = 0x02014b50, SIG_END = 0x06054b50;

  async function pipe(bytes, stream) {
    const out = new Response(new Blob([bytes]).stream().pipeThrough(stream));
    return new Uint8Array(await out.arrayBuffer());
  }
  const inflate = (b) => pipe(b, new DecompressionStream("deflate-raw"));
  const deflate = (b) => pipe(b, new CompressionStream("deflate-raw"));
  const canDeflate = () => typeof CompressionStream === "function";

  let crcTable = null;
  function crc32(bytes) {
    if (!crcTable) {
      crcTable = new Uint32Array(256);
      for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        crcTable[n] = c >>> 0;
      }
    }
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = crcTable[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  // Les tailles et positions font foi dans le RÉPERTOIRE CENTRAL (les en-têtes locaux peuvent
  // les laisser à zéro quand un descripteur de données suit le contenu).
  async function read(input) {
    const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let end = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 65535); i--) {
      if (dv.getUint32(i, true) === SIG_END) { end = i; break; }
    }
    if (end < 0) throw new Error("Fichier illisible : ce n'est pas un document Word (.docx).");
    const count = dv.getUint16(end + 10, true);
    let p = dv.getUint32(end + 16, true);
    const utf8 = new TextDecoder("utf-8");
    const entries = [];
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== SIG_CENTRAL) throw new Error("Archive endommagée.");
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
      const local = dv.getUint32(p + 42, true);
      const name = utf8.decode(bytes.subarray(p + 46, p + 46 + nameLen));
      p += 46 + nameLen + extraLen + commentLen;
      if (dv.getUint32(local, true) !== SIG_LOCAL) throw new Error("Archive endommagée.");
      const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      const raw = bytes.subarray(start, start + csize);
      let data;
      if (method === 0) data = raw.slice();
      else if (method === 8) data = await inflate(raw);
      else throw new Error("Compression non prise en charge (méthode " + method + ").");
      entries.push({ name, data });
    }
    return entries;
  }

  async function write(entries) {
    const enc = new TextEncoder();
    const chunks = [], central = [];
    let offset = 0;
    for (const e of entries) {
      const name = enc.encode(e.name);
      const data = e.data instanceof Uint8Array ? e.data : enc.encode(String(e.data));
      const crc = crc32(data);
      let method = 0, body = data;
      if (canDeflate() && data.length > 64) {
        const z = await deflate(data);
        if (z.length < data.length) { method = 8; body = z; }
      }
      const head = new DataView(new ArrayBuffer(30));
      head.setUint32(0, SIG_LOCAL, true);
      head.setUint16(4, 20, true);       // version requise
      head.setUint16(6, 0x0800, true);   // noms en UTF-8
      head.setUint16(8, method, true);
      head.setUint16(10, 0, true); head.setUint16(12, 0x21, true); // heure / date DOS (1980-01-01)
      head.setUint32(14, crc, true);
      head.setUint32(18, body.length, true);
      head.setUint32(22, data.length, true);
      head.setUint16(26, name.length, true);
      head.setUint16(28, 0, true);
      chunks.push(new Uint8Array(head.buffer), name, body);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, SIG_CENTRAL, true);
      cd.setUint16(4, 20, true); cd.setUint16(6, 20, true);
      cd.setUint16(8, 0x0800, true);
      cd.setUint16(10, method, true);
      cd.setUint16(12, 0, true); cd.setUint16(14, 0x21, true);
      cd.setUint32(16, crc, true);
      cd.setUint32(20, body.length, true);
      cd.setUint32(24, data.length, true);
      cd.setUint16(28, name.length, true);
      cd.setUint32(42, offset, true);
      central.push(new Uint8Array(cd.buffer), name);
      offset += 30 + name.length + body.length;
    }
    const cdSize = central.reduce((n, c) => n + c.length, 0);
    const endRec = new DataView(new ArrayBuffer(22));
    endRec.setUint32(0, SIG_END, true);
    endRec.setUint16(8, entries.length, true);
    endRec.setUint16(10, entries.length, true);
    endRec.setUint32(12, cdSize, true);
    endRec.setUint32(16, offset, true);
    const all = [...chunks, ...central, new Uint8Array(endRec.buffer)];
    const out = new Uint8Array(all.reduce((n, c) => n + c.length, 0));
    let o = 0;
    for (const c of all) { out.set(c, o); o += c.length; }
    return out;
  }

  // base64 ↔ octets (createDocument attend du base64 ; getFileAsync rend des tableaux d'octets).
  function toBase64(bytes) {
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  root.WPZip = { read, write, crc32, toBase64 };
})(typeof window !== "undefined" ? window : globalThis);
