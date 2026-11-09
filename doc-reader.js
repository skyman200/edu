/* 문서 읽기: PDF, Word, PPT, 한글(hwpx), 엑셀, 텍스트, 이미지에서 글을 뽑는 브라우저 도구. 파일은 올리지 않고 글만 꺼냅니다.
   실습실(index.html)과 24시간 조교(ta.html)가 함께 씁니다. 의존: vendor/pdfjs (PDF, 처음 읽을 때 불러옴). */
(function (root) {
'use strict';
var BASE = (function () { try { return new URL('.', document.currentScript.src).href; } catch (e) { return './'; } })();
var ATT = { files: 3, images: 4, textOne: 100000, textAll: 150000, bytes: 25 * 1024 * 1024, imgSide: 1568 };
const extOf = n => (n.split('.').pop() || '').toLowerCase();
const decodeXml = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, '&');
const tidy = s => s.replace(/\r\n?/g, '\n').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
async function unzip(buf) {
  const dv = new DataView(buf), u8 = new Uint8Array(buf), ent = {};
  let e = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
  if (e < 0) throw new Error('not a zip');
  const n = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true);
  for (let k = 0; k < n && dv.getUint32(p, true) === 0x02014b50; k++) {
    const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    ent[new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nl))] = { method: dv.getUint16(p + 10, true), size: dv.getUint32(p + 20, true), off: dv.getUint32(p + 42, true) };
    p += 46 + nl + xl + cl;
  }
  return {
    names: Object.keys(ent),
    async read(name) {
      const f = ent[name]; if (!f) return '';
      const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true), data = u8.subarray(start, start + f.size);
      if (f.method === 0) return new TextDecoder().decode(data);
      return new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
    }
  };
}
const numSort = (a, b) => (parseInt(a.replace(/\D+/g, ''), 10) || 0) - (parseInt(b.replace(/\D+/g, ''), 10) || 0);
const paras = (xml, endTag, textRe) => xml.split(endTag).map(p => { let s = ''; for (const m of p.matchAll(textRe)) s += decodeXml(m[1].replace(/<[^>]+>/g, '')); return s.trim(); }).filter(Boolean).join('\n');
const readers = {
  async docx(buf) { const z = await unzip(buf); return paras(await z.read('word/document.xml'), '</w:p>', /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g); },
  async pptx(buf) {
    const z = await unzip(buf), out = [];
    for (const nm of z.names.filter(x => /^ppt\/slides\/slide\d+\.xml$/.test(x)).sort(numSort)) {
      const t = paras(await z.read(nm), '</a:p>', /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g);
      if (t) out.push('[슬라이드 ' + (parseInt(nm.replace(/\D+/g, ''), 10)) + ']\n' + t);
    }
    return out.join('\n\n');
  },
  async hwpx(buf) {
    const z = await unzip(buf), out = [];
    for (const nm of z.names.filter(x => /^Contents\/section\d+\.xml$/i.test(x)).sort(numSort)) out.push(paras(await z.read(nm), '</hp:p>', /<hp:t(?:\s[^>]*)?>([\s\S]*?)<\/hp:t>/g));
    return out.join('\n');
  },
  async xlsx(buf) {
    const z = await unzip(buf), ss = [];
    for (const m of (await z.read('xl/sharedStrings.xml')).matchAll(/<si>([\s\S]*?)<\/si>/g)) ss.push(decodeXml([...m[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(x => x[1]).join('')));
    const out = [];
    for (const nm of z.names.filter(x => /^xl\/worksheets\/sheet\d+\.xml$/.test(x)).sort(numSort)) {
      const rows = [];
      for (const r of (await z.read(nm)).matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)) {
        const cells = [];
        for (const c of r[1].matchAll(/<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
          const col = ((c[1].match(/r="([A-Z]+)\d+"/) || [])[1] || 'A').split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
          const v = (c[2] || '').match(/<v>([\s\S]*?)<\/v>/), is = (c[2] || '').match(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/);
          cells[col] = /t="s"/.test(c[1]) && v ? ss[+v[1]] || '' : is ? decodeXml(is[1]) : v ? decodeXml(v[1]) : '';
        }
        if (cells.some(Boolean)) rows.push(Array.from(cells, x => x || '').join('\t'));
      }
      if (rows.length) out.push('[시트 ' + (parseInt(nm.replace(/\D+/g, ''), 10)) + ']\n' + rows.join('\n'));
    }
    return out.join('\n\n');
  }
};
async function readTextFile(file) {
  const buf = await file.arrayBuffer();
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { return new TextDecoder('euc-kr').decode(buf); }
}
let pdfLib;
async function loadPdf() {
  if (!pdfLib) {
    pdfLib = await import(BASE + 'vendor/pdfjs/pdf.min.mjs');
    pdfLib.GlobalWorkerOptions.workerSrc = new URL(BASE + 'vendor/pdfjs/pdf.worker.min.mjs', location.href).href;
  }
  return pdfLib;
}
function canvasUrl(src, w, h) {
  const k = Math.min(1, ATT.imgSide / Math.max(w, h)), c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(w * k)); c.height = Math.max(1, Math.round(h * k));
  const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(src, 0, 0, c.width, c.height);
  return c.toDataURL('image/jpeg', 0.82);
}
async function readPdf(file) {
  const lib = await loadPdf();
  const pdf = await lib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  let text = '', chars = 0;
  for (let i = 1; i <= pdf.numPages && text.length < ATT.textOne * 1.3; i++) {
    const tc = await (await pdf.getPage(i)).getTextContent();
    const s = tc.items.map(it => it.str + (it.hasEOL ? '\n' : ' ')).join('').replace(/[ \t]+/g, ' ').replace(/ ?\n ?/g, '\n').trim();
    chars += s.length;
    text += '[쪽 ' + i + ']\n' + s + '\n\n';
  }
  if (chars >= 40 * Math.min(pdf.numPages, 3)) return { text: text.trim(), pages: pdf.numPages };
  const images = [], n = Math.min(pdf.numPages, ATT.images);  // a scan has no text layer: send the first pages as pictures
  for (let i = 1; i <= n; i++) {
    const pg = await pdf.getPage(i), vp0 = pg.getViewport({ scale: 1 }), sc = ATT.imgSide / Math.max(vp0.width, vp0.height), vp = pg.getViewport({ scale: sc });
    const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    await pg.render({ canvasContext: c.getContext('2d'), viewport: vp, background: '#fff' }).promise;
    images.push(c.toDataURL('image/jpeg', 0.82));
  }
  return { images, pages: pdf.numPages };
}
async function readImage(file) {
  let bmp;
  try { bmp = await createImageBitmap(file); } catch (e) { throw new Error('이 이미지 형식은 읽을 수 없습니다. PNG나 JPG로 저장해 올려 주세요. (아이폰 HEIC는 안 됩니다)'); }
  const url = canvasUrl(bmp, bmp.width, bmp.height); if (bmp.close) bmp.close();
  return url;
}

// 파일 하나를 읽어 {text, note} 를 돌려줍니다. 글자가 없는 스캔 PDF와 이미지는 {images:[dataURL], note} 입니다.
async function read(file) {
  const ext = extOf(file.name), isImg = /^image\//.test(file.type) || /^(png|jpe?g|webp)$/.test(ext);
  if (file.size > ATT.bytes) throw new Error('파일이 너무 큽니다. 25MB 이하로 줄여 주세요.');
  if (ext === 'hwp') throw new Error('예전 한글(.hwp)은 읽을 수 없습니다. 한글에서 "다른 이름으로 저장"으로 .hwpx나 PDF로 저장해 올려 주세요.');
  if (/^(ppt|doc|xls)$/.test(ext)) throw new Error('예전 Office 형식(.' + ext + ')은 읽을 수 없습니다. .' + ext + 'x나 PDF로 저장해 올려 주세요.');
  if (isImg) return { images: [await readImage(file)], note: '이미지 1장' };
  let text = '', note = '';
  if (ext === 'pdf') {
    const r = await readPdf(file);
    if (r.images) return { images: r.images, note: '글자 없는 스캔 PDF라 앞 ' + r.images.length + '쪽을 그림으로 읽습니다' };
    text = r.text; note = r.pages + '쪽';
  } else if (readers[ext]) text = tidy(await readers[ext](await file.arrayBuffer()));
  else if (/^(txt|md|csv|tsv|json|html?)$/.test(ext) || file.type.startsWith('text/')) text = tidy(await readTextFile(file));
  else throw new Error('이 형식(.' + ext + ')은 읽을 수 없습니다. PDF, Word, PPT, 한글(hwpx), 엑셀, 텍스트, 이미지를 올려 주세요.');
  text = text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  if (!text.trim()) throw new Error('읽을 수 있는 글자가 없습니다. 스캔한 문서라면 PDF나 이미지(JPG, PNG)로 올려 주세요.');
  return { text: text, note: note };
}
root.DocReader = { ATT: ATT, extOf: extOf, decodeXml: decodeXml, tidy: tidy, readers: readers, readTextFile: readTextFile, readPdf: readPdf, readImage: readImage, read: read };
})(window);
