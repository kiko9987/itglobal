// 플렉스 결재 첨부·본문 이미지 내려받기 — 다시 받기(flex_refresh_browser.js) 뒤 같은 탭에서 실행 (2026-11-20).
// window.__flexR.approvalDocs 중 SINCE(UTC) 뒤 쓰거나 고친 문서의 첨부 파일·본문 이미지를 zip 하나로 내려받는다.
// 11-06 전체 받기 뒤 문서는 첨부가 없어서(11-20 점검: 9건·10개) 이걸로 채웠다. 해지 날 SINCE 를 마지막 받은 날로 바꿔 다시.
// 받은 zip(브라우저 창에선 Downloads 의 .tmp 로 남을 수 있음 — 크기 확인 후 이름 바꿈)을
//   flex_export_20261106/원본/ 에 두고 첨부파일/ 아래로 풀면 기존 폴더 모양(결재첨부/번호_제목/, 결재본문이미지/)과 같다.
const SINCE = '2026-11-20';
(async () => {
  const docs = (window.__flexR && window.__flexR.approvalDocs) || [];
  const files = [];
  for (const x of docs) {
    const d = x.document || x;
    if (String(d.writtenAt || d.createdAt || '') < SINCE && String(d.updatedAt || '') < SINCE) continue;
    for (const a of d.attachments || []) files.push({code: d.code, title: d.title, kind: 'att', name: a.file.fileName, url: a.file.downloadUrl});
    for (const m of (d.content || '').matchAll(/\/api\/v2\/file\/files\/([A-Za-z0-9]+)/g))
      files.push({code: d.code, title: d.title, kind: 'img', name: m[1], url: '/api/v2/file/files/' + m[1]});
  }
  const crcT = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; crcT[n] = c >>> 0; }
  const crc = u => { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = crcT[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const enc = new TextEncoder(), parts = [], cen = [];
  let off = 0, n = 0;
  for (const f of files) {
    const r = await fetch(f.url);
    if (!r.ok) { console.warn('실패', f.code, f.name, r.status); continue; }
    const b = await r.blob(), data = new Uint8Array(await b.arrayBuffer());
    const ext = b.type === 'image/png' ? '.png' : b.type === 'image/jpeg' ? '.jpg' : '';
    const path = f.kind === 'att' ? `결재첨부/${f.code}_${f.title.replace(/[\\/:*?"<>|]/g, '_')}/${f.name}` : `결재본문이미지/${f.name}${ext}`;
    const nm = enc.encode(path), c = crc(data), sz = data.length;
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
    lh.setUint32(14, c, true); lh.setUint32(18, sz, true); lh.setUint32(22, sz, true); lh.setUint16(26, nm.length, true);
    parts.push(new Uint8Array(lh.buffer), nm, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint32(16, c, true); ch.setUint32(20, sz, true); ch.setUint32(24, sz, true); ch.setUint16(28, nm.length, true); ch.setUint32(42, off, true);
    cen.push(new Uint8Array(ch.buffer), nm);
    off += 30 + nm.length + sz; n++;
  }
  const cenSize = cen.reduce((a, b) => a + b.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, n, true); end.setUint16(10, n, true); end.setUint32(12, cenSize, true); end.setUint32(16, off, true);
  const blob = new Blob([...parts, ...cen, new Uint8Array(end.buffer)], {type: 'application/zip'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = 'flex_att_since_' + SINCE.replace(/-/g, '') + '.zip'; document.body.appendChild(a); a.click(); a.remove();
  window.__attDone = `파일 ${n}개 · ${blob.size} bytes`;
})();
