# -*- coding: utf-8 -*-
"""작은 PDF 만들기 (2026-11-20) — 증명서 한 장용. 외부 패키지 없이 표준 라이브러리만.

- 글꼴: 윈도우 맑은 고딕(TTF)을 쓰는 글자만 남겨(나머지 글리프는 비움) 문서 안에 넣는다(CIDFontType2·Identity-H).
  글리프 번호를 바꾸지 않는 '비우기' 방식이라 cmap·hmtx 를 다시 쓸 필요가 없다. 받는 쪽에 글꼴이 없어도 똑같이 보인다.
- 그림: PNG(8비트 RGB·RGBA, 비월 없음)를 풀어 RGB + 투명(SMask)으로 넣는다 — 직인용.
- 좌표는 왼쪽 위 기준 pt(A4 = 595.28 × 841.89). 글자 y 는 글자 위쪽.
"""
import struct
import zlib
from typing import Dict, List, Optional, Tuple

A4 = (595.28, 841.89)


# ── TrueType ──────────────────────────────────────────────────────────
class TTF:
    _cache: Dict[str, 'TTF'] = {}

    @classmethod
    def load(cls, path: str) -> 'TTF':
        if path not in cls._cache:
            cls._cache[path] = cls(open(path, 'rb').read())
        return cls._cache[path]

    def __init__(self, data: bytes):
        self.data = data
        if data[:4] not in (b'\x00\x01\x00\x00', b'true'):
            raise ValueError('TrueType 글꼴이 아님')
        n = struct.unpack('>H', data[4:6])[0]
        self.tables = {}
        for i in range(n):
            tag, _cs, off, ln = struct.unpack('>4sIII', data[12 + 16 * i: 28 + 16 * i])
            self.tables[tag.decode('latin-1')] = (off, ln)
        head = self._t('head')
        self.upm = struct.unpack('>H', head[18:20])[0]
        self.bbox = struct.unpack('>hhhh', head[36:44])
        self.long_loca = struct.unpack('>h', head[50:52])[0] == 1
        hhea = self._t('hhea')
        self.ascent, self.descent = struct.unpack('>hh', hhea[4:8])
        n_hm = struct.unpack('>H', hhea[34:36])[0]
        self.num_glyphs = struct.unpack('>H', self._t('maxp')[4:6])[0]
        hmtx = self._t('hmtx')
        adv = [struct.unpack('>H', hmtx[4 * i: 4 * i + 2])[0] for i in range(n_hm)]
        self.advance = adv + [adv[-1]] * (self.num_glyphs - n_hm)
        os2 = self.tables.get('OS/2')
        self.cap_height = self.ascent
        if os2 and os2[1] >= 90:
            self.cap_height = struct.unpack('>h', self._t('OS/2')[88:90])[0] or self.ascent
        self.cmap = self._read_cmap()
        loca = self._t('loca')
        if self.long_loca:
            self.loca = list(struct.unpack(f'>{self.num_glyphs + 1}I', loca[:4 * (self.num_glyphs + 1)]))
        else:
            self.loca = [x * 2 for x in struct.unpack(f'>{self.num_glyphs + 1}H', loca[:2 * (self.num_glyphs + 1)])]

    def _t(self, tag: str) -> bytes:
        off, ln = self.tables[tag]
        return self.data[off:off + ln]

    def _read_cmap(self) -> Dict[int, int]:
        cm = self._t('cmap')
        n = struct.unpack('>H', cm[2:4])[0]
        subs = {}
        for i in range(n):
            pid, eid, off = struct.unpack('>HHI', cm[4 + 8 * i: 12 + 8 * i])
            subs[(pid, eid)] = off
        out: Dict[int, int] = {}
        if (3, 10) in subs and struct.unpack('>H', cm[subs[(3, 10)]:subs[(3, 10)] + 2])[0] == 12:
            o = subs[(3, 10)]
            ngroups = struct.unpack('>I', cm[o + 12:o + 16])[0]
            for g in range(ngroups):
                s, e, gid = struct.unpack('>III', cm[o + 16 + 12 * g: o + 28 + 12 * g])
                for c in range(s, e + 1):
                    out[c] = gid + (c - s)
            return out
        o = subs.get((3, 1), subs.get((0, 3)))
        if o is None or struct.unpack('>H', cm[o:o + 2])[0] != 4:
            raise ValueError('지원하지 않는 cmap')
        seg2 = struct.unpack('>H', cm[o + 6:o + 8])[0]
        seg = seg2 // 2
        ends = struct.unpack(f'>{seg}H', cm[o + 14:o + 14 + seg2])
        starts = struct.unpack(f'>{seg}H', cm[o + 16 + seg2:o + 16 + 2 * seg2])
        deltas = struct.unpack(f'>{seg}h', cm[o + 16 + 2 * seg2:o + 16 + 3 * seg2])
        ro_pos = o + 16 + 3 * seg2
        ros = struct.unpack(f'>{seg}H', cm[ro_pos:ro_pos + seg2])
        for i in range(seg):
            for c in range(starts[i], ends[i] + 1):
                if c == 0xFFFF:
                    continue
                if ros[i] == 0:
                    g = (c + deltas[i]) & 0xFFFF
                else:
                    p = ro_pos + 2 * i + ros[i] + 2 * (c - starts[i])
                    g = struct.unpack('>H', cm[p:p + 2])[0]
                    if g:
                        g = (g + deltas[i]) & 0xFFFF
                if g:
                    out[c] = g
        return out

    def gid(self, ch: str) -> int:
        return self.cmap.get(ord(ch), 0)

    def width(self, text: str, size: float) -> float:
        return sum(self.advance[self.gid(c)] for c in text) * size / self.upm

    def _glyph(self, g: int) -> bytes:
        glyf_off = self.tables['glyf'][0]
        return self.data[glyf_off + self.loca[g]: glyf_off + self.loca[g + 1]]

    def _components(self, g: int) -> List[int]:
        b = self._glyph(g)
        if len(b) < 10 or struct.unpack('>h', b[:2])[0] >= 0:
            return []
        out, p = [], 10
        while True:
            flags, comp = struct.unpack('>HH', b[p:p + 4])
            out.append(comp)
            p += 4 + (4 if flags & 0x1 else 2)
            if flags & 0x8:
                p += 2
            elif flags & 0x40:
                p += 4
            elif flags & 0x80:
                p += 8
            if not flags & 0x20:
                return out

    def subset(self, gids) -> bytes:
        """gids 만 남긴 TTF — 나머지 글리프는 길이 0(번호는 그대로)."""
        keep = {0} | set(gids)
        todo = list(keep)
        while todo:
            for c in self._components(todo.pop()):
                if c not in keep:
                    keep.add(c)
                    todo.append(c)
        glyf, loca = bytearray(), []
        for g in range(self.num_glyphs):
            loca.append(len(glyf))
            if g in keep:
                glyf += self._glyph(g)
                glyf += b'\x00' * (-len(glyf) % 4)
        loca.append(len(glyf))
        head = bytearray(self._t('head'))
        head[8:12] = b'\x00\x00\x00\x00'                      # checkSumAdjustment — 아래에서 다시
        head[50:52] = struct.pack('>h', 1)                     # long loca
        post = struct.pack('>IIhhIIIII', 0x00030000, 0, 0, 0, 0, 0, 0, 0, 0)
        tables = {'head': bytes(head), 'glyf': bytes(glyf), 'loca': struct.pack(f'>{len(loca)}I', *loca),
                  'post': post}
        for t in ('hhea', 'hmtx', 'maxp', 'cvt ', 'fpgm', 'prep', 'OS/2', 'name', 'cmap'):   # cmap: 맥·폰 보기 프로그램 대비
            if t in self.tables:
                tables[t] = self._t(t)
        return _build_sfnt(tables)


def _checksum(b: bytes) -> int:
    b = b + b'\x00' * (-len(b) % 4)
    return sum(struct.unpack(f'>{len(b) // 4}I', b)) & 0xFFFFFFFF


def _build_sfnt(tables: Dict[str, bytes]) -> bytes:
    tags = sorted(tables)
    n = len(tags)
    es = max(i for i in range(16) if 2 ** i <= n)
    hdr = struct.pack('>IHHHH', 0x00010000, n, 16 * 2 ** es, es, n * 16 - 16 * 2 ** es)
    off = 12 + 16 * n
    dirs, body = b'', b''
    for t in tags:
        data = tables[t]
        dirs += struct.pack('>4sIII', t.encode('latin-1'), _checksum(data), off + len(body), len(data))
        body += data + b'\x00' * (-len(data) % 4)
    font = bytearray(hdr + dirs + body)
    adj = (0xB1B0AFBA - _checksum(bytes(font))) & 0xFFFFFFFF
    head_off = 12 + 16 * n + sum(len(tables[t]) + (-len(tables[t]) % 4) for t in tags[:tags.index('head')])
    font[head_off + 8: head_off + 12] = struct.pack('>I', adj)
    return bytes(font)


# ── PNG ───────────────────────────────────────────────────────────────
def decode_png(data: bytes) -> Tuple[int, int, bytes, Optional[bytes]]:
    """(가로, 세로, RGB 바이트, 투명 바이트 또는 None). 8비트 RGB(2)·RGBA(6)·회색(0)·회색+투명(4), 비월 없음만."""
    if data[:8] != b'\x89PNG\r\n\x1a\n':
        raise ValueError('PNG 아님')
    p, idat, w = 8, b'', 0
    while p < len(data):
        ln = struct.unpack('>I', data[p:p + 4])[0]
        kind, body = data[p + 4:p + 8], data[p + 8:p + 8 + ln]
        if kind == b'IHDR':
            w, h, depth, ctype, _c, _f, inter = struct.unpack('>IIBBBBB', body)
            if depth != 8 or inter or ctype not in (0, 2, 4, 6):
                raise ValueError('지원하지 않는 PNG(8비트 비월 없음만)')
        elif kind == b'IDAT':
            idat += body
        elif kind == b'IEND':
            break
        p += 12 + ln
    bpp = {0: 1, 2: 3, 4: 2, 6: 4}[ctype]
    raw = zlib.decompress(idat)
    stride = w * bpp
    out = bytearray(h * stride)
    prev = bytearray(stride)
    pos = 0
    for y in range(h):
        f = raw[pos]
        line = bytearray(raw[pos + 1:pos + 1 + stride])
        pos += 1 + stride
        if f == 1:
            for i in range(bpp, stride):
                line[i] = (line[i] + line[i - bpp]) & 255
        elif f == 2:
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 255
        elif f == 3:
            for i in range(stride):
                left = line[i - bpp] if i >= bpp else 0
                line[i] = (line[i] + ((left + prev[i]) >> 1)) & 255
        elif f == 4:
            for i in range(stride):
                a = line[i - bpp] if i >= bpp else 0
                b = prev[i]
                c = prev[i - bpp] if i >= bpp else 0
                pa, pb, pc = abs(b - c), abs(a - c), abs(a + b - 2 * c)
                pr = a if pa <= pb and pa <= pc else (b if pb <= pc else c)
                line[i] = (line[i] + pr) & 255
        out[y * stride:(y + 1) * stride] = line
        prev = line
    if ctype == 2:
        return w, h, bytes(out), None
    if ctype == 0:
        return w, h, bytes(b for v in out for b in (v, v, v)), None
    if ctype == 4:
        g, a = out[0::2], out[1::2]
        return w, h, bytes(b for v in g for b in (v, v, v)), bytes(a)
    rgb = bytearray(w * h * 3)
    rgb[0::3], rgb[1::3], rgb[2::3] = out[0::4], out[1::4], out[2::4]
    return w, h, bytes(rgb), bytes(out[3::4])


# ── PDF ───────────────────────────────────────────────────────────────
class Doc:
    """한 쪽짜리 문서. font(key, ttf) 로 글꼴을 등록하고 text/line/rect/image 로 그린 뒤 output()."""

    def __init__(self, size=A4, title: str = ''):
        self.W, self.H = size
        self.title = title
        self.ops: List[str] = []
        self.fonts: Dict[str, TTF] = {}
        self.used: Dict[str, Dict[int, str]] = {}
        self.images: List[Tuple[int, int, bytes, Optional[bytes]]] = []

    def font(self, key: str, ttf: TTF):
        self.fonts[key] = ttf
        self.used[key] = {}

    @staticmethod
    def _c(rgb) -> str:
        return ' '.join(f'{v / 255:.3f}' for v in rgb)

    def text_width(self, s: str, size: float, font: str) -> float:
        return self.fonts[font].width(s, size)

    def text(self, x: float, y: float, s: str, size: float = 10, font: str = 'R', color=(0, 0, 0),
             align: str = 'left', spacing: float = 0):
        f = self.fonts[font]
        w = f.width(s, size) + spacing * max(len(s) - 1, 0)
        if align == 'center':
            x -= w / 2
        elif align == 'right':
            x -= w
        base = self.H - y - f.ascent * size / f.upm               # 글자 위쪽 → 기준선
        hexs = ''
        for ch in s:
            g = f.gid(ch)
            self.used[font][g] = ch
            hexs += f'{g:04X}'
        # Tc(글자 간격)는 BT/ET 밖까지 남는 그래픽 상태 — 매번 적어야 앞 글자의 간격이 번지지 않는다
        self.ops.append(f'BT /{font} {size:.2f} Tf {spacing:.2f} Tc {self._c(color)} rg {x:.2f} {base:.2f} Td <{hexs}> Tj ET')
        return w

    def line(self, x1, y1, x2, y2, width: float = 0.5, color=(0, 0, 0)):
        self.ops.append(f'{width:.2f} w {self._c(color)} RG {x1:.2f} {self.H - y1:.2f} m {x2:.2f} {self.H - y2:.2f} l S')

    def rect(self, x, y, w, h, fill=(240, 240, 240)):
        self.ops.append(f'{self._c(fill)} rg {x:.2f} {self.H - y - h:.2f} {w:.2f} {h:.2f} re f')

    def image(self, png: bytes, x, y, w, h):
        self.images.append(decode_png(png) if isinstance(png, (bytes, bytearray)) else png)
        n = len(self.images)
        self.ops.append(f'q {w:.2f} 0 0 {h:.2f} {x:.2f} {self.H - y - h:.2f} cm /Im{n} Do Q')

    def output(self) -> bytes:
        objs: List[bytes] = []

        def add(b) -> int:
            objs.append(b if isinstance(b, bytes) else b.encode('latin-1'))
            return len(objs)

        def stream(dict_body: str, data: bytes, compress: bool = True) -> int:
            if compress:
                data = zlib.compress(data, 9)
                dict_body += ' /Filter /FlateDecode'
            return add(f'<< {dict_body} /Length {len(data)} >>\nstream\n'.encode('latin-1') + data + b'\nendstream')

        font_refs = {}
        for key, f in self.fonts.items():
            used = self.used[key]
            if not used:
                continue
            sub = f.subset(used.keys())
            ff = stream(f'/Length1 {len(sub)}', sub)
            sc = 1000 / f.upm
            desc = add(f'<< /Type /FontDescriptor /FontName /ITG{key}+Malgun /Flags 32 '
                       f'/FontBBox [{" ".join(str(int(v * sc)) for v in f.bbox)}] /ItalicAngle 0 '
                       f'/Ascent {int(f.ascent * sc)} /Descent {int(f.descent * sc)} /CapHeight {int(f.cap_height * sc)} '
                       f'/StemV 80 /FontFile2 {ff} 0 R >>')
            ws = ' '.join(f'{g} [{int(f.advance[g] * sc)}]' for g in sorted(used))
            cid = add(f'<< /Type /Font /Subtype /CIDFontType2 /BaseFont /ITG{key}+Malgun '
                      f'/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> '
                      f'/FontDescriptor {desc} 0 R /W [{ws}] /CIDToGIDMap /Identity >>')
            cmap = ('/CIDInit /ProcSet findresource begin 12 dict begin begincmap '
                    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def '
                    '/CMapName /Adobe-Identity-UCS def /CMapType 2 def 1 begincodespacerange <0000> <FFFF> endcodespacerange\n')
            items = sorted(used.items())
            for i in range(0, len(items), 100):
                chunk = items[i:i + 100]
                cmap += f'{len(chunk)} beginbfchar\n' + ''.join(
                    f'<{g:04X}> <{"".join(f"{u:04X}" for u in _utf16(ch))}>\n' for g, ch in chunk) + 'endbfchar\n'
            cmap += 'endcmap CMapName currentdict /CMap defineresource pop end end'
            tu = stream('', cmap.encode('latin-1'))
            font_refs[key] = add(f'<< /Type /Font /Subtype /Type0 /BaseFont /ITG{key}+Malgun /Encoding /Identity-H '
                                 f'/DescendantFonts [{cid} 0 R] /ToUnicode {tu} 0 R >>')
        img_refs = []
        for (w, h, rgb, alpha) in self.images:
            smask = ''
            if alpha is not None:
                sm = stream(f'/Type /XObject /Subtype /Image /Width {w} /Height {h} /ColorSpace /DeviceGray '
                            f'/BitsPerComponent 8', alpha)
                smask = f' /SMask {sm} 0 R'
            img_refs.append(stream(f'/Type /XObject /Subtype /Image /Width {w} /Height {h} /ColorSpace /DeviceRGB '
                                   f'/BitsPerComponent 8{smask}', rgb))
        content = stream('', '\n'.join(self.ops).encode('latin-1'))
        res = ('<< /Font << ' + ' '.join(f'/{k} {v} 0 R' for k, v in font_refs.items()) + ' >> /XObject << '
               + ' '.join(f'/Im{i + 1} {r} 0 R' for i, r in enumerate(img_refs)) + ' >> >>')
        pages_id = len(objs) + 2
        page = add(f'<< /Type /Page /Parent {pages_id} 0 R /MediaBox [0 0 {self.W:.2f} {self.H:.2f}] '
                   f'/Resources {res} /Contents {content} 0 R >>')
        add(f'<< /Type /Pages /Kids [{page} 0 R] /Count 1 >>')
        info = add('<< /Producer (ITG groupware) /Title <' + 'FEFF' + ''.join(
            f'{u:04X}' for ch in self.title for u in _utf16(ch)) + '> >>')
        catalog = add(f'<< /Type /Catalog /Pages {pages_id} 0 R >>')
        out = bytearray(b'%PDF-1.7\n%\xe2\xe3\xcf\xd3\n')
        offs = []
        for i, o in enumerate(objs, 1):
            offs.append(len(out))
            out += f'{i} 0 obj\n'.encode('latin-1') + o + b'\nendobj\n'
        xref = len(out)
        out += f'xref\n0 {len(objs) + 1}\n0000000000 65535 f \n'.encode('latin-1')
        out += b''.join(f'{o:010d} 00000 n \n'.encode('latin-1') for o in offs)
        out += (f'trailer\n<< /Size {len(objs) + 1} /Root {catalog} 0 R /Info {info} 0 R >>\n'
                f'startxref\n{xref}\n%%EOF\n').encode('latin-1')
        return bytes(out)


def _utf16(ch: str) -> List[int]:
    b = ch.encode('utf-16-be')
    return [struct.unpack('>H', b[i:i + 2])[0] for i in range(0, len(b), 2)]
