"""Neutralise the text of a .docx WIDTH FOR WIDTH.

Every letter and digit is replaced by another of the SAME ADVANCE WIDTH in
every face the document can be set in, chosen at random per occurrence (not a
cipher: the same letter goes to a different letter each time). Spaces, digits'
class, punctuation and everything outside the run text stay exactly as they
were. So every word keeps its width in every face, every line breaks where it
broke, and a baseline printed from the original stays valid for the result --
which is the check: reprint the result and it must reproduce the baseline.

    python3 tools/neutralise-width-for-width.py original.docx neutral.docx \
        --fonts assets/fonts --face LiberationSerif-Regular.ttf \
        --face LiberationSerif-Bold.ttf --face LiberationSans-Bold.ttf \
        --face LiberationMono-Regular.ttf [--seed 1]

Only Latin letters and ASCII digits are touched, and only within the ONE face
a run is set in (see the rules below); a character with no partner of its own
width in that face is kept as it is. XML entities are left alone.
"""
import argparse
import random
import re
import struct
import sys
import zipfile
from pathlib import Path

LETTERS = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
DIGITS = '0123456789'


def advances(ttf: Path) -> dict[str, int]:
    """Advance width per character of interest, in font units, from cmap + hmtx."""
    data = ttf.read_bytes()
    num_tables = struct.unpack('>H', data[4:6])[0]
    tables = {}
    for i in range(num_tables):
        tag, _, off, length = struct.unpack('>4sIII', data[12 + 16 * i:28 + 16 * i])
        tables[tag.decode('latin1')] = (off, length)
    # cmap: a (3,1) or (0,x) format-4 subtable
    coff, _ = tables['cmap']
    n = struct.unpack('>H', data[coff + 2:coff + 4])[0]
    sub = None
    for i in range(n):
        pid, eid, so = struct.unpack('>HHI', data[coff + 4 + 8 * i:coff + 12 + 8 * i])
        fmt = struct.unpack('>H', data[coff + so:coff + so + 2])[0]
        if fmt == 4 and (pid == 3 and eid == 1 or pid == 0):
            sub = coff + so
            break
    assert sub is not None, f'{ttf.name}: no format-4 cmap'
    segx2 = struct.unpack('>H', data[sub + 6:sub + 8])[0]
    seg = segx2 // 2
    ends = struct.unpack(f'>{seg}H', data[sub + 14:sub + 14 + segx2])
    starts = struct.unpack(f'>{seg}H', data[sub + 16 + segx2:sub + 16 + 2 * segx2])
    deltas = struct.unpack(f'>{seg}h', data[sub + 16 + 2 * segx2:sub + 16 + 3 * segx2])
    range_off_pos = sub + 16 + 3 * segx2
    range_offs = struct.unpack(f'>{seg}H', data[range_off_pos:range_off_pos + segx2])

    def glyph(cp: int) -> int:
        for i in range(seg):
            if starts[i] <= cp <= ends[i]:
                if range_offs[i] == 0:
                    return (cp + deltas[i]) & 0xFFFF
                addr = range_off_pos + 2 * i + range_offs[i] + 2 * (cp - starts[i])
                g = struct.unpack('>H', data[addr:addr + 2])[0]
                return (g + deltas[i]) & 0xFFFF if g else 0
        return 0

    hoff, _ = tables['hhea']
    num_h = struct.unpack('>H', data[hoff + 34:hoff + 36])[0]
    moff, _ = tables['hmtx']

    def advance(g: int) -> int:
        i = min(g, num_h - 1)
        return struct.unpack('>H', data[moff + 4 * i:moff + 4 * i + 2])[0]

    return {c: advance(glyph(ord(c))) for c in LETTERS + DIGITS}


def classes(faces: list[dict[str, int]], alphabet: str) -> dict[str, list[str]]:
    """Characters grouped by their advance in EVERY face."""
    by_key: dict[tuple[int, ...], list[str]] = {}
    for c in alphabet:
        by_key.setdefault(tuple(f[c] for f in faces), []).append(c)
    return {c: members for members in by_key.values() for c in members}




FACE_RULES = '''
A run is set in ONE face, and only that face's widths constrain its letters:
  mono        the run names a monospace family (Courier New)
  sans-bold   the paragraph is a heading (Word's Heading1/Heading2 without fonts
              of their own: LibreOffice sets those in its heading face)
  serif-bold  the run is bold
  serif       everything else
Each face is a --face NAME=FILE.ttf; the four names above are the rule set.
'''


def face_of(paragraph_props: str, run_props: str) -> str:
    fonts = re.search(r'w:ascii="([^"]+)"', run_props)
    if fonts and re.search(r'courier|mono', fonts.group(1), re.I):
        return 'mono'
    if re.search(r'<w:pStyle w:val="Heading\d"', paragraph_props):
        return 'sans-bold'
    if re.search(r'<w:b(?: [^>]*)?/>', run_props) and not re.search(r'<w:b w:val="(?:0|false)"', run_props):
        return 'serif-bold'
    return 'serif'


ENTITY = re.compile(r'&[a-zA-Z#0-9]+;')


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__ + FACE_RULES, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('source')
    ap.add_argument('target')
    ap.add_argument('--fonts', required=True)
    ap.add_argument('--face', action='append', required=True, help='NAME=FILE.ttf, e.g. serif=LiberationSerif-Regular.ttf')
    ap.add_argument('--seed', type=int, default=1)
    args = ap.parse_args()
    faces = {}
    for spec in args.face:
        name, file = spec.split('=', 1)
        faces[name] = advances(Path(args.fonts) / file)
    for needed in ('serif', 'serif-bold', 'sans-bold', 'mono'):
        assert needed in faces, f'--face {needed}=... is missing'
    tables = {name: {**classes([adv], 'abcdefghijklmnopqrstuvwxyz'), **classes([adv], 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), **classes([adv], DIGITS)} for name, adv in faces.items()}
    rng = random.Random(args.seed)
    stats = {'replaced': 0, 'kept': 0, 'runs': 0, 'faces': {}}

    def neutralise_text(text: str, table: dict[str, list[str]]) -> str:
        out = []
        i = 0
        while i < len(text):
            m = ENTITY.match(text, i)
            if m:
                out.append(m.group(0))
                i = m.end()
                continue
            c = text[i]
            members = table.get(c)
            if members is None or len(members) < 2:
                if members is not None:
                    stats['kept'] += 1
                out.append(c)
            else:
                out.append(rng.choice([x for x in members if x != c]))
                stats['replaced'] += 1
            i += 1
        return ''.join(out)

    def neutralise_paragraph(m: re.Match) -> str:
        p = m.group(0)
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S)
        ppr_s = ppr.group(0) if ppr else ''

        def run(rm: re.Match) -> str:
            r = rm.group(0)
            rpr = re.search(r'<w:rPr>.*?</w:rPr>', r, re.S)
            face = face_of(ppr_s, rpr.group(0) if rpr else '')
            stats['runs'] += 1
            stats['faces'][face] = stats['faces'].get(face, 0) + 1
            table = tables[face]
            return re.sub(r'(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)', lambda tm: tm.group(1) + neutralise_text(tm.group(2), table) + tm.group(3), r)

        return re.sub(r'<w:r(?: [^>]*)?>.*?</w:r>', run, p, flags=re.S)

    with zipfile.ZipFile(args.source) as src, zipfile.ZipFile(args.target, 'w') as dst:
        for item in src.infolist():
            body = src.read(item.filename)
            if re.match(r'word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$', item.filename):
                xml = body.decode('utf-8')
                xml = re.sub(r'<w:p(?: [^>]*)?>.*?</w:p>', neutralise_paragraph, xml, flags=re.S)
                body = xml.encode('utf-8')
            dst.writestr(item, body, compress_type=item.compress_type)
    singles = {name: ''.join(sorted(c for c, m in t.items() if len(m) < 2)) for name, t in tables.items()}
    print(f"runs {stats['runs']} by face {stats['faces']}; characters replaced {stats['replaced']}, kept {stats['kept']}; singleton classes per face {singles}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
