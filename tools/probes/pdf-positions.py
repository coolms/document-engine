"""Per-page text WITH its x/y, so alignment can be checked rather than assumed.

Text position in a PDF comes from the text matrix (`a b c d e f Tm`), whose e/f
are the x/y of the run that follows. Subset fonts mean the glyph codes are
meaningless without each font's ToUnicode CMap, so that is decoded too.

Coordinates are PDF points from the BOTTOM-LEFT of the page.
"""
import re
import sys
import zlib

data = open(sys.argv[1], "rb").read()

objects = {}
for match in re.finditer(rb"(\d+)\s+(\d+)\s+obj(.*?)endobj", data, re.S):
    objects[int(match.group(1))] = match.group(3)


def inflate(body):
    match = re.search(rb"stream\r?\n", body)
    if match is None:
        return b""
    raw = body[match.end():body.rindex(b"endstream")]
    try:
        return zlib.decompress(raw)
    except zlib.error:
        return raw


unicode_of = {}
for body in objects.values():
    blob = inflate(body)
    for section in re.findall(rb"beginbfchar(.*?)endbfchar", blob, re.S):
        for code, target in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", section):
            unicode_of.setdefault(int(code, 16), chr(int(target[:4], 16)))
    for section in re.findall(rb"beginbfrange(.*?)endbfrange", blob, re.S):
        for low, high, target in re.findall(
                rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", section):
            base = int(target[:4], 16)
            for offset in range(int(high, 16) - int(low, 16) + 1):
                unicode_of.setdefault(int(low, 16) + offset, chr(base + offset))


def kids(number):
    body = objects.get(number, b"")
    if b"/Pages" not in body:
        return [number]
    listed = re.search(rb"/Kids\s*\[(.*?)\]", body, re.S)
    out = []
    for kid in re.findall(rb"(\d+)\s+\d+\s+R", listed.group(1) if listed else b""):
        out.extend(kids(int(kid)))
    return out


order = []
for root in [n for n, b in objects.items()
             if b"/Pages" in b and b"/Parent" not in b and b"/Kids" in b]:
    order.extend(kids(root))

NUM = rb"(-?[\d.]+)"
# `Tm` sets the matrix outright; `Td` displaces from the line start, which BT
# has just reset to the origin — so inside one BT block it is absolute too.
# LibreOffice emits the Td form, one BT block per line.
MATRIX = re.compile(rb"\s+".join([NUM] * 6) + rb"\s+Tm|" + NUM + rb"\s+" + NUM + rb"\s+Td")
SHOW = re.compile(rb"\[((?:[^\[\]\\]|\\.)*)\]\s*TJ|<([0-9A-Fa-f]*)>\s*Tj")


def decode(source):
    out = []
    for hexed in re.findall(rb"<([0-9A-Fa-f]*)>", source):
        for index in range(0, len(hexed) - 1, 2):
            out.append(unicode_of.get(int(hexed[index:index + 2], 16), "?"))
    return "".join(out)


for index, page in enumerate(order):
    body = objects[page]
    content = re.search(rb"/Contents\s+(\d+)\s+\d+\s+R", body)
    blob = b"" if content is None else inflate(objects[int(content.group(1))])
    print("--- PAGE %d ---" % (index + 1))

    # Walk the stream in order, remembering the most recent text matrix.
    x = y = 0.0
    for token in re.finditer(MATRIX.pattern + rb"|" + SHOW.pattern, blob):
        groups = token.groups()
        if groups[0] is not None:            # a b c d e f Tm
            x, y = float(groups[4]), float(groups[5])
            continue
        if groups[6] is not None:            # x y Td
            x, y = float(groups[6]), float(groups[7])
            continue
        source = groups[8] if groups[8] is not None else b"<" + (groups[9] or b"") + b">"
        text = decode(source)
        if text.strip():
            print("  x=%8.2f  y=%8.2f  %s" % (x, y, text))
