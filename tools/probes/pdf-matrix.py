"""Per-page text with its FULL text matrix, so rotation can be read off.

`pdf-positions.py` keeps only e/f — the translation — which is enough until the
text is turned on its side. `a b c d` carry the rotation and scale: an upright
run is `1 0 0 1`, a quarter turn anticlockwise is `0 1 -1 0`, and clockwise is
`0 -1 1 0`.
"""
import re
import sys

from os.path import dirname, join

source = open(join(dirname(__file__), 'pdf-positions.py')).read()
exec(source.split('NUM = rb')[0])

NUM = rb"(-?[\d.]+)"
TM = re.compile(rb"\s+".join([NUM] * 6) + rb"\s+Tm")
TD = re.compile(NUM + rb"\s+" + NUM + rb"\s+Td")
SHOW = re.compile(rb"\[((?:[^\[\]\\]|\\.)*)\]\s*TJ|<([0-9A-Fa-f]*)>\s*Tj")


def decode(text):
    out = []
    for hexed in re.findall(rb"<([0-9A-Fa-f]*)>", text):
        for index in range(0, len(hexed) - 1, 2):
            out.append(unicode_of.get(int(hexed[index:index + 2], 16), "?"))
    return "".join(out)


for index, page in enumerate(order):
    body = objects[page]
    content = re.search(rb"/Contents\s+(\d+)\s+\d+\s+R", body)
    blob = b"" if content is None else inflate(objects[int(content.group(1))])
    print("--- PAGE %d ---" % (index + 1))

    matrix = (1.0, 0.0, 0.0, 1.0)
    x = y = 0.0
    pattern = TM.pattern + rb"|" + TD.pattern + rb"|" + SHOW.pattern
    for token in re.finditer(pattern, blob):
        groups = token.groups()
        if groups[0] is not None:
            matrix = tuple(float(value) for value in groups[0:4])
            x, y = float(groups[4]), float(groups[5])
            continue
        if groups[6] is not None:
            x, y = float(groups[6]), float(groups[7])
            continue
        text = decode(groups[8] if groups[8] is not None else b"<" + (groups[9] or b"") + b">")
        if not text.strip():
            continue
        turn = '' if matrix == (1.0, 0.0, 0.0, 1.0) else '  [%g %g %g %g]' % matrix
        print("  x=%8.2f  y=%8.2f  %-14s%s" % (x, y, text, turn))
