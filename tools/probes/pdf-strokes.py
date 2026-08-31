"""Every line segment a PDF draws, WITH the stroke width in force for it.

`pdf-rules.py` reads positions and nothing else, which is enough until the
question is how THICK a rule is. Width comes from the `w` operator, which is
state: it applies to every stroke after it until the next one. So the stream
has to be walked in order rather than scanned for one pattern at a time.
"""
import re
import sys
import zlib


def streams(data):
    for match in re.finditer(rb'stream\r?\n', data):
        start = match.end()
        end = data.find(b'endstream', start)
        if end < 0:
            continue
        raw = data[start:end]
        try:
            yield zlib.decompress(raw)
        except zlib.error:
            yield raw


NUM = r'-?[\d.]+'
TOKEN = re.compile(
    r'(?P<w>%s)\s+w\b'
    r'|(?P<x1>%s)\s+(?P<y1>%s)\s+m\s+(?P<x2>%s)\s+(?P<y2>%s)\s+l'
    % (NUM, NUM, NUM, NUM, NUM))

for index, content in enumerate(streams(open(sys.argv[1], 'rb').read())):
    text = content.decode('latin-1')
    width = 1.0
    found = []

    for match in TOKEN.finditer(text):
        if match.group('w') is not None:
            width = float(match.group('w'))
            continue
        found.append((
            float(match.group('x1')), float(match.group('y1')),
            float(match.group('x2')), float(match.group('y2')), width))

    if not found:
        continue

    print('--- stream %d ---' % index)
    for x1, y1, x2, y2, w in found:
        print('  line %9.3f %9.3f %9.3f %9.3f   width %6.3f' % (x1, y1, x2, y2, w))
