#!/usr/bin/env python3
"""Where a PDF places each image, and which way up.

An image is drawn as `q <a b c d e f> cm /Name Do Q`: the matrix maps the unit
square onto the page, so `a` and `d` are its width and height, `e` and `f` its
lower-left corner, and `b`/`c` say it was turned. `[0 h -w 0]` is a quarter
turn anticlockwise, `[0 -h w 0]` clockwise.

Neither pdf-positions.py nor pdf-rules.py sees any of this: an image is not
text and not a rule. Reach for this one when the question is where a PICTURE
landed.

    python3 tools/probes/pdf-images.py /tmp/out.pdf
"""
import re
import sys
import zlib


def streams(raw: bytes):
    """Every content stream in the file, decompressed."""
    for match in re.finditer(rb'stream\r?\n', raw):
        start = match.end()
        end = raw.find(b'endstream', start)
        if end < 0:
            continue
        body = raw[start:end]
        try:
            yield zlib.decompress(body).decode('latin-1')
        except zlib.error:
            try:
                yield body.decode('latin-1')
            except UnicodeDecodeError:
                continue


NUMBER = r'(-?[\d.]+)'
PLACED = re.compile(
    r'\s+'.join([NUMBER] * 6) + r'\s+cm\s+(?:[^D]*?)/([A-Za-z0-9_.]+)\s+Do',
    re.S,
)


def turn_of(b: float, c: float) -> str:
    if abs(b) < 1e-9 and abs(c) < 1e-9:
        return 'upright'
    if b > 0 and c < 0:
        return 'turned anticlockwise'
    if b < 0 and c > 0:
        return 'turned clockwise'

    return 'turned [b=%.3f c=%.3f]' % (b, c)


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2

    with open(sys.argv[1], 'rb') as handle:
        raw = handle.read()

    found = 0
    for index, stream in enumerate(streams(raw)):
        placements = PLACED.findall(stream)
        if not placements:
            continue
        print('--- stream %d ---' % index)
        for a, b, c, d, e, f, name in placements:
            found += 1
            a, b, c, d, e, f = (float(v) for v in (a, b, c, d, e, f))
            # The turned cases carry their size in b/c rather than a/d.
            across = abs(a) if abs(a) > 1e-9 else abs(c)
            along = abs(d) if abs(d) > 1e-9 else abs(b)
            print('  /%-10s x=%8.2f y=%8.2f  %6.2f x %6.2f  %s'
                  % (name, e, f, across, along, turn_of(b, c)))

    if 0 == found:
        print('no placed images')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
