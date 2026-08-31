"""Filled rectangles with the COLOUR in force, so a shaded cell can be named.

`pdf-rules.py` reports every rectangle and `pdf-strokes.py` carries the stroke
width; neither carries the FILL, and a conditional table format is a fill.

The colour is graphics state — `rg` sets it and it persists until the next one —
so the stream is walked in order, exactly as the stroke probe walks `w`. Walked
as TOKENS rather than lines: a content stream puts as many operators on a line
as it likes, and a line-by-line reader finds nothing at all.
"""
import re
import sys
import zlib

NUMBER = re.compile(r'^-?[\d.]+$')


def streams(data):
    for match in re.finditer(rb'stream\r?\n(.*?)\r?\nendstream', data, re.S):
        raw = match.group(1)
        try:
            yield zlib.decompress(raw)
        except zlib.error:
            yield raw


def main(path):
    data = open(path, 'rb').read()
    for index, stream in enumerate(streams(data)):
        tokens = stream.decode('latin1').replace('\n', ' ').replace('\r', ' ').split()
        printed = False
        fill = None
        operands = []
        for token in tokens:
            if NUMBER.match(token):
                try:
                    operands.append(float(token))
                except ValueError:
                    # A stream that is not a content stream at all — a font
                    # program, say — can hold anything. Skipped rather than
                    # crashing the walk over the streams that ARE.
                    operands = []
                continue

            if 'rg' == token and len(operands) >= 3:
                fill = '#%02X%02X%02X' % tuple(
                    int(round(max(0.0, min(1.0, value)) * 255)) for value in operands[-3:])
            elif 're' == token and len(operands) >= 4 and fill is not None:
                x, y, width, height = operands[-4:]
                # The page's own background is not a cell.
                if not (width > 500 and height > 700):
                    if not printed:
                        print('--- stream %d ---' % index)
                        printed = True
                    print('  fill %-8s x=%8.2f y=%8.2f w=%7.2f h=%6.2f'
                          % (fill, x, y, width, height))

            operands = []


if __name__ == '__main__':
    for argument in sys.argv[1:]:
        main(argument)
