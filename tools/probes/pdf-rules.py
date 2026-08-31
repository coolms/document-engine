"""Print every rectangle and line segment a PDF's content streams draw."""
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


NUM = r'([-\d.]+)'

for index, content in enumerate(streams(open(sys.argv[1], 'rb').read())):
    text = content.decode('latin-1')
    found = []

    for m in re.finditer(r'%s %s %s %s re' % (NUM, NUM, NUM, NUM), text):
        found.append(('re  ', *(float(g) for g in m.groups())))
    for m in re.finditer(r'%s %s m\s+%s %s l' % (NUM, NUM, NUM, NUM), text):
        found.append(('line', *(float(g) for g in m.groups())))

    if not found:
        continue

    print('--- stream %d ---' % index)
    for entry in found:
        print('  ' + ' '.join(
            part if isinstance(part, str) else '%9.3f' % part for part in entry))
