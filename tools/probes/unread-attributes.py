"""Which `w:` ATTRIBUTES appear in the fixtures and nowhere in src/.

The element sweep found `w:lvlOverride`. The same net over attributes
catches a different class: an element the reader knows, carrying a modifier it
does not — `w:hRule` on a row height, `w:type` on a break, and their like.

Attributes carried on elements the reader has never heard of are noise here, so
each one is reported with the elements it was seen on.
"""
import collections
import glob
import os
import re
import zipfile

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '../..')

source = ''
for path in glob.glob(os.path.join(ROOT, 'src/**/*.ts'), recursive=True):
    source += open(path, encoding='utf-8').read()

seen = collections.Counter()
carriers = collections.defaultdict(collections.Counter)
for path in sorted(glob.glob(os.path.join(ROOT, 'tests/fixtures/docx/*.docx'))):
    try:
        with zipfile.ZipFile(path) as package:
            for entry in package.namelist():
                if not entry.startswith('word/') or not entry.endswith('.xml'):
                    continue
                xml = package.read(entry).decode('utf-8', 'replace')
                for tag, body in re.findall(r'<(w:[A-Za-z0-9]+)((?:\s+[^<>]*)?)/?>', xml):
                    for attribute in re.findall(r'\b(w:[A-Za-z0-9]+)=', body):
                        seen[attribute] += 1
                        carriers[attribute][tag] += 1
    except zipfile.BadZipFile:
        pass

unknown = []
for attribute, count in seen.items():
    local = attribute.split(':', 1)[1]
    if ("'%s'" % attribute) in source or ("'%s'" % local) in source:
        continue
    unknown.append((count, attribute, carriers[attribute].most_common(4)))

print('%d distinct attributes, %d never mentioned in src/\n' % (len(seen), len(unknown)))
for count, attribute, where in sorted(unknown, reverse=True):
    print('%5d  %-20s %s' % (count, attribute,
                             ', '.join('%s(%d)' % pair for pair in where)))
