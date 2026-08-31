"""Which WordprocessingML elements appear in the fixtures and NOWHERE in src/.

The reader's own diagnostics list what it knows it cannot do. This lists what it
does not know about at all: every element name in every vendored .docx, less
every name mentioned anywhere in the package's source.
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
files = collections.defaultdict(set)
for path in sorted(glob.glob(os.path.join(ROOT, 'tests/fixtures/docx/*.docx'))):
    name = os.path.basename(path)
    try:
        with zipfile.ZipFile(path) as package:
            for entry in package.namelist():
                if not entry.startswith('word/') or not entry.endswith('.xml'):
                    continue
                xml = package.read(entry).decode('utf-8', 'replace')
                for tag in re.findall(r'<(w:[A-Za-z0-9]+)[ />]', xml):
                    seen[tag] += 1
                    files[tag].add(name)
    except zipfile.BadZipFile:
        print('!! not a zip: %s' % name)

unknown = []
for tag, count in seen.items():
    local = tag.split(':', 1)[1]
    # The reader matches on the qualified name for `w:` elements and on the
    # LOCAL name for anything reached by local-name search, so look for both.
    if ("'%s'" % tag) in source or ("'%s'" % local) in source:
        continue
    unknown.append((count, tag, sorted(files[tag])))

print('%d distinct elements, %d never mentioned in src/\n' % (len(seen), len(unknown)))
for count, tag, where in sorted(unknown, reverse=True):
    print('%5d  %-22s %s' % (count, tag, ', '.join(where[:4])
                             + ('' if len(where) < 5 else ' +%d' % (len(where) - 4))))
