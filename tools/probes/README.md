# The instruments

Every geometric number in this engine was read off a PDF that LibreOffice
printed. These four scripts are how it was read. They lived on one developer's
machine for the whole of the arc that measured the engine — which meant the
numbers were reproducible and the *method* was not.

## The loop

Build a `.docx` stating exactly one thing, print it through LibreOffice, and
read the coordinates back out:

```bash
# Render the probe through a headless LibreOffice, then measure the PDF.
soffice --headless --convert-to pdf probe.docx
python3 tools/probes/pdf-positions.py probe.pdf
```

If `soffice` is not on the machine running the probe, any headless
LibreOffice conversion service does the same job. What matters is that the
SAME renderer prints every fixture in a comparison: these numbers only mean
something relative to one engine.

## Which instrument answers which question

| Script | Reads | Use it for |
| --- | --- | --- |
| `pdf-positions.py` | text with its x/y, decoded through each subset font's ToUnicode CMap | where a baseline sits, where a line wraps, how far a step is |
| `pdf-rules.py` | every `re` rectangle and `m … l` segment | borders and rules, which are invisible to text extraction |
| `pdf-matrix.py` | the full text matrix, not only its translation | rotated text — `[0 1 -1 0]` is a quarter turn anticlockwise |
| `pdf-strokes.py` | segments **with the stroke width in force** | how THICK a rule is; `w` is state, so the stream is walked in order |
| `pdf-images.py` | every `cm … Do` placement, with its size and turn | where a PICTURE landed; an image is neither text nor a rule, so nothing else sees it |
| `unread-elements.py` | every element name in the fixtures, less every name mentioned in `src/` | what the reader has never HEARD of — the gaps its own diagnostics cannot report |
| `unread-attributes.py` | the same net over ATTRIBUTES, reported with the elements carrying them | a modifier on an element the reader DOES know — `w:gutter` on `w:pgMar` was found this way |
| `pdf-fills.py` | every filled rectangle with the COLOUR in force | which cells a shading reached — a conditional table format is a fill, and neither the rule nor the stroke probe carries one |
| `run-width.mjs` | how wide a STRING is, through the engine's own `advanceOf` | the run you have to SUBTRACT to turn absolute page positions into a feature's own geometry |

## What they taught, the hard way

- **A rule is invisible to text extraction.** `pdf-positions.py` shows a
  bordered paragraph exactly as it shows a plain one. Reach for
  `pdf-rules.py` before concluding that a border does nothing.
- **A null result needs a stimulus big enough to see.** `w:tblCellSpacing` was
  closed at 120 twips and re-checked at 400; section `w:vAlign` on A4 and
  re-checked on a page 200pt tall. Both held — but a weak probe and a real
  null look identical.
- **Count what came back.** Invalid XML — `<w:br/>` inside `<w:t>` — made
  LibreOffice silently drop two of three tables, which reads exactly like a
  feature that does nothing.
- **A null result needs a CONTROL, not just a bigger stimulus.** `pdf-images.py`
  reported no picture inside a turned cell. That is only evidence once the same
  probe finds the picture in a document known to have one, and the same fixture
  with the turn removed prints it — otherwise a typo in the drawing XML and a
  real null look identical.
- **One line cannot report a line BOX.** A single-line fixture says where one
  baseline landed and nothing else, so a placement and the box it sits in look
  identical — one probe read three baselines and could not separate them. Wrap
  the text: the SECOND baseline is the first line's whole height, less its own
  ascent.
- **A value that is almost always ZERO hides in plain sight.** `w:gutter` is
  on 75 of these fixtures and was read in none: Word writes it on every
  document and only a file set up for binding gives it a value, so every
  test passed while the attribute did nothing. When a sweep turns up
  an attribute this common, check what the fixtures SAY it is before
  concluding it does not matter.
- **The room a thing keeps and the thing it draws must be the SAME number.**
  A table reserved its own declared rule while drawing the cells' — four
  times the gap, every row. When a box is the right size but in the
  wrong place, look for two code paths answering one question.
- **Round the DIFFERENCE, not the ends.** One spacing printed as 10.47 in
  one table and 10.48 in another because each baseline was rounded before
  subtracting. It reads exactly like a real divergence.
- **A specification is not a measurement.** ECMA-376 lists the conditional
  table formats weakest-first with the VERTICAL bands below the horizontal
  ones. LibreOffice draws them the other way round, and a table defining both
  came out entirely the vertical band colour. Read the order off the
  page, not off the standard.
- **An unknown enum value is not ignored — it is applied to EVERYTHING.** A
  `w:tblStylePr` typed `firstColumn` (the `w:tblLook` spelling) rather than
  `firstCol` (the `w:tblStylePr` one) dressed every cell of the table, which
  reads exactly like a condition that means "all of it".
- **A guard is invisible while the case it guards is undefined.** Excluding a
  header row from the banding changed nothing while the style defined only ONE
  band: the header's parity fell on the other. Define both.
- **A feature nobody parses cannot report itself.** The reader's diagnostics
  list what it knows it cannot do; `unread-elements.py` lists what it has never
  heard of. That is how `w:lvlOverride` was found — in a Word-AUTHORED fixture,
  ignored outright, and numbering three lists wrong.
- **Two passes over one input must agree, and that is answerable WITHOUT
  LibreOffice.** A table wrapped in a content control measured 6.67px to a
  column and the identical table unwrapped measured 312px — the same table,
  read twice, read differently. No reference renderer is needed to call that
  a defect, and the test asserts the two are EQUAL rather than asserting a
  number.
- **Measure the run you SUBTRACT, or subtract nothing.** A feature's own
  geometry usually has to be recovered by taking the text before it off an
  absolute position. Estimating that text's width by eye read one box as
  17.50, then 16.40, then 18.01 — three readings of an
  unchanging page, two of them published, because `P-before` was called "about
  36" where the face says 34.43. `run-width.mjs` is that subtraction, and it is
  the only one of these instruments that needs no PDF.
- **Do not carry more decimals than the page supports.** Positions come back to
  two decimals and the print already runs about 0.10 right of the engine, so
  agreement in the SECOND decimal is arithmetic, not measurement. The box above
  is 18.0 and change; the page first gave 18.02 and the instrument said 18.01.
- **Check the CODE before calling a number missing.** The inline VML box's line
  height was carried as the blocker for two slices and four prints measured the
  rule — which the engine already implemented, to a fiftieth of a point.
  "Not built" had been inferred from a diagnostic saying the box was DROPPED,
  which was true of its text and false about its geometry. Run the fixture
  through `openWordFile` -> `layoutSections` -> `renderPage` and dump the ops
  first; it takes seconds and says whether you are measuring a gap or a
  conformance.
- **An unexplained constant is an unasked question.** An 18.0 that "nothing in
  the file accounts for" was re-measured across five slices and read three
  different ways before anyone asked what it was MADE of: wrap distance, 9pt a
  side, DEFAULTED. A file is silent about a default precisely because
  the default is what it wants, so "not in the file" is the signature of one.
  Probe such an attribute at three values — unset, zero, and something large.
  The large one is the control that makes the zero mean anything.
- **Diff the PART LIST before theorising about the content.** Two fixtures with
  the same grid pitch and the same floor printed 18.00 and 11.50, and the font
  took the blame for two slices. It was not the font: one fixture carried a
  `word/settings.xml` and the hand-built one did not — and an EMPTY settings
  part behaves like a full one, so there was no declaration to find.
  Hand-built fixtures are minimal by design, which is their virtue for
  isolating an element and their trap at the level of PARTS: a missing part can
  put the renderer in legacy defaults and have the fixture measure behaviour no
  real document has. When a hand-built fixture disagrees with a Word-authored
  one, suspect the hand-built one. And remove a part's `[Content_Types].xml`
  Override and its relationship along with it, or the file is invalid and
  prints EMPTY — which reads exactly like the ablation having worked.
- **Probing a suspected gap often deletes it.** Seven elements left the backlog
  that way; `w:docGrid` went the other way after being dismissed twice from its
  name, and moves Latin text by 6.5pt a line.
