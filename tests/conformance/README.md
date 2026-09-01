# Conformance

The engine never asks the DOM how wide text is. It reads advances out of the
font file, so line breaking is deterministic and identical in Node, in a
browser, and in a test — and the browser cannot disagree with the server about
where a line ends.

That removes one risk and creates another: **if our reading of the font is
wrong, every page boundary is wrong and nothing else in the system would
notice.** These cases exist to catch exactly that.

## How it is checked

`expected-widths.spec.ts` measures every case in `browser-parity.cases.ts`
straight from the vendored fonts and writes `.conformance/parity-expectations.json`.
The browser side loads the same font files, measures with canvas `measureText`,
and compares.

Emitting a fixture rather than hard-coding numbers keeps ONE source of truth —
the font files. A constant pasted into the browser check would keep passing
against a stale expectation after a font is re-vendored, which is worse than no
check at all.

## Result after GPOS kerning (Chrome)

| case | engine | browser | delta |
|---|---|---|---|
| `xxxxxxxxxx` monospaced | 96.015625 | 96.01563 | **0** |
| `Договор аренды` monospaced | 134.421875 | 134.42188 | **0** |
| `nnnnmmmm` Caladea (1000 upem) | 87.744 | 87.74396 | **~0** |
| `nonsense` Liberation Serif (2048 upem) | 58.65625 | 58.65625 | **0** |
| `AVAVAVAV` Liberation Serif (kerned) | 78 | 78 | **0** |
| `AVAVAVAV` Liberation Mono (no kern feature) | 76.8125 | 76.8125 | **0** |
| `Wave To Avoid` Carlito (GPOS-only) | 94.7109375 | 94.71094 | **0** |

Worst delta **0.00004px**, and that is float formatting rather than a
disagreement. The last three rows carry the weight: we kern exactly as the
browser does, we do NOT invent kerning for a font that declares none, and the
default font — whose kerning exists only in GPOS — is right too.

Before GPOS kerning was built the fifth row read **92.4375 vs 78, a
−14.4375px gap**: the browser kerned and we did not. That number was recorded here as the target GPOS
had to hit, and hitting it exactly is the evidence the implementation is right
rather than merely closer.

 **A lesson from building this.** The Liberation Mono row was briefly reported
as a −19.2px disagreement, which sent me looking for a parser bug that did not
exist. The cause was a hand-typed expectation: the ten-character value pasted
into an eight-character case. The engine and the browser had agreed all along.
**Never hand-type an expected value into the browser check** — read it from the
generated fixture, which is the whole reason the fixture exists.

## Running it

```
npm test                 # writes .conformance/parity-expectations.json
```

The browser half is still driven manually against a served copy of the fonts.
Automating it (Playwright, as `@coolms/designer` does for visual tests) is the
follow-up — worth doing, because a parity check nobody runs is a parity check
that fails silently.
