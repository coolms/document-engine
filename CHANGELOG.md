# Changelog

All notable changes to `@coolms/document-engine` are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

This file starts at the version named below, which is what the registry
currently serves. Earlier alphas are deliberately not reconstructed: entries are written
in the same commit as the work they describe, and inventing the ones that
predate this file would be a worse record than not having them.

## 0.1.0-alpha.2 — 2026-09-03

**A pre-release, carrying no compatibility promise.** Published under the
`alpha` dist-tag.

A framework-agnostic document layout engine in vanilla TypeScript: the OOXML
model, deterministic text measurement read from the font files themselves, and
pagination with reflow.

**Zero runtime dependencies**, deliberately — the arithmetic is ours, so the
browser and the server agree by construction rather than by both happening to
call the same library. The fonts are vendored rather than installed for the
same reason: if the browser and the renderer do not shape text from the same
bytes, page boundaries diverge and the document on screen stops agreeing with
the `.docx` it produces.

1285 tests across 22 files (`npm test`).

### Added

- A Licence section in the README, which the package had never carried. It
  names where each vendored font came from, recorded in
  `assets/fonts/fonts.manifest.json`, rather than asserting licence
  identifiers.
