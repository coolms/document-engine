# Changelog

All notable changes to `@coolms/document-engine` are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

This file starts at the version named below, which is what the registry
currently serves. Earlier alphas are deliberately not reconstructed: entries are written
in the same commit as the work they describe, and inventing the ones that
predate this file would be a worse record than not having them.

## 0.1.0-alpha.3 — 2026-09-04

### Added

**This repository now publishes to Packagist as well as npm.** The engine
stays `@coolms/document-engine` on npm; the font set is `coolms/document-fonts`
on Packagist. Two registries, two names, each true where it appears -- the
composer archive contains no engine, so calling it one would have been a
promise the package could not keep.

```bash
composer require coolms/document-fonts
```

That installs `assets/fonts/` -- the 24 faces and `fonts.manifest.json` --
and nothing else. No PHP, no autoloader, and **no PHP version constraint**,
because there is no PHP in it to run: constraining a runtime the package
never invokes would refuse installations for no reason.

⚠️ **Why it exists.** The CoolMS application read the manifest from a path
that only exists in a development checkout, so no installed application
could ever have found it. A PHP application needs these files and cannot
fetch them from a registry it does not use.

### Changed

- The composer archive is `assets/fonts` plus the documents that explain it.
  It had been carrying `tools/` -- developer probes in JavaScript and Python
  -- along with the TypeScript config; a PHP project has no use for any of
  it.
- The README opens by saying which of the two packages you installed. Anyone
  arriving from Packagist previously met a Node badge, `npm install` as the
  only install line, and "Requirements: Node >=20".

The npm package is unchanged in content.
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
