# Changelog

All notable changes to `@coolms/document-engine` are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

This file starts at the version named below, which is what the registry
currently serves. Earlier alphas are deliberately not reconstructed: entries are written
in the same commit as the work they describe, and inventing the ones that
predate this file would be a worse record than not having them.

## Unreleased

## 0.1.0-alpha.4 -- 2026-10-16

### Added

- **Each font family's licence text travels beside its files.** All 24 faces
  are licensed under the SIL Open Font License 1.1, which requires its text to
  go wherever the fonts go, and the package had shipped none of it:
  `assets/fonts/OFL-Liberation.txt` (Liberation Serif, Sans and Mono),
  `assets/fonts/OFL-Carlito.txt`, `assets/fonts/OFL-Caladea.txt` and
  `assets/fonts/OFL-Gelasio.txt`, each the text its authors publish, copied
  byte for byte. Both registries carry them from this version on: npm in the
  tarball, Packagist in the `coolms/document-fonts` archive. The README's
  Licence section lists them, and says MIT covers this package's code, not the
  fonts.
- Declares `support` (`issues`, `source`) in composer.json and `repository`
  and `bugs` in package.json -- so a page imported from this package, and the
  catalogue, know where a correction is filed. Packagist filled the gap from
  GitHub when the manifest was silent; the declared field is the one that holds
  on any registry.

### Changed

- The README and this changelog name a shipped file by its path from the
  package root: `assets/fonts/fonts.manifest.json`, where they used to give the
  manifest's file name alone, which names no file at that root.

### Notes

- On npm this version follows 0.1.0-alpha.2: 0.1.0-alpha.3 was published to
  Packagist only.

## 0.1.0-alpha.3 -- 2026-09-04

### Added

**This repository now publishes to Packagist as well as npm.** The engine
stays `@coolms/document-engine` on npm; the font set is `coolms/document-fonts`
on Packagist. Two registries, two names, each true where it appears -- the
composer archive contains no engine, so calling it one would have been a
promise the package could not keep.

```bash
composer require coolms/document-fonts
```

That installs `assets/fonts/` -- the 24 faces and `assets/fonts/fonts.manifest.json` --
and nothing else. No PHP, no autoloader, and **no PHP version constraint**,
because there is no PHP in it to run: constraining a runtime the package
never invokes would refuse installations for no reason.

!! **Why it exists.** The CoolMS application read the manifest from a path
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
## 0.1.0-alpha.2 -- 2026-09-03

**A pre-release, carrying no compatibility promise.** Published under the
`alpha` dist-tag.

A framework-agnostic document layout engine in vanilla TypeScript: the OOXML
model, deterministic text measurement read from the font files themselves, and
pagination with reflow.

**Zero runtime dependencies**, deliberately -- the arithmetic is ours, so the
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
