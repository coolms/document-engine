# @coolms/document-engine

> **Reading this from Packagist?** This repository publishes twice. The
> TypeScript engine below is the npm package. What Composer installs is
> `coolms/document-fonts`: the 24 font files in `assets/fonts/`, each
> family's licence beside them, and `assets/fonts/fonts.manifest.json`, which
> names them -- with no PHP and no autoloader. It exists so a
> PHP application can reach the same fonts, from the same version, without
> a Node toolchain. If you are looking for a document layout library for
> PHP, this is not one.

[![CI](https://github.com/coolms/document-engine/actions/workflows/ci.yml/badge.svg)](https://github.com/coolms/document-engine/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@coolms/document-engine?label=npm%20%40coolms%2Fdocument-engine)](https://www.npmjs.com/package/@coolms/document-engine)
[![Packagist](https://img.shields.io/packagist/v/coolms/document-fonts?label=composer%20coolms%2Fdocument-fonts)](https://packagist.org/packages/coolms/document-fonts)
[![Node](https://img.shields.io/badge/node-%E2%89%A5%2020-339933)](https://nodejs.org)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

**Framework-agnostic vanilla TypeScript document layout engine: OOXML model, deterministic text measurement from font files, and pagination with reflow. Zero runtime dependencies — the arithmetic is ours, so the browser and the server agree by construction.**

## Why it has no dependencies

The engine decides where a line breaks and where a page ends. If that arithmetic
came from a library, the browser and the server would each get whatever version
they resolved, and a document would paginate one way in the editor and another
way in the PDF. Owning the arithmetic is what makes the two agree by
construction, so the absence of dependencies is the feature rather than an
economy.

Text measurement reads the font files in `assets/fonts` directly -- the metrics
come from the same faces the renderer uses, not from a table that can drift away
from them.

## Install

**The engine, from npm:**

```bash
npm install @coolms/document-engine
```

**The fonts, from Packagist:**

```bash
composer require coolms/document-fonts
```

That gives you `vendor/coolms/document-fonts/assets/fonts/` -- the 24 faces,
each family's licence (`assets/fonts/OFL-*.txt`) and
`assets/fonts/fonts.manifest.json`. Serve them at whatever URL your application uses;
CoolMS serves them under the admin at `assets/document-fonts/`.

```ts
import { ... } from '@coolms/document-engine';
```

## Scripts

| | |
|---|---|
| `npm run build` | `tsc -p tsconfig.build.json` |
| `npm run test` | `vitest run` |
| `npm run test:watch` | `vitest` |
| `npm run typecheck` | `tsc --noEmit` |
| `npm run audit:claims` | `node tools/audit-claims.mjs` |

## Requirements

**npm:** Node 22.12 or later on the 22 line, 24 or 26 (`engines.node`:
`^22.12.0 || ^24.0.0 || ^26.0.0`, the lines CI runs). No runtime
dependencies; TypeScript and vitest are dev-only.

**Composer:** none. The package declares no PHP constraint because it
contains no PHP -- it is font files and a JSON manifest, and constraining a
runtime it never invokes would refuse installations for no reason.

## Branches

`develop` is the default and where work lands; `main` carries releases. The package is a pre-release and carries no compatibility promise.

## Licence

MIT — see [LICENSE](LICENSE), which covers this package's own code.

The font files under `assets/fonts/` are third-party, and MIT does not cover
them. Every one is licensed under the SIL Open Font License 1.1, which says so in
the font itself, and each family's licence text travels beside its files:

| Family | Faces | Licence text |
|---|---|---|
| Liberation Serif, Sans and Mono 2.1.5 | 12 | `assets/fonts/OFL-Liberation.txt` |
| Carlito 1.104 | 4 | `assets/fonts/OFL-Carlito.txt` |
| Caladea 1.001 | 4 | `assets/fonts/OFL-Caladea.txt` |
| Gelasio 1.008 | 4 | `assets/fonts/OFL-Gelasio.txt` |

Each text is the one its authors publish, copied byte for byte, copyright and
Reserved Font Names included. `assets/fonts/fonts.manifest.json` records where
each font came from: Liberation,
Carlito and Caladea were extracted from the `gotenberg/gotenberg:8` image
(Debian packages `fonts-liberation`, `fonts-crosextra-carlito`,
`fonts-crosextra-caladea`), and Gelasio came from the designer's own repository
under OFL-1.1. They are vendored rather than installed because the browser and
the renderer must shape text from the same bytes or the page boundaries
diverge.
