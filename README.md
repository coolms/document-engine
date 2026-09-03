# @coolms/document-engine

[![CI](https://github.com/coolms/document-engine/actions/workflows/ci.yml/badge.svg)](https://github.com/coolms/document-engine/actions/workflows/ci.yml)
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

```bash
npm install @coolms/document-engine
```

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

Node >=20. No runtime dependencies; TypeScript and vitest are dev-only.

## Branches

`develop` is the default and where work lands; `main` carries releases. The package is a pre-release and carries no compatibility promise.

## Licence

MIT — see [LICENSE](LICENSE), which covers this package's own code.

The font files under `assets/fonts` are third-party and carry their own terms.
`assets/fonts/fonts.manifest.json` records where each one came from: Liberation,
Carlito and Caladea were extracted from the `gotenberg/gotenberg:8` image
(Debian packages `fonts-liberation`, `fonts-crosextra-carlito`,
`fonts-crosextra-caladea`), and Gelasio came from the designer's own repository
under OFL-1.1. They are vendored rather than installed because the browser and
the renderer must shape text from the same bytes or the page boundaries
diverge.
