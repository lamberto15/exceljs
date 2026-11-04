# Contributing

The library lives in `next/` of this repository; run everything below from that directory.

## Setup

Requires Node 22 or newer and [pnpm](https://pnpm.io) (version pinned in `package.json`).

```sh
pnpm install --frozen-lockfile
```

## Scripts

| Script           | What it does                                                           |
| ---------------- | ---------------------------------------------------------------------- |
| `pnpm test`      | Vitest suite (`pnpm test:watch` to iterate)                            |
| `pnpm typecheck` | `tsc --noEmit`                                                         |
| `pnpm lint`      | Biome check (`pnpm format` to fix formatting)                          |
| `pnpm build`     | Build `dist/` with tsdown                                              |
| `pnpm size`      | Bundle-size budgets (build first)                                      |
| `pnpm bench`     | turbosheet vs ExcelJS, see [docs/benchmarks.md](./docs/benchmarks.md)  |

Before opening a PR, `pnpm typecheck && pnpm test && pnpm lint && pnpm build && pnpm size` should pass.

## The Excel pack

Excel is stricter than LibreOffice and ExcelJS. `node test/smoke/excel-pack.mjs [outDir]` (after
`pnpm build`) generates one file per feature plus a checklist, to open in real Microsoft Excel and check for
"We found a problem with some content" prompts. Run it whenever you change a writer.

## Adding a feature

Features (tables, charts, images, ...) are self-contained modules. Follow
[src/xlsx/features/README.md](./src/xlsx/features/README.md), add tests, and document it in
[docs/features.md](./docs/features.md). Keep `src/` free of `node:*` imports and `Buffer`, and stay within the
size budgets.

## Changesets

User-visible changes need a changeset:

```sh
pnpm changeset
```

Pick the bump, write a one-line summary, and commit the generated `.changeset/*.md`. Docs-only, test-only and
CI changes do not need one.

## Pull requests

- Use conventional commit messages (`feat:`, `fix:`, `docs:`, ...).
- Open PRs against `next`. Required checks: the test matrix (Node 22 and 24 on ubuntu, macos and windows) and
  `quality` (lint, build, size, LibreOffice).
- Keep PRs focused, with tests for behavior changes.

## Releases

Maintainers release through CI with npm provenance; see [RELEASING.md](./RELEASING.md).
