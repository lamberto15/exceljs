# Releasing

xlsxkit uses [Changesets](https://github.com/changesets/changesets) and a GitHub Actions workflow
([`.github/workflows/next-release.yml`](../.github/workflows/next-release.yml)). Nothing is published
from a developer machine in the normal flow.

## Day to day: add a changeset

For every user-visible change (feature, fix, behavior change), in `next/`:

```sh
pnpm changeset
```

Pick the bump (`patch` fix, `minor` feature or, while 0.x, a breaking change; `major` after 1.0), write a
one-line summary for the changelog, and commit the generated `.changeset/*.md` file with your PR.
Docs-only, test-only and CI changes do not need one (`pnpm changeset --empty` if CI asks).

## How a release happens

1. PRs with changesets land on `next`.
2. A maintainer runs the workflow by hand (GitHub, Actions, **next-release**, Run workflow on `next`).
   It is manual-only on purpose so nothing reaches npm by accident; to automate it later add a
   `push` trigger. It runs `typecheck`, `test`, `build`, `size`, then the Changesets action:
   - with pending changesets it opens (or updates) a **"chore(release): version packages"** PR that
     bumps `package.json` and writes `CHANGELOG.md`;
   - when that PR is merged there are no pending changesets and the version is not on npm, so the action
     runs `pnpm release` (`changeset publish`), which publishes to npm with `--access public` and
     provenance (`NPM_CONFIG_PROVENANCE=true`, `id-token: write`) and creates the git tag.
3. Runs with no changesets and a version already on npm do nothing.

Checks before merging the version PR: `CHANGELOG.md` reads well, `pnpm build && pnpm size` is green,
`npm pack --dry-run` lists only `dist/`, `README.md`, `LICENSE`, `CHANGELOG.md`, `package.json`.

## One-time setup (done by the maintainer, not by the repo)

Publishing uses **npm trusted publishing (OIDC)**: GitHub Actions proves its identity to npm, so there
is no npm token secret, and provenance is attached automatically. Requirements: npm CLI >= 11.5.1 (the
workflow upgrades npm) and Node >= 22.14 (the workflow uses Node 24).

- An npm account with **2FA enabled**.
- Settings, Actions, General, Workflow permissions: **Read and write**, and tick "Allow GitHub Actions
  to create and approve pull requests" (for the version PR).
- The repository must be public (it is) for provenance.
- After the first publish (below): on npmjs.com open the `xlsxkit` package, **Settings, Trusted
  publishing**, add a GitHub Actions publisher: owner `lamberto15`, repository `exceljs`, workflow
  `next-release.yml` (environment empty). Then, under the same settings, set publishing access to
  require 2FA and disallow tokens.

## 0.1.0: the first publish (manual, once)

npm only lets a trusted publisher be configured for a package that already exists, so 0.1.0 is
published from a maintainer's machine. `package.json` is already `0.1.0` and `CHANGELOG.md` has the
entry, so no changeset is needed.

```sh
cd next
pnpm install --frozen-lockfile
npm login                      # browser / 2FA
npm pack --dry-run             # inspect the file list
npm publish --access public    # prepublishOnly runs build + tests; a local publish has no provenance
git tag xlsxkit@0.1.0 && git push origin xlsxkit@0.1.0
```

Then configure the trusted publisher (above). From 0.2.0 on, every release goes through changesets and
the **next-release** workflow (Actions, next-release, Run workflow), which publishes over OIDC with
provenance via `scripts/publish-if-new.mjs` (skips versions already on npm).

## Deprecating or fixing a bad release

`npm deprecate xlsxkit@0.1.1 "reason"`; publish a fixed version through the normal flow. Unpublishing is
only possible within 72 hours and is discouraged.
