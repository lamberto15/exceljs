# Releasing

turbosheet uses [Changesets](https://github.com/changesets/changesets) and a GitHub Actions workflow
([`.github/workflows/release.yml`](./.github/workflows/release.yml)). Nothing is published
from a developer machine in the normal flow.

## Day to day: add a changeset

For every user-visible change (feature, fix, behavior change):

```sh
pnpm changeset
```

Pick the bump (`patch` fix, `minor` feature or, while 0.x, a breaking change; `major` after 1.0), write a
one-line summary for the changelog, and commit the generated `.changeset/*.md` file with your PR.
Docs-only, test-only and CI changes do not need one (no CI check enforces changesets).

## How a release happens

1. PRs with changesets land on `main`.
2. The **release** workflow (`release.yml`) runs on every push to `main` (and can be run by hand). It runs
   `typecheck`, `test`, `build`, `size`, then the Changesets action:
   - with pending changesets it opens (or updates) a **"chore(release): version packages"** PR that
     bumps `package.json` and writes `CHANGELOG.md`;
   - when that PR is merged there are no pending changesets and the version is not on npm, so the action
     runs `pnpm release` (`changeset publish`), which publishes to npm (public access, from
     `.changeset/config.json`) over npm trusted publishing (OIDC, `id-token: write`, no npm token), with
     provenance, and creates the git tag.
3. Runs with no changesets and a version already on npm do nothing.

Checks before merging the version PR: `CHANGELOG.md` reads well, `pnpm build && pnpm size` is green,
`npm pack --dry-run` lists only `dist/`, `README.md`, `LICENSE`, `CHANGELOG.md`, `package.json`.

## One-time setup (done by the maintainer, not by the repo)

Publishing uses **npm trusted publishing (OIDC)**: GitHub Actions proves its identity to npm, so there
is no npm token secret (`NODE_AUTH_TOKEN` is deliberately unset), and provenance is attached automatically.
Requirements: Node >= 22.14 (the workflow uses Node 24) and a pnpm that falls back to the OIDC token (>= 11.1.3;
the repo pins pnpm 12 in `packageManager`).

- An npm account with **2FA enabled** (a passkey works; a local `npm publish` then needs
  `--auth-type=web` in an interactive terminal).
- Optional repo secret **`RELEASE_PAT`** (fine-grained PAT, Contents + Pull requests read/write, this
  repo only) so the "chore(release): version packages" PR triggers CI; without it the workflow uses
  `GITHUB_TOKEN` and that PR gets no checks.
- Settings, Actions, General, Workflow permissions: **Read and write**, and tick "Allow GitHub Actions
  to create and approve pull requests" (for the version PR).
- The repository must be public (it is) for provenance.
- After the first publish (below): on npmjs.com open the `turbosheet` package, **Settings, Trusted
  publishing**, add a GitHub Actions publisher: owner `lamberto15`, repository `turbosheet`, workflow
  `release.yml` (exact, case-sensitive; environment empty). The repository was renamed from
  `lamberto15/exceljs`; the publisher must name the current repository. Then, under the same settings, set publishing access to
  require 2FA and disallow tokens.

## 0.1.0: the first publish (manual, done)

npm only lets a trusted publisher be configured for a package that already exists, so 0.1.0 was
published from a maintainer's machine (before the package moved to the repository root, #31), roughly:

```sh
pnpm install --frozen-lockfile
npm login                      # browser / 2FA
npm pack --dry-run             # inspect the file list
npm publish --access public    # prepublishOnly runs build + tests; a local publish has no provenance
git tag turbosheet@0.1.0 && git push origin turbosheet@0.1.0
```

Then the trusted publisher was configured (above). From 0.2.0 on, every release goes through changesets and
the **release** workflow, which publishes over OIDC with provenance via `pnpm run release`
(`changeset publish`, which skips versions already on npm).

## Deprecating or fixing a bad release

`npm deprecate turbosheet@0.1.1 "reason"`; publish a fixed version through the normal flow. Unpublishing is
only possible within 72 hours and is discouraged.
