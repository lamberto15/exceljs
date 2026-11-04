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
2. The workflow runs on every push to `next`: it runs `typecheck`, `test`, `build`, `size`, then the
   Changesets action:
   - with pending changesets it opens (or updates) a **"chore(release): version packages"** PR that
     bumps `package.json` and writes `CHANGELOG.md`;
   - when that PR is merged there are no pending changesets and the version is not on npm, so the action
     runs `pnpm release` (`changeset publish`), which publishes to npm with `--access public` and
     provenance (`NPM_CONFIG_PROVENANCE=true`, `id-token: write`) and creates the git tag.
3. Pushes with no changesets and a version already on npm do nothing.

Checks before merging the version PR: `CHANGELOG.md` reads well, `pnpm build && pnpm size` is green,
`npm pack --dry-run` lists only `dist/`, `README.md`, `LICENSE`, `CHANGELOG.md`, `package.json`.

## One-time setup (done by the maintainer, not by the repo)

- An npm account that owns the `xlsxkit` name, with **2FA enabled**.
- An npm **automation** access token (granular token with publish rights on `xlsxkit`, 2FA bypass for
  CI is required for token publishing) saved as the repository secret **`NPM_TOKEN`**
  (Settings, Secrets and variables, Actions).
- Settings, Actions, General: allow GitHub Actions to create pull requests ("Allow GitHub Actions to
  create and approve pull requests"), and give workflows read and write permission.
- The repository must be public for npm provenance.
- Optional: once the package exists on npm, switch to npm trusted publishing (OIDC) in the package
  settings and drop `NPM_TOKEN`.

## 0.1.0: the first publish

`package.json` is already at `0.1.0` and `CHANGELOG.md` already has the 0.1.0 entry, so no changeset is
needed for it. Two options:

1. **Automated.** After `NPM_TOKEN` is set, push/merge to `next`. The workflow finds no pending
   changesets and `0.1.0` missing on npm, and publishes it.
2. **Manual first publish** (to claim the name or to test locally):

   ```sh
   cd next
   pnpm install --frozen-lockfile
   npm login                      # 2FA prompt
   npm pack --dry-run             # inspect the file list
   npm publish --access public    # prepublishOnly runs build + tests; add --provenance only from CI
   git tag xlsxkit@0.1.0 && git push origin xlsxkit@0.1.0
   ```

   `--provenance` only works from a supported CI provider, so a local publish has none.

After 0.1.0, every release goes through changesets as above.

## Deprecating or fixing a bad release

`npm deprecate xlsxkit@0.1.1 "reason"`; publish a fixed version through the normal flow. Unpublishing is
only possible within 72 hours and is discouraged.
