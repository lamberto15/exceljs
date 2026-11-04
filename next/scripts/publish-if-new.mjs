// Publishes the package with the npm CLI (trusted publishing / OIDC needs npm >= 11.5.1) only when
// the current version is not on the registry yet, then creates the git tag via changesets.
// Run by the changesets action in .github/workflows/next-release.yml; safe to re-run.
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const { name, version } = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
const run = (cmd, args) => execFileSync(cmd, args, { stdio: 'inherit' });

let published = false;
try {
  execFileSync('npm', ['view', `${name}@${version}`, 'version'], { stdio: 'pipe' });
  published = true;
} catch {
  // npm view exits non-zero (E404) when the version does not exist
}
if (published) {
  console.log(`${name}@${version} is already on npm; nothing to publish.`);
} else {
  run('npm', ['publish', '--access', 'public']);
  run('npx', ['changeset', 'tag']);
  // changesets/action reads this line to create the GitHub release
  console.log(`New tag: ${name}@${version}`);
}
