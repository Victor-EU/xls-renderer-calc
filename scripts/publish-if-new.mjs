/**
 * Publish a workspace, unless that exact version is already on the registry.
 *
 * npm versions are immutable: republishing one fails with E403. That matters
 * because this release publishes two packages in sequence, and anything that
 * interrupts it between them — a network blip, a revoked token, a runner dying —
 * leaves the engine published and the preview not. Re-running the job would then
 * fail on the *engine* step and never reach the package that actually needs
 * publishing, so the only way out would be burning a version number on a release
 * whose contents nobody changed.
 *
 * Treating "already there" as done rather than as an error makes the job safe to
 * re-run, which is the property you want from the step you are least able to
 * predict. It cannot mask a real problem: the version guard has already checked
 * that the tag and the manifests agree, so a version that exists is one this
 * same tag published.
 *
 *   node scripts/publish-if-new.mjs packages/formula-engine
 */

import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const dir = process.argv[2];
if (!dir) {
  console.error('::error::usage: publish-if-new.mjs <workspace-dir>');
  process.exit(1);
}

const { name, version } = JSON.parse(readFileSync(new URL(`../${dir}/package.json`, import.meta.url), 'utf8'));

const npm = (args) => execFileSync('npm', args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });

let published = false;
try {
  published = npm(['view', `${name}@${version}`, 'version']).trim() === version;
} catch {
  // A 404 is the ordinary case for a first publish, and for a scope that has no
  // packages yet. Anything else worth knowing about will surface on publish.
  published = false;
}

if (published) {
  console.log(`  ${name}@${version} is already on the registry — nothing to do`);
  process.exit(0);
}

console.log(`  publishing ${name}@${version}`);
execFileSync('npm', ['publish', '-w', name, '--access', 'public', '--provenance'], { stdio: 'inherit' });
