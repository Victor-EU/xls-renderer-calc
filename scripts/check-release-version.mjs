/**
 * The tag must be the version, in every manifest, before anything is published.
 *
 * A mis-tagged release cannot be taken back. npm versions are immutable, so a
 * `v0.3.0` tag that publishes 0.2.0 burns a version number and leaves a
 * permanently wrong tag pointing at the wrong tree. This costs a second and runs
 * before the publish steps.
 *
 * It also checks the thing that is easy to forget: the two packages are released
 * together, and the preview's dependency range has to admit the engine being
 * published beside it. Bumping the pair to 0.3.0 while the preview still asks
 * for `^0.2.0` produces a release that installs the *old* engine — green tests,
 * wrong library, and no error anywhere.
 *
 *   TAG=v0.3.0 node scripts/check-release-version.mjs
 */

import { readFileSync } from 'node:fs';

const read = (p) => JSON.parse(readFileSync(new URL(`../${p}/package.json`, import.meta.url), 'utf8'));

const tag = (process.env.TAG ?? '').replace(/^v/, '');
if (!/^\d+\.\d+\.\d+(-[\w.]+)?$/.test(tag)) {
  fail(`TAG is "${process.env.TAG}", which is not a version. Tag releases as v0.3.0.`);
}

const PACKAGES = ['packages/formula-engine', 'packages/xlsx-preview'];
for (const pkg of PACKAGES) {
  const { name, version } = read(pkg);
  if (version !== tag) fail(`${name} is ${version}, but the tag says ${tag}`);
  console.log(`  ok  ${name} ${version}`);
}

// `^0.2.0` admits 0.2.x and, while the major is 0, nothing else — npm treats a
// caret on a 0.x version as locking the minor. So the range's minor has to match
// the version being published exactly, not merely be less than it.
const preview = read('packages/xlsx-preview');
const range = preview.dependencies['@xlscalc/formula-engine'];
const bare = range.replace(/^[\^~>=\s]*/, '');
const [tagMajor, tagMinor] = tag.split('.').map(Number);
const [depMajor, depMinor] = bare.split('.').map(Number);
if (depMajor !== tagMajor || (tagMajor === 0 && depMinor !== tagMinor) || depMinor > tagMinor) {
  fail(`xlsx-preview depends on "${range}", which does not admit the ${tag} engine it ships with`);
}
console.log(`  ok  xlsx-preview depends on "${range}", satisfied by ${tag}`);

function fail(message) {
  console.error(`::error::${message}`);
  process.exit(1);
}
