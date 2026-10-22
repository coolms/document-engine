#!/usr/bin/env node
// The Node versions this package declares are the ones its CI runs and its dependencies accept.
//
//   node engines-check.mjs [--root DIR] [--workflow FILE | --majors 22,24,26]
//
// The majors of package.json's engines.node are exactly the Node lines CI runs: the one `node: [...]` matrix list in
// the workflow (default: <root>/.github/workflows/ci.yml), or the --majors given by a caller that holds the matrix
// itself. Within each of them, the lowest version engines.node admits is no lower than the highest floor any package
// in package-lock.json declares for that major, and every package admits that major at all. The lockfile's own copy
// of the root's engines.node is the same text as package.json's. A package with no dependencies and no lockfile has
// no floor but the major's x.0.0.
//
// Last line: `engines CLEAR: ...`, `engines FOUND: ...` or `engines UNEVALUABLE: ...`.
// Exit: 0 CLEAR, 1 FOUND, 2 UNEVALUABLE.

import { existsSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';

function option(name) {
  const at = process.argv.indexOf(name);
  return at === -1 ? null : process.argv[at + 1] ?? '';
}

const root = resolve(option('--root') ?? join(import.meta.dirname, '..'));
// Majors above the highest one CI runs that engines.node is checked not to admit: an open range such as ">=26"
// admits every one of them.
const MAJORS_ABOVE = 100;

function finish(line, code) {
  process.stdout.write(line + '\n');
  process.exit(code);
}

// semver from the package's own install, else the copy npm ships with (a package with no dependencies installs none):
// npm sits in <prefix>/lib/node_modules/npm for some prefix above the node binary.
const places = [join(root, 'package.json')];
for (let dir = dirname(process.execPath); dir !== dirname(dir); dir = dirname(dir)) {
  places.push(join(dir, 'lib', 'node_modules', 'npm', 'package.json'));
}
let semver;
for (const from of places) {
  try {
    semver = createRequire(from)('semver');
    break;
  } catch {
    // the next place
  }
}
if (!semver) finish('engines UNEVALUABLE: semver is found neither in the package (run npm ci) nor in npm', 2);

function readJson(rel) {
  try {
    return JSON.parse(readFileSync(join(root, rel), 'utf8'));
  } catch (error) {
    finish(`engines UNEVALUABLE: ${rel} cannot be read: ${error.message}`, 2);
  }
}

/** The Node lines CI runs: --majors, else the one `node: [...]` list in the workflow. */
function matrixMajors() {
  let source;
  let list;
  const given = option('--majors');
  if (given !== null) {
    source = '--majors';
    list = given.replace(/^\s*\[|\]\s*$/g, '');
  } else {
    const workflow = resolve(option('--workflow') ?? join(root, '.github/workflows/ci.yml'));
    source = workflow;
    let text;
    try {
      text = readFileSync(workflow, 'utf8');
    } catch (error) {
      finish(`engines UNEVALUABLE: ${workflow} cannot be read: ${error.message}`, 2);
    }
    const lists = [...text.matchAll(/^\s+node:\s*\[([^\]]*)\]\s*$/gm)];
    if (lists.length !== 1) finish(`engines UNEVALUABLE: ${workflow} has ${lists.length} matrix lists of Node lines, not one`, 2);
    list = lists[0][1];
  }
  const majors = list.split(',').map((item) => item.trim().replace(/^["']|["']$/g, ''));
  if (majors.length === 0 || majors.some((major) => !/^\d+$/.test(major))) {
    finish(`engines UNEVALUABLE: ${source} lists "${list.trim()}", not whole Node majors`, 2);
  }
  return majors.map(Number).sort((a, b) => a - b);
}

/**
 * The lowest version a range admits within one major, or null when it admits none there; undefined when the range
 * does not parse. The lowest admitted version is the major's x.0.0 or one of the range's own lower bounds (the next
 * patch for a ">"), so those are the versions tried. (semver's minVersion is not used: some copies return null for a
 * comparator set with two lower bounds.)
 */
function floorIn(range, major) {
  let parsed;
  try {
    parsed = new semver.Range(range);
  } catch {
    return undefined;
  }
  const candidates = [semver.parse(`${major}.0.0`)];
  for (const set of parsed.set) {
    for (const comparator of set) {
      if (!comparator.semver || comparator.semver.major !== major) continue;
      const bound = semver.parse(`${comparator.semver.major}.${comparator.semver.minor}.${comparator.semver.patch}`);
      if (comparator.operator === '>=' || comparator.operator === '') candidates.push(bound);
      if (comparator.operator === '>') candidates.push(semver.parse(semver.inc(bound, 'patch')));
    }
  }
  candidates.sort(semver.compare);
  return candidates.find((version) => parsed.test(version)) ?? null;
}

const pkg = readJson('package.json');
const engines = pkg.engines?.node;
if (typeof engines !== 'string' || engines.trim() === '') finish('engines UNEVALUABLE: package.json declares no engines.node', 2);
if (floorIn(engines, 0) === undefined) finish(`engines FOUND: package.json's engines.node "${engines}" is not a version range`, 1);
const majors = matrixMajors();

const findings = [];
let packages;
if (existsSync(join(root, 'package-lock.json'))) {
  const lock = readJson('package-lock.json');
  packages = Object.entries(lock.packages ?? {});
  if (packages.length === 0) finish('engines UNEVALUABLE: package-lock.json lists no packages', 2);
  const lockEngines = lock.packages['']?.engines?.node;
  if (lockEngines !== engines) {
    findings.push(`package-lock.json's root engines.node is "${lockEngines ?? '(none)'}", package.json's is "${engines}"`);
  }
} else {
  const kinds = ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies'];
  const declared = kinds.filter((kind) => pkg[kind] && Object.keys(pkg[kind]).length > 0);
  if (declared.length > 0) finish(`engines UNEVALUABLE: no package-lock.json to read the floors of the ${declared.join(', ')} from`, 2);
  packages = [];
}

const named = [];
for (let major = 0; major <= majors[majors.length - 1] + MAJORS_ABOVE; major++) {
  if (floorIn(engines, major) !== null) named.push(major);
}
const extra = named.filter((major) => !majors.includes(major));
const missing = majors.filter((major) => !named.includes(major));
if (extra.length > 0) {
  const shown = extra.length > 3 ? `${extra.slice(0, 3).join(', ')}, ... (${extra.length} majors)` : extra.join(', ');
  findings.push(`engines.node "${engines}" admits Node ${shown}, which CI (${majors.join(', ')}) does not run`);
}
if (missing.length > 0) {
  findings.push(`CI runs Node ${missing.join(', ')}, which engines.node "${engines}" does not admit`);
}

// For each major CI runs: the highest floor a package declares, and which packages declare it.
const read = packages.filter(([path, entry]) => path !== '' && typeof entry.engines?.node === 'string').length;
const floors = [];
for (const major of majors) {
  let highest = null;
  let by = [];
  for (const [path, entry] of packages) {
    const range = entry.engines?.node;
    if (path === '' || typeof range !== 'string') continue;
    const name = `${path.replace(/^.*node_modules\//, '')}@${entry.version ?? '(no version)'}`;
    const floor = floorIn(range, major);
    if (floor === undefined) {
      findings.push(`${name} declares engines.node "${range}", which is not a version range`);
      continue;
    }
    if (floor === null) {
      findings.push(`${name} declares engines.node "${range}", which admits no Node ${major}, a line CI runs`);
      continue;
    }
    if (highest === null || semver.gt(floor, highest)) {
      highest = floor;
      by = [name];
    } else if (semver.eq(floor, highest) && !by.includes(name)) {
      by.push(name);
    }
  }
  const required = highest ?? semver.parse(`${major}.0.0`);
  const declared = floorIn(engines, major);
  floors.push({ major, required });
  if (declared && semver.lt(declared, required)) {
    const shown = by.slice(0, 3).join(', ') + (by.length > 3 ? `, ... (${by.length} packages)` : '');
    findings.push(`engines.node "${engines}" admits Node ${declared.version}, below ${required.version}, the floor ${shown} declare${by.length === 1 ? 's' : ''} for Node ${major}`);
  }
}

const computed = floors.map(({ required }) => `^${required.version}`).join(' || ');
for (const finding of findings) process.stdout.write(`engines: ${finding}\n`);
if (findings.length > 0) {
  finish(`engines FOUND: ${findings.length} finding(s); the lockfile's floors for CI's Node ${majors.join(', ')} give "${computed}"`, 1);
}
const from = packages.length === 0
  ? 'no lockfile and no dependencies, so each from its first release'
  : `each from the lockfile's floor (${floors.map(({ required }) => required.version).join(', ')}); ${read} package${read === 1 ? '' : 's'} with engines.node read`;
finish(`engines CLEAR: "${engines}" admits exactly CI's Node ${majors.join(', ')}, ${from}`, 0);
