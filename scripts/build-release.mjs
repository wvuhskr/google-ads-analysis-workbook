#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const VERSION = 'v1.0.0';
const DEFAULT_OUTPUT = path.join(ROOT, 'release-assets', VERSION);
const RELEASE_INPUTS = [
  'google-ads-analysis-workbook.js',
  path.join('tools', 'sanitize-downloaded-xlsx.js'),
  path.join('docs', 'release', 'QUICKSTART.md'),
  path.join('docs', 'release', 'DATA-HANDLING.md'),
  'LICENSE',
];
const ASSETS = [
  {
    source: 'google-ads-analysis-workbook.js',
    releaseName: `google-ads-analysis-workbook-${VERSION}.js`,
  },
  {
    source: path.join('tools', 'sanitize-downloaded-xlsx.js'),
    releaseName: `sanitize-downloaded-xlsx-${VERSION}.js`,
  },
];
const ARCHIVE_MEMBERS = [
  `google-ads-analysis-workbook-${VERSION}.js`,
  `sanitize-downloaded-xlsx-${VERSION}.js`,
  'QUICKSTART.md',
  'DATA-HANDLING.md',
  'LICENSE',
];
const CHECKSUMMED_ASSETS = [
  `google-ads-analysis-workbook-${VERSION}.js`,
  `sanitize-downloaded-xlsx-${VERSION}.js`,
  `google-ads-analysis-workbook-${VERSION}.zip`,
];
const FIXED_MTIME = new Date(Date.UTC(2000, 0, 1, 0, 0, 0));

function fail(message) {
  throw new Error(message);
}

function sha256File(filePath) {
  return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

function git(args, options = {}) {
  try {
    return execFileSync('git', args, {
      cwd: ROOT,
      encoding: options.encoding === undefined ? 'utf8' : options.encoding,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
  } catch (error) {
    const detail = String(error.stderr || error.message || '').trim();
    fail(`Git release check failed${detail ? `: ${detail}` : '.'}`);
  }
}

function assertCleanTrackedTree() {
  const repositoryRoot = path.resolve(String(git(['rev-parse', '--show-toplevel'])).trim());
  if (repositoryRoot !== ROOT) {
    fail(`Release builder must run from its own repository root: ${ROOT}`);
  }
  const status = String(git(['status', '--porcelain=v1', '--untracked-files=all']));
  if (status.trim()) {
    fail('Release source must be clean; commit or remove all tracked and untracked changes first.');
  }
}

function assertReleaseInputsMatchHead() {
  const snapshots = new Map();
  for (const relativePath of RELEASE_INPUTS) {
    git(['ls-files', '--error-unmatch', '--', relativePath]);
    const worktreeBytes = fs.readFileSync(path.join(ROOT, relativePath));
    const headBytes = git(['show', `HEAD:${relativePath}`], { encoding: null });
    if (!worktreeBytes.equals(headBytes)) {
      fail(`Release input bytes differ from HEAD: ${relativePath}`);
    }
    snapshots.set(relativePath, headBytes);
  }
  return snapshots;
}

function assertAnnotatedReleaseTagAtHead() {
  const tagRef = `refs/tags/${VERSION}`;
  const objectType = String(git(['cat-file', '-t', tagRef])).trim();
  if (objectType !== 'tag') {
    fail(`${VERSION} must exist as an annotated tag before a release build.`);
  }
  const tagObject = String(git(['cat-file', '-p', tagRef]));
  const typeMatch = /^type ([^\n]+)$/m.exec(tagObject);
  const objectMatch = /^object ([0-9a-f]+)$/m.exec(tagObject);
  if (!typeMatch || typeMatch[1] !== 'commit' || !objectMatch) {
    fail(`${VERSION} must be an annotated tag that directly targets a commit.`);
  }
  const headCommit = String(git(['rev-parse', 'HEAD^{commit}'])).trim();
  const taggedCommit = objectMatch[1];
  if (headCommit !== taggedCommit) {
    fail(`${VERSION} must point to HEAD before a release build.`);
  }
}

function assertReleaseSource(options = {}) {
  assertCleanTrackedTree();
  const snapshots = assertReleaseInputsMatchHead();
  if (!options.allowUntagged) assertAnnotatedReleaseTagAtHead();
  return snapshots;
}

function assertEmptyDestination(destination) {
  if (!fs.existsSync(destination)) return;
  if (!fs.statSync(destination).isDirectory()) fail(`Release destination is not a directory: ${destination}`);
  if (fs.readdirSync(destination).length !== 0) {
    fail(`Release destination must be empty: ${destination}`);
  }
}

function copyInput(snapshots, sourceRelativePath, destinationPath) {
  const sourceBytes = snapshots.get(sourceRelativePath);
  if (!sourceBytes) fail(`Required release input is missing from HEAD: ${sourceRelativePath}`);
  fs.writeFileSync(destinationPath, sourceBytes);
  fs.utimesSync(destinationPath, FIXED_MTIME, FIXED_MTIME);
}

function build(outputDirectory = DEFAULT_OUTPUT, options = {}) {
  const snapshots = assertReleaseSource(options);
  const destination = path.resolve(outputDirectory);
  assertEmptyDestination(destination);
  fs.mkdirSync(destination, { recursive: true });
  const stage = fs.mkdtempSync(path.join(os.tmpdir(), 'google-ads-analysis-workbook-release-'));
  try {
    for (const asset of ASSETS) copyInput(snapshots, asset.source, path.join(stage, asset.releaseName));
    copyInput(snapshots, path.join('docs', 'release', 'QUICKSTART.md'), path.join(stage, 'QUICKSTART.md'));
    copyInput(snapshots, path.join('docs', 'release', 'DATA-HANDLING.md'), path.join(stage, 'DATA-HANDLING.md'));
    copyInput(snapshots, 'LICENSE', path.join(stage, 'LICENSE'));

    const archivePath = path.join(destination, `google-ads-analysis-workbook-${VERSION}.zip`);
    execFileSync('zip', ['-X', '-q', archivePath, ...ARCHIVE_MEMBERS], {
      cwd: stage,
      env: { ...process.env, TZ: 'UTC' },
      stdio: 'pipe',
    });
    for (const asset of ASSETS) fs.copyFileSync(path.join(stage, asset.releaseName), path.join(destination, asset.releaseName));

    const sums = CHECKSUMMED_ASSETS
      .map((name) => `${sha256File(path.join(destination, name))}  ${name}`)
      .join('\n') + '\n';
    fs.writeFileSync(path.join(destination, 'SHA256SUMS.txt'), sums, 'utf8');
  } finally {
    fs.rmSync(stage, { recursive: true, force: true });
  }
}

function parseArguments(args) {
  let allowUntagged = false;
  let parsingOptions = true;
  const positional = [];
  for (const argument of args) {
    if (parsingOptions && argument === '--') {
      parsingOptions = false;
    } else if (parsingOptions && argument === '--allow-untagged') {
      if (allowUntagged) fail('Duplicate option: --allow-untagged');
      allowUntagged = true;
    } else if (parsingOptions && argument.startsWith('-')) {
      fail(`Unknown option: ${argument}`);
    } else {
      positional.push(argument);
    }
  }
  if (positional.length > 1) {
    fail('Usage: node scripts/build-release.mjs [--allow-untagged] [--] [output-directory]');
  }
  return { allowUntagged, outputDirectory: positional[0] || DEFAULT_OUTPUT };
}

function main() {
  const options = parseArguments(process.argv.slice(2));
  build(options.outputDirectory, { allowUntagged: options.allowUntagged });
}

try {
  main();
} catch (error) {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
}

export { ARCHIVE_MEMBERS, CHECKSUMMED_ASSETS, RELEASE_INPUTS, build, parseArguments };
