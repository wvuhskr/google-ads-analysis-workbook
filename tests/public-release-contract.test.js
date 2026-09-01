'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const { execFileSync, spawnSync } = require('node:child_process');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');

const PACKAGE_ROOT = path.resolve(__dirname, '..');
const EXPORTER = path.join(PACKAGE_ROOT, 'google-ads-analysis-workbook.js');
const SANITIZER = path.join(PACKAGE_ROOT, 'tools', 'sanitize-downloaded-xlsx.js');
const BUILD_RELEASE = path.join(PACKAGE_ROOT, 'scripts', 'build-release.mjs');
const AUDIT_PUBLIC_PACKAGE = path.join(PACKAGE_ROOT, 'scripts', 'audit-public-package.mjs');
const CANONICAL_REPOSITORY = 'https://github.com/wvuhskr/google-ads-analysis-workbook';
const APPROVED_PREVIEW_SHA256 = '828cdbd700d568c87eab874a21dafea17fe7878ee9c18d33daa9ca334a473392';
const RELEASE_CANDIDATE_LANGUAGE = new RegExp(
  ['R' + 'C15', 'release' + ' candidate'].join('|'),
  'i',
);

function loadExporter(filePath) {
  const source = fs.readFileSync(filePath, 'utf8');
  const context = {
    console,
    Date,
    JSON,
    Math,
    Object,
    RegExp,
    String,
    Number,
    Boolean,
    Array,
    Error,
    Set,
    Map,
  };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: filePath });
  return { context, source };
}

function temporaryDirectory(prefix) {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

function sha256File(filePath) {
  return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

function runNode(script, args, options = {}) {
  return spawnSync(process.execPath, [script, ...args], {
    cwd: PACKAGE_ROOT,
    encoding: 'utf8',
    env: { ...process.env, ...options.env },
    ...options,
  });
}

function workbookUrl() {
  return ['https://docs.google.com', 'spreadsheets', 'd', 'review-fixture-sheet', 'edit'].join('/');
}

function customerId() {
  return ['246', '801', '3579'].join('-');
}

function unformattedCustomerId() {
  return ['246', '801', '3579'].join('');
}

function githubToken() {
  return ['gh', 'p_', '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'].join('');
}

function modernGithubToken() {
  return ['github', '_pat_', '11AA22BB33CC44DD55EE_', '0123456789abcdefghijklmnopqrstuvwxyz'].join('');
}

function googleApiKey() {
  return ['AI', 'za', 'SyA1234567890abcdefghijklmnopqrstuv'].join('');
}

function privateKeyMarker() {
  return ['-----BEGIN ', 'PRIVATE KEY-----'].join('');
}

function archiveMember(archivePath, member) {
  return execFileSync('unzip', ['-p', archivePath, member]);
}

function writeChecksums(directory) {
  const names = [
    'google-ads-analysis-workbook-v1.0.0.js',
    'sanitize-downloaded-xlsx-v1.0.0.js',
    'google-ads-analysis-workbook-v1.0.0.zip',
  ];
  fs.writeFileSync(
    path.join(directory, 'SHA256SUMS.txt'),
    names.map((name) => `${sha256File(path.join(directory, name))}  ${name}`).join('\n') + '\n',
  );
}

function runGit(cwd, args) {
  return execFileSync('git', args, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
}

function createReleaseGitFixture(options = {}) {
  const fixture = temporaryDirectory('google-ads-analysis-workbook-git-fixture-');
  const inputs = [
    'google-ads-analysis-workbook.js',
    path.join('tools', 'sanitize-downloaded-xlsx.js'),
    path.join('docs', 'release', 'QUICKSTART.md'),
    path.join('docs', 'release', 'DATA-HANDLING.md'),
    'LICENSE',
    path.join('scripts', 'build-release.mjs'),
  ];
  for (const relativePath of inputs) {
    const destination = path.join(fixture, relativePath);
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    fs.copyFileSync(path.join(PACKAGE_ROOT, relativePath), destination);
  }
  runGit(fixture, ['init', '-q', '-b', 'main']);
  runGit(fixture, ['config', 'user.name', 'Release Test']);
  runGit(fixture, ['config', 'user.email', 'release-test@example.invalid']);
  runGit(fixture, ['add', '--', ...inputs]);
  runGit(fixture, ['commit', '-q', '-m', 'fixture']);
  if (options.tag === 'annotated') {
    runGit(fixture, ['tag', '-a', 'v1.0.0', '-m', 'fixture v1.0.0']);
  } else if (options.tag === 'lightweight') {
    runGit(fixture, ['tag', 'v1.0.0']);
  }
  return fixture;
}

function listPublicFiles(directory, root = directory) {
  const files = [];
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      if (!new Set(['.git', 'node_modules', 'release-assets']).has(entry.name)) {
        files.push(...listPublicFiles(path.join(directory, entry.name), root));
      }
    } else if (entry.isFile()) {
      files.push(path.relative(root, path.join(directory, entry.name)));
    }
  }
  return files.sort();
}

function createCompletePublicGitFixture(options = {}) {
  const fixture = temporaryDirectory('google-ads-analysis-workbook-public-fixture-');
  const tracked = listPublicFiles(PACKAGE_ROOT);
  for (const relativePath of tracked) {
    const destination = path.join(fixture, relativePath);
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    fs.copyFileSync(path.join(PACKAGE_ROOT, relativePath), destination);
  }
  runGit(fixture, ['init', '-q', '-b', 'main']);
  runGit(fixture, ['config', 'user.name', 'Audit Test']);
  runGit(fixture, ['config', 'user.email', 'audit-test@example.invalid']);
  runGit(fixture, ['add', '--', ...tracked]);
  runGit(fixture, ['commit', '-q', '-m', 'public fixture']);
  if (options.tag === 'annotated') {
    runGit(fixture, ['tag', '-a', 'v1.0.0', '-m', 'fixture v1.0.0']);
  }
  return fixture;
}

function runFixtureBuild(fixture, args, options = {}) {
  return spawnSync(process.execPath, [path.join(fixture, 'scripts', 'build-release.mjs'), ...args], {
    cwd: fixture,
    encoding: 'utf8',
    env: { ...process.env, ...options.env },
  });
}

function withTamperedRelease(callback) {
  const releaseDirectory = temporaryDirectory('google-ads-analysis-workbook-tampered-release-');
  const fixture = createCompletePublicGitFixture({ tag: 'annotated' });
  const build = runFixtureBuild(fixture, [releaseDirectory]);
  assert.equal(build.status, 0, build.stderr);
  callback(releaseDirectory, fixture);
}

test('stable public source exposes the verified v1.0.0 contract', () => {
  const source = fs.readFileSync(EXPORTER, 'utf8');
  const sanitizerSource = fs.readFileSync(SANITIZER, 'utf8');

  assert.match(source, /Google Ads Analysis Workbook/);
  assert.match(source, /SPDX-License-Identifier: MIT/);
  assert.match(sanitizerSource, /SPDX-License-Identifier: MIT/);
  assert.doesNotMatch(source, RELEASE_CANDIDATE_LANGUAGE);
  assert.doesNotMatch(sanitizerSource, RELEASE_CANDIDATE_LANGUAGE);
  assert.doesNotMatch(source, /UrlFetchApp|ScriptApp\.getOAuthToken|DriveApp/);
  assert.doesNotMatch(source, /AUTOMATIC_XLSX|exportSanitizedXlsx|RUN_MODE/);
  assert.doesNotMatch(sanitizerSource, /https?\.request|\bfetch\s*\(/i);

  const loaded = loadExporter(EXPORTER);
  assert.equal(loaded.context.VERSION, 'v1.0.0');
  assert.equal(loaded.context.OWNER_KEY, 'google-ads-analysis-workbook');
  assert.equal(loaded.context.OUTPUT_SCHEMA_VERSION, 9);
  assert.equal(loaded.context.RUNTIME_CONTRACT_VERSION, 10);
  const manifest = loaded.context.getManifestDefinition_();
  assert.equal(manifest.length, 39);
  assert.equal(loaded.context.preferredTabOrder_(manifest).length, 41);

  const sanitizer = require(SANITIZER);
  assert.equal(typeof sanitizer.runCli, 'function');
  assert.equal(typeof sanitizer.sanitizeWorkbookEntries, 'function');
});

test('release builder emits only the approved assets and a verifiable flat archive', () => {
  // Break caught: an allowlist or checksum regression could ship extra files,
  // omit a required member, or record hashes that do not match the assets.
  const outputDirectory = temporaryDirectory('google-ads-analysis-workbook-release-');
  const fixture = createReleaseGitFixture({ tag: 'annotated' });
  const result = runFixtureBuild(fixture, [outputDirectory]);
  assert.equal(result.status, 0, result.stderr);

  const files = fs.readdirSync(outputDirectory).sort();
  assert.deepEqual(files, [
    'SHA256SUMS.txt',
    'google-ads-analysis-workbook-v1.0.0.js',
    'google-ads-analysis-workbook-v1.0.0.zip',
    'sanitize-downloaded-xlsx-v1.0.0.js',
  ]);

  const archivePath = path.join(outputDirectory, 'google-ads-analysis-workbook-v1.0.0.zip');
  const members = execFileSync('unzip', ['-Z', '-1', archivePath], { encoding: 'utf8' })
    .trim()
    .split('\n')
    .filter(Boolean)
    .sort();
  assert.deepEqual(members, [
    'DATA-HANDLING.md',
    'LICENSE',
    'QUICKSTART.md',
    'google-ads-analysis-workbook-v1.0.0.js',
    'sanitize-downloaded-xlsx-v1.0.0.js',
  ]);
  for (const member of members) {
    assert.doesNotMatch(member, /[\\/]/, `${member} must be a flat archive member`);
    assert.match(member, /\.(?:js|md)$|^LICENSE$/, `${member} has an unexpected extension`);
  }
  const archiveQuickstart = archiveMember(archivePath, 'QUICKSTART.md').toString('utf8');
  const archiveDataHandling = archiveMember(archivePath, 'DATA-HANDLING.md').toString('utf8');
  assert.match(archiveQuickstart, /google-ads-analysis-workbook-v1\.0\.0\.js/);
  assert.match(archiveQuickstart, /sanitize-downloaded-xlsx-v1\.0\.0\.js/);
  assert.match(archiveQuickstart, /\[Data handling\]\(DATA-HANDLING\.md\)/);
  assert.doesNotMatch(archiveQuickstart, /\.\.\/|tools\//);
  assert.doesNotMatch(archiveDataHandling, /\]\(SECURITY\.md\)/);

  const sums = fs.readFileSync(path.join(outputDirectory, 'SHA256SUMS.txt'), 'utf8')
    .trim()
    .split('\n');
  assert.equal(sums.length, 3);
  for (const line of sums) {
    const match = /^([a-f0-9]{64})  ([^\\/]+)$/.exec(line);
    assert.ok(match, `invalid checksum line: ${line}`);
    assert.equal(match[1], sha256File(path.join(outputDirectory, match[2])));
  }
});

test('release builder refuses to overwrite a nonempty destination', () => {
  // Break caught: accepting an existing release directory can silently mix
  // stale files with newly built assets.
  const outputDirectory = temporaryDirectory('google-ads-analysis-workbook-dirty-release-');
  fs.writeFileSync(path.join(outputDirectory, 'existing.txt'), 'do not overwrite');
  const fixture = createReleaseGitFixture({ tag: 'annotated' });
  const result = runFixtureBuild(fixture, [outputDirectory]);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /nonempty|empty|refus/i);
});

test('strict release builder requires a clean annotated v1.0.0 tag at HEAD', () => {
  const clean = createReleaseGitFixture({ tag: 'annotated' });
  assert.equal(runFixtureBuild(clean, [temporaryDirectory('strict-release-')]).status, 0);

  const missing = createReleaseGitFixture();
  const missingResult = runFixtureBuild(missing, [temporaryDirectory('missing-tag-release-')]);
  assert.notEqual(missingResult.status, 0);
  assert.match(`${missingResult.stdout}\n${missingResult.stderr}`, /annotated|tag/i);
  assert.equal(
    runFixtureBuild(missing, ['--allow-untagged', temporaryDirectory('ci-release-')]).status,
    0,
  );

  const lightweight = createReleaseGitFixture({ tag: 'lightweight' });
  const lightweightResult = runFixtureBuild(lightweight, [temporaryDirectory('lightweight-release-')]);
  assert.notEqual(lightweightResult.status, 0);
  assert.match(`${lightweightResult.stdout}\n${lightweightResult.stderr}`, /annotated|tag/i);

  const nested = createReleaseGitFixture();
  runGit(nested, ['tag', '-a', 'inner-release', '-m', 'inner release']);
  runGit(nested, ['tag', '-a', 'v1.0.0', 'inner-release', '-m', 'nested release']);
  const nestedResult = runFixtureBuild(nested, [temporaryDirectory('nested-tag-release-')]);
  assert.notEqual(nestedResult.status, 0);
  assert.match(`${nestedResult.stdout}\n${nestedResult.stderr}`, /directly|commit|tag/i);

  const behind = createReleaseGitFixture({ tag: 'annotated' });
  fs.writeFileSync(path.join(behind, 'tracked-note.txt'), 'second commit\n');
  runGit(behind, ['add', 'tracked-note.txt']);
  runGit(behind, ['commit', '-q', '-m', 'move head']);
  const behindResult = runFixtureBuild(behind, [temporaryDirectory('behind-tag-release-')]);
  assert.notEqual(behindResult.status, 0);
  assert.match(`${behindResult.stdout}\n${behindResult.stderr}`, /HEAD|tag/i);
});

test('release builder refuses dirty source in strict and allow-untagged modes', () => {
  for (const dirtyKind of ['unstaged', 'staged', 'untracked']) {
    const fixture = createReleaseGitFixture({ tag: 'annotated' });
    if (dirtyKind === 'unstaged' || dirtyKind === 'staged') {
      fs.appendFileSync(path.join(fixture, 'LICENSE'), '\nmodified\n');
      if (dirtyKind === 'staged') runGit(fixture, ['add', 'LICENSE']);
    } else {
      fs.writeFileSync(path.join(fixture, 'unexpected.txt'), 'untracked\n');
    }
    for (const args of [[], ['--allow-untagged']]) {
      const result = runFixtureBuild(
        fixture,
        [...args, temporaryDirectory(`dirty-${dirtyKind}-release-`)],
      );
      assert.notEqual(result.status, 0, `${dirtyKind} ${args.join(' ')}`);
      assert.match(`${result.stdout}\n${result.stderr}`, /clean|dirty|tracked|untracked/i);
    }
  }
});

test('release builder compares every release input with HEAD even when status is hidden', () => {
  const fixture = createReleaseGitFixture({ tag: 'annotated' });
  const releaseInputs = [
    'google-ads-analysis-workbook.js',
    path.join('tools', 'sanitize-downloaded-xlsx.js'),
    path.join('docs', 'release', 'QUICKSTART.md'),
    path.join('docs', 'release', 'DATA-HANDLING.md'),
    'LICENSE',
  ];
  for (const relativePath of releaseInputs) {
    runGit(fixture, ['update-index', '--no-assume-unchanged', '--', ...releaseInputs]);
    runGit(fixture, ['reset', '--hard', '-q', 'HEAD']);
    fs.appendFileSync(path.join(fixture, relativePath), '\nhidden worktree substitution\n');
    runGit(fixture, ['update-index', '--assume-unchanged', '--', relativePath]);
    const result = runFixtureBuild(
      fixture,
      ['--allow-untagged', temporaryDirectory('head-mismatch-release-')],
    );
    assert.notEqual(result.status, 0, relativePath);
    assert.match(`${result.stdout}\n${result.stderr}`, /HEAD|byte|input|match/i);
  }
});

test('release builder rejects unknown options and duplicate output paths', () => {
  const fixture = createReleaseGitFixture({ tag: 'annotated' });
  for (const args of [
    ['--unsafe'],
    [temporaryDirectory('one-output-'), temporaryDirectory('two-output-')],
    ['--allow-untagged', '--allow-untagged'],
  ]) {
    const result = runFixtureBuild(fixture, args);
    assert.notEqual(result.status, 0);
    assert.match(`${result.stdout}\n${result.stderr}`, /usage|option|argument/i);
  }
});

test('public package audit rejects a candidate containing a workbook URL and customer ID', () => {
  // Break caught: an audit that only checks file names can miss live advertiser
  // identifiers embedded in otherwise harmless text files.
  const fixtureDirectory = temporaryDirectory('google-ads-analysis-workbook-audit-fixture-');
  fs.writeFileSync(
    path.join(fixtureDirectory, 'contaminated.txt'),
    `Workbook ${workbookUrl()} for ${customerId()}`,
  );
  const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixtureDirectory]);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /workbook|customer|sheet/i);
});

test('public package audit rejects retired product and release names', () => {
  // Break caught: a partially renamed release can retain stale branding in a
  // secondary document or generated asset even when the primary README is new.
  const fixture = createCompletePublicGitFixture();
  const retiredName = ['Master Account', ' Export'].join('');
  fs.appendFileSync(path.join(fixture, 'SUPPORT.md'), `\n${retiredName}\n`);
  runGit(fixture, ['add', 'SUPPORT.md']);
  runGit(fixture, ['commit', '-q', '-m', 'reintroduce retired product name']);

  const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixture]);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /retired product or release name/i);
});

test('public package audit rejects a candidate containing a GitHub access token', () => {
  // Break caught: a text-only privacy audit can still miss a usable secret.
  const fixtureDirectory = temporaryDirectory('google-ads-analysis-workbook-secret-fixture-');
  fs.writeFileSync(
    path.join(fixtureDirectory, 'config.txt'),
    `token=${githubToken()}`,
  );
  const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixtureDirectory]);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /secret|token|credential/i);
});

test('public package audit rejects modern GitHub, Google API, and private-key secrets', () => {
  const fixtures = [
    ['github-modern.txt', `token=${modernGithubToken()}\n`, /GitHub|token|secret/i],
    ['google-api.txt', `key=${googleApiKey()}\n`, /Google|API|secret|key/i],
    ['private-key.pem', `${privateKeyMarker()}\n`, /private|key|secret/i],
  ];
  for (const [name, content, expected] of fixtures) {
    const fixtureDirectory = temporaryDirectory('google-ads-analysis-workbook-modern-secret-');
    fs.writeFileSync(path.join(fixtureDirectory, name), content);
    const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixtureDirectory]);
    assert.notEqual(result.status, 0, name);
    assert.match(`${result.stdout}\n${result.stderr}`, expected);
  }
});

test('public package audit rejects unexpected binaries and a tampered approved preview', () => {
  const binaryFixture = temporaryDirectory('google-ads-analysis-workbook-binary-fixture-');
  fs.writeFileSync(path.join(binaryFixture, 'unexpected.bin'), Buffer.from([1, 0, 2, 3]));
  const binaryResult = runNode(AUDIT_PUBLIC_PACKAGE, [binaryFixture]);
  assert.notEqual(binaryResult.status, 0);
  assert.match(`${binaryResult.stdout}\n${binaryResult.stderr}`, /binary|unexpected|allowed/i);

  const previewFixture = createCompletePublicGitFixture();
  const previewPath = path.join(previewFixture, 'docs', 'assets', 'google-ads-analysis-workbook-synthetic.png');
  fs.writeFileSync(previewPath, Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    Buffer.from('tampered'),
  ]));
  runGit(previewFixture, ['add', 'docs/assets/google-ads-analysis-workbook-synthetic.png']);
  runGit(previewFixture, ['commit', '-q', '-m', 'alternate preview']);
  const previewResult = runNode(AUDIT_PUBLIC_PACKAGE, [previewFixture]);
  assert.notEqual(previewResult.status, 0);
  assert.match(`${previewResult.stdout}\n${previewResult.stderr}`, /hash|PNG|preview|approved/i);
  assert.equal(sha256File(path.join(PACKAGE_ROOT, 'docs', 'assets', 'google-ads-analysis-workbook-synthetic.png')), APPROVED_PREVIEW_SHA256);

  const textFixture = createCompletePublicGitFixture();
  fs.writeFileSync(path.join(textFixture, 'README.md'), Buffer.from([0x74, 0x65, 0x78, 0x74, 0, 0x78]));
  runGit(textFixture, ['add', 'README.md']);
  runGit(textFixture, ['commit', '-q', '-m', 'binary in expected text path']);
  const textResult = runNode(AUDIT_PUBLIC_PACKAGE, [textFixture]);
  assert.notEqual(textResult.status, 0);
  assert.match(`${textResult.stdout}\n${textResult.stderr}`, /binary|UTF-8/i);
});

test('public package audit enforces the exact tracked-file manifest', () => {
  const fixture = createCompletePublicGitFixture();
  fs.writeFileSync(path.join(fixture, 'unexpected-public-file.txt'), 'unexpected but harmless\n');
  runGit(fixture, ['add', 'unexpected-public-file.txt']);
  runGit(fixture, ['commit', '-q', '-m', 'add unexpected tracked file']);
  const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixture]);
  assert.notEqual(result.status, 0);
  assert.match(
    `${result.stdout}\n${result.stderr}`,
    /Public candidate manifest mismatch.*Unexpected: unexpected-public-file\.txt/i,
  );
});

test('public package audit independently rejects unsupported claims across public surfaces', () => {
  // Break caught: otherwise alternate wording or a less prominent public file
  // can bypass an audit that recognizes only one exact README phrase.
  const cases = [
    ['README.md', 'This exporter is API-free.'],
    ['docs/QUICKSTART.md', 'This works without any API.'],
    ['docs/release/QUICKSTART.md', 'Use this as a full restorable backup.'],
    ['package.json', 'One-click Google Ads export.'],
    ['CHANGELOG.md', 'Exports every Google Ads record.'],
    ['google-ads-analysis-workbook.js', 'Works across MCCs.'],
    ['SUPPORT.md', 'Provides a real-time Google Ads dashboard.'],
    ['DATA-HANDLING.md', 'Automatically optimizes accounts.'],
    ['docs/release/DATA-HANDLING.md', 'Anonymizes every output.'],
    ['docs/TESTING.md', 'Automatically creates an XLSX.'],
    ['CONTRIBUTING.md', 'A READY workbook authorizes unrestricted sharing.'],
  ];

  for (const [relativePath, claim] of cases) {
    const fixture = createCompletePublicGitFixture();
    const target = path.join(fixture, relativePath);
    if (relativePath === 'package.json') {
      const manifest = JSON.parse(fs.readFileSync(target, 'utf8'));
      manifest.description = claim;
      fs.writeFileSync(target, `${JSON.stringify(manifest, null, 2)}\n`);
    } else {
      fs.appendFileSync(target, `\n${claim}\n`);
    }
    runGit(fixture, ['add', relativePath]);
    runGit(fixture, ['commit', '-q', '-m', 'overclaim product scope']);

    const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixture]);
    assert.notEqual(result.status, 0, `${relativePath}: ${claim}`);
    assert.match(
      `${result.stdout}\n${result.stderr}`,
      /unsupported product scope claim/i,
      `${relativePath}: ${claim}`,
    );
  }
});

test('public package audit rejects adoption copy without the required scope boundaries', () => {
  // Break caught: privacy-safe release files can still leave users with an
  // ambiguous promise about API setup, Google Ads writes, backups, or sharing.
  const fixture = createCompletePublicGitFixture();
  fs.writeFileSync(
    path.join(fixture, 'README.md'),
    '# Google Ads Analysis Workbook\n\nA useful Google Ads spreadsheet exporter.\n',
  );
  runGit(fixture, ['add', 'README.md']);
  runGit(fixture, ['commit', '-q', '-m', 'remove product scope boundaries']);

  const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixture]);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /required|claim|scope|boundary/i);
});

test('public package audit independently requires every major adoption boundary', () => {
  // Break caught: a release can retain generic setup warnings while silently
  // dropping one material limitation that changes how users interpret output.
  const cases = [
    [/individual\s+advertiser\s+account/g, 'advertiser account', 'single-advertiser scope'],
    [
      /It is not a real-time dashboard, recurring connector, continuous sync, or\n  atomic account snapshot\./,
      'It is a point-in-time export.',
      'non-atomic scope',
    ],
    [/It is not an automated optimizer/, 'It reviews account data', 'optimizer exclusion'],
    [/It is not an exhaustive export/, 'It exports selected context', 'non-exhaustive scope'],
    [
      /It does not (?:create or distribute XLSX files automatically|automatically create or distribute XLSX files)\./,
      'XLSX files are available.',
      'manual XLSX boundary',
    ],
    [/It does not anonymize/, 'It processes', 'anonymization boundary'],
  ];

  for (const [pattern, replacement, label] of cases) {
    const fixture = createCompletePublicGitFixture();
    const target = path.join(fixture, 'README.md');
    const original = fs.readFileSync(target, 'utf8');
    const revised = original.replace(pattern, replacement);
    assert.notEqual(revised, original, label);
    fs.writeFileSync(target, revised);
    runGit(fixture, ['add', 'README.md']);
    runGit(fixture, ['commit', '-q', '-m', 'remove adoption boundary']);

    const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixture]);
    assert.notEqual(result.status, 0, label);
    assert.match(`${result.stdout}\n${result.stderr}`, /required product scope boundary/i, label);
  }
});

test('public package audit rejects a plausible unformatted customer ID', () => {
  // Break caught: identifiers can appear without display separators and still
  // expose an advertiser when they follow a customer/account ID label.
  const fixtureDirectory = temporaryDirectory('google-ads-analysis-workbook-unformatted-id-');
  fs.writeFileSync(path.join(fixtureDirectory, 'config.txt'), `customer_id=${unformattedCustomerId()}\n`);
  const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixtureDirectory]);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /customer/i);
});

test('public package audit scans a contaminating file placed under tests', () => {
  // Break caught: excluding the test tree lets a tracked-looking fixture carry
  // the very advertiser identifiers that the repository audit promises to find.
  const fixture = createCompletePublicGitFixture();
  fs.writeFileSync(
    path.join(fixture, 'tests', 'campaign-geo-regression.test.js'),
    `module.exports = ${JSON.stringify(workbookUrl())};\n`,
  );
  runGit(fixture, ['add', 'tests/campaign-geo-regression.test.js']);
  runGit(fixture, ['commit', '-q', '-m', 'contaminate test fixture']);
  const result = runNode(AUDIT_PUBLIC_PACKAGE, [fixture]);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /workbook|sheet/i);
});

test('public package audit rejects a checksum-rewritten release JS mismatch', () => {
  // Break caught: hashes can be rewritten together with a substituted release
  // file, so the release bytes must also be compared with tracked source.
  withTamperedRelease((releaseDirectory, candidateRoot) => {
    const exporter = path.join(releaseDirectory, 'google-ads-analysis-workbook-v1.0.0.js');
    fs.appendFileSync(exporter, '\n// package mismatch\n');
    writeChecksums(releaseDirectory);
    const result = runNode(AUDIT_PUBLIC_PACKAGE, [candidateRoot, releaseDirectory]);
    assert.notEqual(result.status, 0);
    assert.match(`${result.stdout}\n${result.stderr}`, /mismatch|release|source/i);
  });
});

test('public package audit rejects a checksum-rewritten ZIP member containing advertiser data', () => {
  // Break caught: checking only the archive member names and outer hash leaves
  // changed member content uninspected after an attacker rewrites SHA256SUMS.
  withTamperedRelease((releaseDirectory, candidateRoot) => {
    const stage = temporaryDirectory('google-ads-analysis-workbook-zip-tamper-');
    const archive = path.join(releaseDirectory, 'google-ads-analysis-workbook-v1.0.0.zip');
    fs.writeFileSync(path.join(stage, 'QUICKSTART.md'), `Download ${workbookUrl()}\n`);
    execFileSync('zip', ['-X', '-q', archive, 'QUICKSTART.md'], { cwd: stage });
    writeChecksums(releaseDirectory);
    const result = runNode(AUDIT_PUBLIC_PACKAGE, [candidateRoot, releaseDirectory]);
    assert.notEqual(result.status, 0);
    assert.match(`${result.stdout}\n${result.stderr}`, /workbook|member|mismatch/i);
  });
});

test('release builder produces byte-identical ZIPs across time zones', () => {
  // Break caught: zip stores local-time metadata unless the builder fixes its
  // environment, making equivalent release bytes differ by build timezone.
  const utcDirectory = temporaryDirectory('google-ads-analysis-workbook-utc-');
  const easternDirectory = temporaryDirectory('google-ads-analysis-workbook-eastern-');
  const fixture = createReleaseGitFixture({ tag: 'annotated' });
  assert.equal(runFixtureBuild(fixture, [utcDirectory], { env: { TZ: 'UTC' } }).status, 0);
  assert.equal(runFixtureBuild(fixture, [easternDirectory], { env: { TZ: 'America/New_York' } }).status, 0);
  assert.deepEqual(
    fs.readFileSync(path.join(utcDirectory, 'google-ads-analysis-workbook-v1.0.0.zip')),
    fs.readFileSync(path.join(easternDirectory, 'google-ads-analysis-workbook-v1.0.0.zip')),
  );
});

test('public adoption docs define advertiser-account scope and safe _export_info acceptance', () => {
  for (const relativePath of ['README.md', 'docs/QUICKSTART.md', 'docs/release/QUICKSTART.md']) {
    const content = fs.readFileSync(path.join(PACKAGE_ROOT, relativePath), 'utf8');
    assert.match(content, /individual advertiser account/i, relativePath);
    assert.match(content, /never.*(?:MCC|manager account)|(?:MCC|manager account).*never/i, relativePath);
    assert.match(content, /workbook_status/i, relativePath);
    assert.match(content, /READY_WITH_LIMITATIONS/, relativePath);
    assert.match(content, /review every `?LIMITED`? row/i, relativePath);
    assert.match(content, /IN_PROGRESS/, relativePath);
    assert.match(content, /NEEDS_REVIEW/, relativePath);
    assert.match(content, /COMPLETE_WITH_ERRORS/, relativePath);
    assert.match(content, /do not (?:analyze|use|share)/i, relativePath);
    assert.match(content, /next_action/i, relativePath);
  }
});

test('quickstarts explain the Preview proceed-or-stop decision in plain language', () => {
  for (const relativePath of ['docs/QUICKSTART.md', 'docs/release/QUICKSTART.md']) {
    const content = fs.readFileSync(path.join(PACKAGE_ROOT, relativePath), 'utf8');
    assert.match(content, /Proceed to (?:\*\*)?Run(?:\*\*)? only when/i, relativePath);
    assert.match(content, /SUPPORTED; rows_read=0[\s\S]{0,180}(?:normal|no matching row)/i, relativePath);
    assert.match(content, /UNSUPPORTED_OR_ERROR[\s\S]{0,350}(?:optional|LIMITED)/i, relativePath);
    assert.match(content, /Stop[\s\S]{0,180}(?:Preview|native_sheet_target|advertiser)/i, relativePath);
  }
  const packagedQuickstart = fs.readFileSync(
    path.join(PACKAGE_ROOT, 'docs', 'release', 'QUICKSTART.md'),
    'utf8',
  );
  assert.match(packagedQuickstart, /open[\s\S]{0,120}paste all (?:of (?:the|its) )?contents/i);
});

test('public metadata and offline archive docs use canonical absolute support routes', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(PACKAGE_ROOT, 'package.json'), 'utf8'));
  assert.equal(manifest.repository?.url, `${CANONICAL_REPOSITORY}.git`);
  assert.equal(manifest.homepage, `${CANONICAL_REPOSITORY}#readme`);
  assert.equal(manifest.bugs?.url, `${CANONICAL_REPOSITORY}/issues`);
  for (const relativePath of [
    'SUPPORT.md',
    'SECURITY.md',
    'DATA-HANDLING.md',
    path.join('docs', 'release', 'QUICKSTART.md'),
    path.join('docs', 'release', 'DATA-HANDLING.md'),
    path.join('.github', 'ISSUE_TEMPLATE', 'config.yml'),
  ]) {
    assert.match(
      fs.readFileSync(path.join(PACKAGE_ROOT, relativePath), 'utf8'),
      /https:\/\/github\.com\/wvuhskr\/google-ads-analysis-workbook/,
      relativePath,
    );
  }
});

test('bug form collects only safe release diagnostics', () => {
  const form = fs.readFileSync(path.join(PACKAGE_ROOT, '.github', 'ISSUE_TEMPLATE', 'bug-report.yml'), 'utf8');
  assert.match(form, /exporter version/i);
  for (const phase of ['Preview', 'Run/export', 'Resume', 'XLSX sanitizer', 'Other']) {
    assert.match(form, new RegExp(phase.replace('/', '\\/'), 'i'));
  }
  assert.match(form, /overall_status/);
  assert.match(form, /workbook_status/);
  assert.match(form, /Node\.js version/i);
  assert.match(form, /Do not include workbooks|forbid|removed advertiser data/i);
});

test('CI pins full-history gitleaks v3 and uses untagged mode only for CI builds', () => {
  const workflow = fs.readFileSync(path.join(PACKAGE_ROOT, '.github', 'workflows', 'ci.yml'), 'utf8');
  const action = 'gitleaks/gitleaks-action@e0c47f4f8be36e29cdc102c57e68cb5cbf0e8d1e';
  assert.match(workflow, new RegExp(action.replace('/', '\\/')));
  const stepStart = workflow.indexOf(action);
  const stepEnd = workflow.indexOf('\n      - ', stepStart);
  const gitleaksStep = workflow.slice(stepStart, stepEnd < 0 ? workflow.length : stepEnd);
  assert.match(gitleaksStep, /GITLEAKS_ENABLE_COMMENTS:\s*["']false["']/);
  assert.doesNotMatch(gitleaksStep, /continue-on-error/);
  assert.match(workflow, /fetch-depth:\s*0/);
  assert.match(workflow, /permissions:\s*\n\s+contents:\s*read/);
  assert.match(workflow, /npm run build:ci/);
  assert.doesNotMatch(workflow, /run:\s*npm run build:release/);
  const manifest = JSON.parse(fs.readFileSync(path.join(PACKAGE_ROOT, 'package.json'), 'utf8'));
  assert.match(manifest.scripts['build:ci'], /--allow-untagged/);
  assert.doesNotMatch(manifest.scripts['build:release'], /allow-untagged/);
});

test('release maintainer guide declares supported Info-ZIP toolchain', () => {
  const guide = fs.readFileSync(path.join(PACKAGE_ROOT, 'docs', 'TESTING.md'), 'utf8');
  assert.match(guide, /Info-ZIP zip 3\.x/i);
  assert.match(guide, /Info-ZIP unzip 6\.x/i);
});

test('maintainer release procedure preserves approval, clean-root, live, and review gates', () => {
  // Break caught: maintainers can otherwise publish locally green bytes without
  // proving the clean package, live workbook, privacy scan, or human gate.
  const guide = fs.readFileSync(path.join(PACKAGE_ROOT, 'docs', 'RELEASING.md'), 'utf8');
  assert.match(guide, /explicit approval[\s\S]{0,160}(?:commit|tag)/i);
  assert.match(guide, /fresh temporary[\s\S]{0,100}(?:repository|clone)/i);
  assert.match(guide, /gitleaks/i);
  assert.match(guide, /blank Google Sheet/i);
  assert.match(guide, /_export_info/i);
  assert.match(guide, /marketer-adoption review/i);
  assert.match(guide, /technical and privacy review/i);
  assert.match(guide, /do not[\s\S]{0,120}(?:push|publish|create a remote)/i);
});
