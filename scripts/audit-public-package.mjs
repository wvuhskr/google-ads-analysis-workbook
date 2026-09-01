#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const VERSION = 'v1.0.0';
const RELEASE_DIRECTORY = path.join(ROOT, 'release-assets', VERSION);
const RELEASE_FILES = [
  'SHA256SUMS.txt',
  `google-ads-analysis-workbook-${VERSION}.js`,
  `google-ads-analysis-workbook-${VERSION}.zip`,
  `sanitize-downloaded-xlsx-${VERSION}.js`,
];
const ARCHIVE_MEMBERS = [
  'DATA-HANDLING.md',
  'LICENSE',
  'QUICKSTART.md',
  `google-ads-analysis-workbook-${VERSION}.js`,
  `sanitize-downloaded-xlsx-${VERSION}.js`,
];
const CHECKSUMMED_ASSETS = RELEASE_FILES.filter((name) => name !== 'SHA256SUMS.txt');
const APPROVED_TRACKED_FILES = [
  '.gitattributes',
  '.github/ISSUE_TEMPLATE/bug-report.yml',
  '.github/ISSUE_TEMPLATE/config.yml',
  '.github/ISSUE_TEMPLATE/feature-request.yml',
  '.github/pull_request_template.md',
  '.github/workflows/ci.yml',
  '.gitignore',
  'CHANGELOG.md',
  'CONTRIBUTING.md',
  'DATA-HANDLING.md',
  'LICENSE',
  'README.md',
  'SECURITY.md',
  'SUPPORT.md',
  'docs/QUICKSTART.md',
  'docs/RELEASING.md',
  'docs/TESTING.md',
  'docs/assets/google-ads-analysis-workbook-synthetic.png',
  'docs/release/DATA-HANDLING.md',
  'docs/release/QUICKSTART.md',
  'google-ads-analysis-workbook.js',
  'package.json',
  'scripts/audit-public-package.mjs',
  'scripts/build-release.mjs',
  'tests/campaign-geo-regression.test.js',
  'tests/hyperlink-regression.test.js',
  'tests/load-exporter.js',
  'tests/master-export.test.js',
  'tests/native-google-sheet-regression.test.js',
  'tests/output-contract-regression.test.js',
  'tests/public-facing-regression.test.js',
  'tests/public-release-contract.test.js',
  'tests/release-readiness-regression.test.js',
  'tests/slim-v1-release-contract.test.js',
  'tests/start-here-regression.test.js',
  'tests/xlsx-hyperlink-sanitizer.test.js',
  'tools/sanitize-downloaded-xlsx.js',
];
const APPROVED_BINARY = 'docs/assets/google-ads-analysis-workbook-synthetic.png';
const APPROVED_BINARY_SHA256 = '828cdbd700d568c87eab874a21dafea17fe7878ee9c18d33daa9ca334a473392';
const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const EXCLUDED_DIRECTORIES = new Set(['.git', 'node_modules', 'release-assets']);
const PROHIBITED_EXTENSIONS = new Set(['.xlsx', '.xls', '.csv', '.tsv', '.pdf', '.log']);
const PROHIBITED_DIRECTORIES = new Set(['work', 'evidence', 'private', 'logs']);
const PROHIBITED_PATTERNS = [
  {
    // Private work-machine login. The public GitHub owner "wvuhskr" is a
    // separate, intentionally public handle and is not denylisted here.
    name: 'known private username',
    pattern: new RegExp(['alex', 'murthawork'].join(''), 'i'),
  },
  {
    name: 'retired product or release name',
    pattern: new RegExp([
      ['Master Account', ' Export'].join(''),
      ['master-account', '-export'].join(''),
      ['google-ads-master', '-export'].join(''),
    ].join('|'), 'i'),
  },
  { name: 'workbook URL', pattern: /https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]+\/edit\b/i },
  { name: 'local absolute path', pattern: /\/(?:Users|home)\/[A-Za-z0-9_.-]+/ },
  {
    name: 'GitHub access token',
    pattern: new RegExp('\\bgh[pousr]_[A-Za-z0-9]{20,255}\\b'),
  },
  {
    name: 'GitHub fine-grained access token',
    pattern: new RegExp('\\b' + ['github', '_pat_'].join('') + '[A-Za-z0-9_]{20,255}\\b'),
  },
  {
    name: 'Google API key',
    pattern: new RegExp('\\b' + ['AI', 'za'].join('') + '[A-Za-z0-9_-]{30,255}\\b'),
  },
  {
    name: 'Google OAuth access token',
    pattern: new RegExp('\\b' + ['ya', '29\\.'].join('') + '[A-Za-z0-9._-]{20,255}\\b'),
  },
  { name: 'AWS access key', pattern: /\b(?:AKIA|ASIA)[A-Z0-9]{16}\b/ },
  {
    name: 'Stripe live secret',
    pattern: new RegExp('\\b' + ['sk', '_live_'].join('') + '[A-Za-z0-9]{16,255}\\b'),
  },
  {
    name: 'Slack access token',
    pattern: new RegExp('\\b' + ['xox', '[abprs]-'].join('') + '[A-Za-z0-9-]{20,255}\\b'),
  },
  {
    name: 'GitLab access token',
    pattern: new RegExp('\\b' + ['glpat', '-'].join('') + '[A-Za-z0-9_-]{20,255}\\b'),
  },
  {
    name: 'private key marker',
    pattern: new RegExp(['-----BEGIN ', '(?:RSA |EC |OPENSSH )?', 'PRIVATE KEY-----'].join('')),
  },
  {
    name: 'release-candidate language',
    pattern: new RegExp(['R' + 'C15', 'release' + ' candidate'].join('|'), 'i'),
  },
];
const CUSTOMER_ID_PATTERN = /\b(?:customer|account)[_. -]?id\b\D{0,16}((?:\d{3}[- ]?){2}\d{4})\b/gi;
const ADOPTION_DOCUMENTS = [
  'README.md',
  path.join('docs', 'QUICKSTART.md'),
  path.join('docs', 'release', 'QUICKSTART.md'),
];
const PUBLIC_CLAIM_DOCUMENTS = [
  'README.md',
  'CHANGELOG.md',
  'CONTRIBUTING.md',
  'DATA-HANDLING.md',
  'SUPPORT.md',
  path.join('docs', 'QUICKSTART.md'),
  path.join('docs', 'TESTING.md'),
  path.join('docs', 'release', 'DATA-HANDLING.md'),
  path.join('docs', 'release', 'QUICKSTART.md'),
  'google-ads-analysis-workbook.js',
  'package.json',
];
const UNSUPPORTED_PRODUCT_CLAIMS = [
  {
    name: 'API independence claim',
    pattern: /\b(?:API[- ]free|(?:works?|runs?|operates?)\s+(?:entirely\s+)?without\s+(?:any\s+)?(?:Google Ads\s+)?API|no\s+(?:Google Ads\s+)?API\s+(?:is\s+)?required)\b/i,
  },
  {
    name: 'restorable backup claim',
    pattern: /\b(?:complete|full|restorable)(?:\s+(?:complete|full|restorable))*\s+(?:Google Ads\s+)?(?:account\s+)?backup\b/i,
  },
  {
    name: 'exhaustive data claim',
    pattern: /\bexports?\s+(?:all|every)\s+(?:Google Ads\s+)?(?:data|records?|resources?|fields?)\b/i,
  },
  { name: 'one-step execution claim', pattern: /\b(?:one|1)[- ](?:click|run)\b/i },
  {
    name: 'manager-account support claim',
    pattern: /\b(?:works?|runs?|exports?)\s+(?:from|across|for)\s+(?:Google Ads\s+)?(?:manager accounts?|MCCs?)\b/i,
  },
  {
    name: 'real-time output claim',
    pattern: /\b(?:is|provides?|creates?|delivers?|offers?)\s+(?:an?\s+)?real[- ]time\b/i,
  },
  {
    name: 'automatic optimization claim',
    pattern: /\b(?:(?:automatically|autonomously)\s+(?:optimizes?|applies?)|(?:is|provides?|includes?)\s+(?:an?\s+)?automated\s+optimizer)\b/i,
  },
  {
    name: 'anonymization claim',
    pattern: /(?<!not\s)\b(?:anonymizes?|redacts?|de-identifies?)\s+(?:all|every|the)\s+(?:outputs?|workbooks?|data)\b/i,
  },
  {
    name: 'automatic XLSX claim',
    pattern: /(?:^|[\n.!?]\s*)(?:(?:the\s+)?(?:exporter|tool|script|project|it)\s+)?automatically\s+(?:creates?|exports?|distributes?|uploads?)\s+(?:an?\s+)?XLSX\b/im,
  },
  {
    name: 'sharing authorization claim',
    pattern: /\bREADY(?:\s+workbook)?\s+(?:automatically\s+)?(?:authorizes?|grants?\s+permission\s+for?)\s+(?:unrestricted\s+)?(?:sharing|uploading)\b/i,
  },
];
const REQUIRED_PRODUCT_DISCLOSURES = [
  {
    name: 'single-advertiser scope',
    pattern: /individual\s+advertiser\s+account/i,
  },
  {
    name: 'Google Ads read-only boundary',
    pattern: /read-only\s+with\s+respect\s+to\s+Google\s+Ads/i,
  },
  {
    name: 'configured Google Sheet write boundary',
    pattern: /writes?[\s\S]{0,120}configured\s+Google\s+Sheet/i,
  },
  {
    name: 'authorized human or LLM-assisted analysis purpose',
    pattern: /authorized\s+human\s+or\s+LLM-assisted\s+analysis/i,
  },
  {
    name: 'separate Google Ads API setup distinction',
    pattern: /without\s+separately\s+provisioning\s+a\s+Google\s+Ads\s+API\s+developer\s+token/i,
  },
  {
    name: 'Google authorization requirement',
    pattern: /Google\s+authorization\s+is\s+still\s+required/i,
  },
  {
    name: 'editable backup exclusion',
    pattern: /not\s+an\s+editable\s+Google\s+Ads\s+account\s+backup/i,
  },
  {
    name: 'non-atomic snapshot boundary',
    pattern: /not[\s\S]{0,160}\batomic(?:\s+account)?\s+snapshot/i,
  },
  {
    name: 'optimizer exclusion',
    pattern: /not[\s\S]{0,180}\bautomated\s+optimizer/i,
  },
  {
    name: 'non-exhaustive coverage boundary',
    pattern: /not[\s\S]{0,240}\bexhaustive\s+export/i,
  },
  {
    name: 'manual XLSX boundary',
    pattern: /does\s+not[\s\S]{0,120}\bautomatically[\s\S]{0,50}\bcreate[\s\S]{0,50}\bXLSX/i,
  },
  {
    name: 'anonymization exclusion',
    pattern: /(?:not[\s\S]{0,280}\banonymization\s+tool|does\s+not\s+anonymize)/i,
  },
  {
    name: 'READY sharing-permission boundary',
    pattern: /\bREADY\b[\s\S]{0,180}does not grant permission to share or upload/i,
  },
];

function fail(message) {
  throw new Error(message);
}

function sha256File(filePath) {
  return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

function sha256Bytes(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function hasOnlyObviousSyntheticCustomerId(value) {
  const digits = value.replace(/[- ]/g, '');
  return /^([0-9])\1{9}$/.test(digits) || digits === '1234567890';
}

function assertTextSafe(text, description) {
  for (const prohibited of PROHIBITED_PATTERNS) {
    if (prohibited.pattern.test(text)) fail(`Found ${prohibited.name} in ${description}`);
  }
  for (const match of text.matchAll(CUSTOMER_ID_PATTERN)) {
    if (!hasOnlyObviousSyntheticCustomerId(match[1])) {
      fail(`Found customer ID in ${description}`);
    }
  }
}

function decodeStrictText(content, description) {
  if (content.includes(0)) fail(`Unexpected binary content in ${description}`);
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(content);
  } catch (_error) {
    fail(`Unexpected binary or non-UTF-8 content in ${description}`);
  }
}

function assertApprovedBinary(content, relativePath, description) {
  if (relativePath !== APPROVED_BINARY) {
    fail(`Unexpected binary file in public candidate: ${relativePath}`);
  }
  if (content.length < PNG_SIGNATURE.length || !content.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)) {
    fail(`Approved preview has an invalid PNG signature: ${description}`);
  }
  if (sha256Bytes(content) !== APPROVED_BINARY_SHA256) {
    fail(`Approved preview PNG hash mismatch: ${description}`);
  }
}

function assertSafeBlob(content, relativePath, description) {
  if (relativePath === APPROVED_BINARY) {
    assertApprovedBinary(content, relativePath, description);
    return;
  }
  assertTextSafe(decodeStrictText(content, description), description);
}

function assertBytesMatch(description, actual, expected) {
  if (!actual.equals(expected)) fail(`Release byte mismatch for ${description}`);
}

function readArchiveMember(archivePath, member) {
  return execFileSync('unzip', ['-p', archivePath, member]);
}

function walk(directory, root = directory) {
  const entries = [];
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      if (!EXCLUDED_DIRECTORIES.has(entry.name)) entries.push(...walk(path.join(directory, entry.name), root));
      continue;
    }
    if (entry.isFile()) entries.push(path.relative(root, path.join(directory, entry.name)));
  }
  return entries.sort();
}

function git(candidateRoot, args, encoding = 'buffer') {
  return execFileSync('git', args, {
    cwd: candidateRoot,
    encoding,
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

function isGitRepositoryRoot(candidateRoot) {
  try {
    const detected = String(git(candidateRoot, ['rev-parse', '--show-toplevel'], 'utf8')).trim();
    return fs.realpathSync(detected) === fs.realpathSync(candidateRoot);
  } catch (_error) {
    return false;
  }
}

function assertPathAllowed(relativePath) {
    const components = relativePath.split(path.sep);
    if (components.some((component) => PROHIBITED_DIRECTORIES.has(component))) {
      fail(`Prohibited private directory in public candidate: ${relativePath}`);
    }
    if (PROHIBITED_EXTENSIONS.has(path.extname(relativePath).toLowerCase())) {
      fail(`Prohibited live-data file type in public candidate: ${relativePath}`);
    }
}

function parseTree(candidateRoot) {
  const raw = git(candidateRoot, ['ls-tree', '-rz', 'HEAD']);
  return raw.toString('utf8').split('\0').filter(Boolean).map((record) => {
    const separator = record.indexOf('\t');
    if (separator < 0) fail('Could not parse Git tree entry.');
    const [mode, type, object] = record.slice(0, separator).split(' ');
    return { mode, type, object, relativePath: record.slice(separator + 1) };
  });
}

function assertGitCandidateSafe(candidateRoot) {
  const status = String(git(
    candidateRoot,
    ['status', '--porcelain=v1', '--untracked-files=all'],
    'utf8',
  ));
  if (status.trim()) fail('Public candidate Git tree must be clean with no tracked or untracked changes.');

  const entries = parseTree(candidateRoot);
  const actualPaths = entries.map((entry) => entry.relativePath);
  if (actualPaths.length !== APPROVED_TRACKED_FILES.length ||
      actualPaths.some((relativePath, index) => relativePath !== APPROVED_TRACKED_FILES[index])) {
    const unexpected = actualPaths.filter((candidate) => !APPROVED_TRACKED_FILES.includes(candidate));
    const missing = APPROVED_TRACKED_FILES.filter((candidate) => !actualPaths.includes(candidate));
    fail(`Public candidate manifest mismatch. Unexpected: ${unexpected.join(', ') || 'none'}; missing: ${missing.join(', ') || 'none'}`);
  }
  for (const entry of entries) {
    if (entry.mode !== '100644' || entry.type !== 'blob') {
      fail(`Public candidate contains a symlink, submodule, or unsupported mode: ${entry.relativePath}`);
    }
    assertPathAllowed(entry.relativePath);
    const content = git(candidateRoot, ['cat-file', 'blob', entry.object]);
    assertSafeBlob(content, entry.relativePath, `HEAD blob: ${entry.relativePath}`);
  }
}

function assertDirectoryCandidateSafe(candidateRoot) {
  for (const relativePath of walk(candidateRoot)) {
    assertPathAllowed(relativePath);
    const content = fs.readFileSync(path.join(candidateRoot, relativePath));
    assertSafeBlob(content, relativePath, `public candidate: ${relativePath}`);
  }
}

function assertCandidateSafe(candidateRoot) {
  if (isGitRepositoryRoot(candidateRoot)) assertGitCandidateSafe(candidateRoot);
  else assertDirectoryCandidateSafe(candidateRoot);
}

function sourceBytes(candidateRoot, relativePath) {
  if (isGitRepositoryRoot(candidateRoot)) {
    return git(candidateRoot, ['show', `HEAD:${relativePath}`]);
  }
  return fs.readFileSync(path.join(candidateRoot, relativePath));
}

function assertSupportedProductClaims(candidateRoot) {
  for (const relativePath of PUBLIC_CLAIM_DOCUMENTS) {
    const content = decodeStrictText(
      sourceBytes(candidateRoot, relativePath),
      `public claim surface: ${relativePath}`,
    );
    for (const unsupported of UNSUPPORTED_PRODUCT_CLAIMS) {
      if (unsupported.pattern.test(content)) {
        fail(`Unsupported product scope claim in ${relativePath}: ${unsupported.name}`);
      }
    }
  }
  for (const relativePath of ADOPTION_DOCUMENTS) {
    const content = decodeStrictText(
      sourceBytes(candidateRoot, relativePath),
      `adoption document: ${relativePath}`,
    );
    for (const required of REQUIRED_PRODUCT_DISCLOSURES) {
      if (!required.pattern.test(content)) {
        fail(`Required product scope boundary is missing from ${relativePath}: ${required.name}`);
      }
    }
  }
}

function assertReleaseAssets(releaseDirectory, candidateRoot) {
  if (!fs.existsSync(releaseDirectory)) fail(`Release directory is missing: ${releaseDirectory}`);
  const actual = fs.readdirSync(releaseDirectory).sort();
  if (actual.length !== RELEASE_FILES.length || actual.some((name, index) => name !== RELEASE_FILES[index])) {
    fail(`Release directory contains unexpected files: ${actual.join(', ')}`);
  }
  const archivePath = path.join(releaseDirectory, `google-ads-analysis-workbook-${VERSION}.zip`);
  const members = execFileSync('unzip', ['-Z', '-1', archivePath], { encoding: 'utf8' })
    .trim().split('\n').filter(Boolean).sort();
  if (members.length !== ARCHIVE_MEMBERS.length || members.some((name, index) => name !== ARCHIVE_MEMBERS[index])) {
    fail(`Release ZIP contains unexpected members: ${members.join(', ')}`);
  }
  for (const member of members) {
    if (/[\\/]/.test(member) || !(/\.(?:js|md)$/.test(member) || member === 'LICENSE')) {
      fail(`Release ZIP contains unsafe member name: ${member}`);
    }
  }
  const releaseInputs = [
    {
      releaseName: `google-ads-analysis-workbook-${VERSION}.js`,
      sourcePath: 'google-ads-analysis-workbook.js',
    },
    {
      releaseName: `sanitize-downloaded-xlsx-${VERSION}.js`,
      sourcePath: path.join('tools', 'sanitize-downloaded-xlsx.js'),
    },
  ];
  for (const input of releaseInputs) {
    const release = fs.readFileSync(path.join(releaseDirectory, input.releaseName));
    const zipped = readArchiveMember(archivePath, input.releaseName);
    assertTextSafe(release.toString('utf8'), `release asset: ${input.releaseName}`);
    assertTextSafe(zipped.toString('utf8'), `ZIP member: ${input.releaseName}`);
    const source = sourceBytes(candidateRoot, input.sourcePath);
    assertBytesMatch(input.releaseName, release, source);
    assertBytesMatch(`ZIP member ${input.releaseName}`, zipped, source);
  }
  const archiveDocuments = [
    { member: 'QUICKSTART.md', sourcePath: path.join('docs', 'release', 'QUICKSTART.md') },
    { member: 'DATA-HANDLING.md', sourcePath: path.join('docs', 'release', 'DATA-HANDLING.md') },
    { member: 'LICENSE', sourcePath: 'LICENSE' },
  ];
  for (const document of archiveDocuments) {
    const zipped = readArchiveMember(archivePath, document.member);
    assertTextSafe(zipped.toString('utf8'), `ZIP member: ${document.member}`);
    assertBytesMatch(`ZIP member ${document.member}`, zipped, sourceBytes(candidateRoot, document.sourcePath));
  }
  const sums = fs.readFileSync(path.join(releaseDirectory, 'SHA256SUMS.txt'), 'utf8').trim().split('\n');
  if (sums.length !== CHECKSUMMED_ASSETS.length) fail('SHA256SUMS.txt has an unexpected number of entries.');
  const seen = new Set();
  for (const line of sums) {
    const match = /^([a-f0-9]{64})  ([^\\/]+)$/.exec(line);
    if (!match || !CHECKSUMMED_ASSETS.includes(match[2]) || seen.has(match[2])) {
      fail(`Invalid checksum entry: ${line}`);
    }
    seen.add(match[2]);
    if (match[1] !== sha256File(path.join(releaseDirectory, match[2]))) {
      fail(`Checksum mismatch for ${match[2]}`);
    }
  }
}

function audit(candidateRoot = ROOT, releaseDirectory) {
  const root = path.resolve(candidateRoot);
  assertCandidateSafe(root);
  assertSupportedProductClaims(root);
  const release = releaseDirectory ? path.resolve(releaseDirectory) : (root === ROOT ? RELEASE_DIRECTORY : '');
  if (release) assertReleaseAssets(release, root);
}

function main() {
  if (process.argv.length > 4) {
    fail('Usage: node scripts/audit-public-package.mjs [candidate-directory] [release-directory]');
  }
  audit(process.argv[2] || ROOT, process.argv[3]);
}

try {
  main();
} catch (error) {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
}

export { audit };
