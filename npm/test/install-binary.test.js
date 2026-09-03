'use strict';

const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const installer = require('../lib/install-binary');

assert.equal(installer.VERSION, '1.0.146-paperai.1');
assert.equal(installer.TAG, 'v1.0.146-paperai.1');
assert.equal(installer.REPO, 'cacdcaecawae/OfficeCLI');
assert.equal(installer.GITHUB_BASE, 'https://github.com/cacdcaecawae/OfficeCLI');

const asset = 'officecli-win-x64.exe';
assert.deepEqual(installer.assetUrls(asset), [
  'https://github.com/cacdcaecawae/OfficeCLI/releases/download/' +
    'v1.0.146-paperai.1/officecli-win-x64.exe'
]);
assert.deepEqual(installer.sumsUrls(), [
  'https://github.com/cacdcaecawae/OfficeCLI/releases/download/' +
    'v1.0.146-paperai.1/SHA256SUMS'
]);
for (const url of installer.assetUrls(asset).concat(installer.sumsUrls())) {
  assert.equal(url.includes('d.officecli.ai'), false);
  assert.equal(url.includes('iOfficeAI/OfficeCLI'), false);
}

const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'officecli-npm-test-'));
try {
  const binary = path.join(temp, asset);
  fs.writeFileSync(binary, 'paperai-test-binary');
  const digest = crypto.createHash('sha256').update(fs.readFileSync(binary)).digest('hex');
  assert.doesNotThrow(() =>
    installer.verifyChecksumText(asset, binary, digest + '  ' + asset + '\n')
  );
  assert.throws(
    () => installer.verifyChecksumText(asset, binary, '0'.repeat(64) + '  ' + asset + '\n'),
    /Checksum mismatch/
  );
  assert.throws(
    () => installer.verifyChecksumText(asset, binary, digest + '  another-asset\n'),
    /must appear exactly once/
  );
  assert.throws(
    () => installer.verifyChecksumText(
      asset,
      binary,
      digest + '  ' + asset + '\n' + digest + '  ' + asset + '\n'
    ),
    /found 2/
  );
} finally {
  fs.rmSync(temp, { recursive: true, force: true });
}

process.stdout.write('install-binary tests passed\n');
