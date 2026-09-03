'use strict';

// Installer for the PaperAI OfficeCLI fork. The package ships no native code:
// it fetches the platform binary only from the matching immutable release in
// cacdcaecawae/OfficeCLI. Do not add the upstream mirror as a fallback: it can
// return a same-base-version binary without the PaperAI compatibility patch.

const fs = require('fs');
const path = require('path');
const https = require('https');
const crypto = require('crypto');
const { execSync } = require('child_process');

const REPO = 'cacdcaecawae/OfficeCLI';
const GITHUB_BASE = 'https://github.com/' + REPO;

// The package version is set to the release version at publish time, so the
// release tag we download from is derived directly from it (immutable, never
// stale). Keep the prerelease suffix because v1.0.146 and
// v1.0.146-paperai.1 are different binaries. Build metadata is not used in the
// release tag.
const VERSION = require('../package.json').version;
const TAG = 'v' + VERSION.split('+')[0];

const PKG_ROOT = path.join(__dirname, '..');
// Native binary lives under vendor/, NOT bin/: the repo's root .gitignore
// ignores `bin/`, and keeping the download target out of bin/ avoids any
// collision with the launcher shim.
const BIN_DIR = path.join(PKG_ROOT, 'vendor');

function log(msg) {
  // postinstall output goes to stderr so it never pollutes a command's stdout.
  process.stderr.write('[officecli] ' + msg + '\n');
}

// musl detection, mirroring install.sh's gnu-vs-musl branch. process.report
// exposes glibcVersionRuntime on a glibc system; its absence (plus the Alpine
// marker / `ldd` text) means musl.
function isMusl() {
  if (process.platform !== 'linux') return false;
  try {
    const report = process.report && process.report.getReport();
    const header = report && report.header;
    if (header && header.glibcVersionRuntime) return false;
    if (header && header.glibcVersionRuntime === undefined) {
      // No glibc runtime reported — treat as musl, but confirm below.
    }
  } catch (_) { /* fall through to filesystem/ldd probes */ }
  try {
    if (fs.existsSync('/etc/alpine-release')) return true;
  } catch (_) { /* ignore */ }
  try {
    const out = execSync('ldd --version 2>&1 || true', { encoding: 'utf8' });
    if (/musl/i.test(out)) return true;
  } catch (_) { /* ignore */ }
  // Default to glibc when nothing positively indicates musl.
  return false;
}

function detectAsset() {
  const platform = process.platform;
  const arch = process.arch;
  if (platform === 'darwin') {
    if (arch === 'arm64') return 'officecli-mac-arm64';
    if (arch === 'x64') return 'officecli-mac-x64';
  } else if (platform === 'linux') {
    const musl = isMusl();
    if (arch === 'x64') return musl ? 'officecli-linux-alpine-x64' : 'officecli-linux-x64';
    if (arch === 'arm64') return musl ? 'officecli-linux-alpine-arm64' : 'officecli-linux-arm64';
  } else if (platform === 'win32') {
    if (arch === 'x64') return 'officecli-win-x64.exe';
    if (arch === 'arm64') return 'officecli-win-arm64.exe';
  }
  throw new Error(
    'Unsupported platform: ' + platform + ' ' + arch +
    '. Download manually from ' + GITHUB_BASE + '/releases'
  );
}

function binaryName() {
  return process.platform === 'win32' ? 'officecli.exe' : 'officecli';
}

function binaryPath() {
  return path.join(BIN_DIR, binaryName());
}

function assetUrls(asset) {
  return [
    GITHUB_BASE + '/releases/download/' + TAG + '/' + asset
  ];
}

function sumsUrls() {
  return [
    GITHUB_BASE + '/releases/download/' + TAG + '/SHA256SUMS'
  ];
}

function httpGet(url, onResponse, onError, redirects) {
  redirects = redirects || 0;
  if (redirects > 10) {
    onError(new Error('Too many redirects for ' + url));
    return;
  }
  const req = https.get(
    url,
    { headers: { 'User-Agent': 'officecli-npm-installer' } },
    function (res) {
      const code = res.statusCode;
      if (code >= 300 && code < 400 && res.headers.location) {
        res.resume();
        const next = new URL(res.headers.location, url).toString();
        httpGet(next, onResponse, onError, redirects + 1);
        return;
      }
      if (code !== 200) {
        res.resume();
        onError(new Error('HTTP ' + code + ' for ' + url));
        return;
      }
      onResponse(res);
    }
  );
  req.on('error', onError);
  req.setTimeout(300000, function () {
    req.destroy(new Error('Timeout downloading ' + url));
  });
}

function fetchToFile(url, dest) {
  return new Promise(function (resolve, reject) {
    httpGet(
      url,
      function (res) {
        const tmp = dest + '.download';
        const out = fs.createWriteStream(tmp);
        res.pipe(out);
        out.on('error', reject);
        out.on('finish', function () {
          out.close(function () {
            try {
              fs.renameSync(tmp, dest);
              resolve();
            } catch (e) {
              reject(e);
            }
          });
        });
      },
      reject
    );
  });
}

function fetchBuffer(url) {
  return new Promise(function (resolve, reject) {
    httpGet(
      url,
      function (res) {
        const chunks = [];
        res.on('data', function (c) { chunks.push(c); });
        res.on('end', function () { resolve(Buffer.concat(chunks)); });
        res.on('error', reject);
      },
      reject
    );
  });
}

function verifyChecksumText(asset, file, sums) {
  // SHA256SUMS rows are "<hex>  <name>" (sha256sum text mode). Match the
  // filename column exactly (a leading '*' marks binary mode), never a
  // substring. Reject malformed or duplicate rows instead of guessing.
  const matches = [];
  for (const line of sums.split('\n')) {
    const match = /^([a-f0-9]{64})\s+\*?(.+?)\s*$/i.exec(line);
    if (match && match[2] === asset) matches.push(match[1]);
  }
  if (matches.length !== 1) {
    throw new Error(
      asset + ' must appear exactly once in SHA256SUMS (found ' + matches.length + ')'
    );
  }
  const expected = matches[0];
  const actual = crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
  if (actual.toLowerCase() !== expected.toLowerCase()) {
    throw new Error('Checksum mismatch for ' + asset + ' (expected ' + expected + ', got ' + actual + ')');
  }
}

async function verifyChecksum(asset, file) {
  const url = sumsUrls()[0];
  let sums;
  try {
    sums = (await fetchBuffer(url)).toString('utf8');
  } catch (error) {
    throw new Error('Could not download mandatory SHA256SUMS from ' + url + ': ' + error.message);
  }
  verifyChecksumText(asset, file, sums);
  log('  checksum verified.');
}

// Download the platform binary into bin/ if it is not already present.
// Idempotent: a non-empty binary is treated as already installed (the package
// version pins the release, so existence is sufficient).
async function ensureBinary() {
  const dest = binaryPath();
  if (fs.existsSync(dest) && fs.statSync(dest).size > 0) {
    return dest;
  }
  fs.mkdirSync(BIN_DIR, { recursive: true });
  const asset = detectAsset();
  let lastErr = null;
  for (const url of assetUrls(asset)) {
    try {
      log('Downloading ' + asset + ' (' + TAG + ') from ' + url + ' ...');
      await fetchToFile(url, dest);
      await verifyChecksum(asset, dest);
      if (process.platform !== 'win32') {
        fs.chmodSync(dest, 0o755);
      }
      log('OfficeCLI ' + VERSION + ' installed.');
      return dest;
    } catch (e) {
      lastErr = e;
      try { fs.rmSync(dest, { force: true }); } catch (_) { /* ignore */ }
      try { fs.rmSync(dest + '.download', { force: true }); } catch (_) { /* ignore */ }
      log('  failed: ' + e.message);
    }
  }
  throw new Error(
    'Could not download OfficeCLI binary (' + asset + ' @ ' + TAG + '). ' +
    'Last error: ' + (lastErr && lastErr.message) +
    '. Install manually from ' + GITHUB_BASE + '/releases'
  );
}

module.exports = {
  ensureBinary: ensureBinary,
  binaryPath: binaryPath,
  detectAsset: detectAsset,
  assetUrls: assetUrls,
  sumsUrls: sumsUrls,
  verifyChecksumText: verifyChecksumText,
  REPO: REPO,
  GITHUB_BASE: GITHUB_BASE,
  VERSION: VERSION,
  TAG: TAG
};
