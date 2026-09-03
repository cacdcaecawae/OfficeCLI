# OfficeCLI — PaperAI fork

This is the PaperAI compatibility fork, version `1.0.146-paperai.1`, based on
upstream OfficeCLI `v1.0.146`. It adds the explicit `strict` and
`office-compatible` DOCX validation profiles. It is not an upstream OfficeCLI
release and is not published to npm by this repository's release workflow.

The package intentionally keeps the internal name `@officecli/officecli` so
existing consumers retain the `officecli` bin. Install it from the immutable
fork Release tarball, not from the npm registry:

```bash
npm install --save-exact \
  https://github.com/cacdcaecawae/OfficeCLI/releases/download/v1.0.146-paperai.1/officecli-officecli-1.0.146-paperai.1.tgz
npx --no-install officecli --version
```

On install, the native binary for your platform (macOS / Linux / Windows,
x64 / arm64) is downloaded only from the matching immutable Release in
`cacdcaecawae/OfficeCLI`. The installer requires the Release's `SHA256SUMS`,
requires the selected asset to appear exactly once, and verifies SHA-256 before
using the binary. It never falls back to `d.officecli.ai` or the upstream
`iOfficeAI/OfficeCLI` repository.

The PaperAI native binary disables automatic global installation and its
upstream self-updater, including previously staged `.update` files. No global
OfficeCLI or per-machine update setting is required. `config autoUpdate`
reports `false`, and enabling it is rejected. Upgrade only by selecting a
reviewed fork Release tarball and updating the lockfile URL and integrity.

## Usage

```bash
officecli create report.docx
officecli add report.docx /body --type paragraph --prop text="Hello"
officecli get report.docx '/body/p[1]'
officecli --help
```

## Notes

- Supported platforms: macOS (arm64/x64), Linux glibc & musl/Alpine
  (arm64/x64), Windows (arm64/x64).
- Set `OFFICECLI_SKIP_BINARY_DOWNLOAD=1` to skip the download during
  `npm install` (the binary is then fetched on first run).
- Fork source and issues: <https://github.com/cacdcaecawae/OfficeCLI>
- Upstream project: <https://github.com/iOfficeAI/OfficeCLI>

Licensed under Apache-2.0.
