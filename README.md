# mobile-cockpit-pwa (GitHub Pages mirror)

This repository is a **published mirror** of the `mobile-cockpit` PWA
static bundle. It exists solely so GitHub Pages can serve the PWA as
real HTML/JS/CSS from a controlled HTTPS origin.

## Why a mirror (and not OneDrive)

OneDrive's preview pipeline wraps user-uploaded HTML inside an
`<iframe srcdoc>` with `sandbox=""` (no `allow-scripts` permission), so
MSAL.js never boots and `app.js` never runs. GitHub Pages is a real
HTTPS origin with no sandbox wrapper. See the upstream knowledge base
for the full incident write-up.

## Do NOT edit files here directly

The canonical source-of-truth lives in the Nokia GitLab repo:

- Upstream: <https://gitlabe2.ext.net.nokia.com/ncom_rd_management/management_automation>
- Path:     `flows/mobile-cockpit/pwa/`

Any change made directly in this mirror will be **overwritten** the
next time `scripts/publish-pwa-github.mjs --execute` runs upstream.

## How to update this mirror

From the upstream working tree:

```bash
cd flows/mobile-cockpit
node scripts/publish-pwa-github.mjs --execute --target-dir <path-to-this-clone>
# Then in the target clone:
git add -A && git commit -m 'publish: <reason>' && git push
```

## Build stamp

`2026-06-10 23:58 CEST f270f68`
