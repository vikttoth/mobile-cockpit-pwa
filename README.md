# mobile-cockpit-pwa (GitHub Pages mirror)

This repository is a **published mirror** of the `mobile-cockpit` PWA
static bundle. It exists solely so GitHub Pages can serve the PWA as
real HTML/JS/CSS from a controlled HTTPS origin. Two independent mirrors
exist since 2026-09-30, when the corporate proxy started blocking
`git push` to github.com (reads, and api.github.com writes, still pass):
GitLab Pages (pushed with `git`, reachable only on the Nokia network/VPN)
and GitHub Pages (published via the GitHub REST "Git Data" API instead
of `git push` -- see `scripts/publish-pwa-github-api.mjs` -- reachable
from anywhere, no VPN). `pwa/config.json`'s `shareLinkBase` says which one
is the one guests are actually sent.

## Why a mirror (and not OneDrive)

OneDrive's preview pipeline wraps user-uploaded HTML inside an
`<iframe srcdoc>` with `sandbox=""` (no `allow-scripts` permission), so
MSAL.js never boots and `app.js` never runs. Pages is a real
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
bash scripts/deploy-pwa-github.sh --execute                  # GitLab Pages (git push)
node scripts/publish-pwa-github-api.mjs --execute             # GitHub Pages (REST API, no git push)
```

## Build stamp

`2026-10-08 10:40 CEST f211de0`
