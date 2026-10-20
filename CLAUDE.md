# office-extensions

Small Office add-ins for the owner, one per subfolder.

## Stack
- Plain static HTML/CSS/JS, no build step, no dependencies. Office.js loaded from Microsoft's CDN.
- Add-in-only XML manifests (`<addin>/manifest.xml`).
- Outlook event handlers (`hold-mode.js`) must stay plain ES5 with no imports: classic Outlook on Windows
  runs them in a JavaScript-only runtime, and that same file is loaded by `commands.html` and the task pane.

## Deployment
- GitHub Pages from `main` of the public repo `ktgre514-hash/office-extensions`:
  https://ktgre514-hash.github.io/office-extensions/
- Pushing to `main` deploys to what the owner's Outlook loads. Work on a branch and merge via PR.
- The owner installs manifests with **Add a custom add-in → Add from File** (Add from URL isn't offered on their account).
  Code changes need no reinstall; manifest changes do, so bump `<Version>` and send the owner the new file.
- Keep this repo public: Pages on a private repo needs a paid plan, and making it private switched Pages off
  (2026-10-19). The PAT cannot turn Pages back on; that is Settings → Pages → branch `main`, `/ (root)` in the browser.

## Checks
- `node --test outlook-hold-mode/test/parse.test.js`
- `npx office-addin-manifest validate outlook-hold-mode/manifest.xml`
- Preview a task pane in a browser: serve the repo root (`npx http-server -c-1`) and open the pane's HTML;
  outside Outlook, Hold Mode stores settings in localStorage.
