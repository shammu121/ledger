# Ledger

Personal expense + task tracker. Static web app (this repo, served by GitHub Pages) talking to a private Google Apps Script API that stores everything in a Google Sheet.

- `index.html` — the app. Set `API` near the top of the script to your Apps Script `/exec` URL.
- `apps-script/Code.gs` — paste into the Sheet's Apps Script project. Set your own `SECRET` there; never commit it.
- Open the app once with `#key=YOUR_SECRET` (or type it on the lock screen); it stays on that device.
