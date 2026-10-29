# Preflight A11y — static browser-only build

This folder is the public static version of Preflight A11y. It scans DOCX/PPTX files **entirely in the browser** with JSZip and a rules engine. No backend, no upload, no analytics.

## Develop

```bash
npm install
npm test
npm run build
python3 -m http.server 8877
# open http://127.0.0.1:8877
```

`npm run build` bundles `src/main.js` into `dist/app.js` with esbuild.

## Deploy

The repository root is served by GitHub Pages. `dist/app.js` is committed, so no build step is required on Pages.

## Limits

Rules-based, not a compliance certification. No PDF. No cloud AI in this build. Manual review still required.
