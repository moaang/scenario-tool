# Scenario Tool

Open `index.html` to run the application. PDF.js, its worker, CMaps and standard fonts load from the configured CDN.

Development sources live in `src/`. Run `node build.mjs` after editing them, and `node build.mjs --check` to verify the generated HTML. The build uses Node's standard library; it does not install packages. Deploy the generated `index.html`.

- `import-docx.js`: OOXML source model, styles, pagination and source annotation anchors.
- `import-pdf.js`: PDF.js extraction, structure coverage and selective logical fallback.
- `interpretation.js`: shared screenplay semantics, block identity and geometry ownership.
- `reader.js`: prepared page display, PDF/DOCX overlays and optional Word PDF attachment.
- `app.js`: application state, persistence, domain helpers, controls and shared annotation UI.
- `index.template.html`: markup and styles.

DOCX pages are composed in the browser using available fonts. For Word's exact page layout, attach a PDF exported from the same manuscript using the DOCX reader's “Attach Word PDF” control. The application validates the block mapping before attaching it.
