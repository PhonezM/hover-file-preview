# Contributing

## Development setup

Use Node.js 22 or later:

```bash
npm ci
bash build.sh
```

Rebuild the committed PowerPoint renderer only when `src/pptx.src.js` changes:

```bash
npm run build:renderer
```

## Checks

```bash
npm test
npm run test:e2e
npm run lint
npm run size
```

The unit suite covers readers and helpers. The six Chromium e2e cases exercise
the packaged extension: PowerPoint, spreadsheet, Word, Markdown, image, and
keyboard previews. Lint checks the Firefox package; the size check enforces the
renderer, unpacked-extension, and zip budgets.

## Code and safety

- Use vanilla JavaScript and do not add runtime dependencies.
- Keep `src/content.js` free of `eval` and `new Function`.
- Escape or safely encode every value derived from a file before inserting it
  into HTML. Prefer DOM APIs and `textContent` where HTML is not required.
- Preserve parser and archive size limits; do not trust declared sizes in an
  input file.

## Adding a preview kind

Add extension detection and the kind's size limit, implement its parser or
renderer, and connect it to the content script's fetch/cache/render flow. Keep
the package build and size budgets in view, and add unit and end-to-end coverage
for the new format and its failure cases.

## Pull requests

- Explain the user-visible change and link any related issue.
- Run the relevant checks above; run all four for changes to parsing, security,
  packaging, or the preview UI.
- Parser and security changes must include unit tests for malformed or hostile
  input. Add an e2e case when behavior in the packaged extension changes.
- Include screenshots for UI changes.
- Update documentation when behavior, limits, or supported formats change.
- Do not add a runtime dependency or exceed a size budget without discussing it
  in the pull request.
