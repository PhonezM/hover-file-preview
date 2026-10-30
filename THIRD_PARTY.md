# Third-party code

Everything the extension ships is first-party except the libraries bundled
inside `pptx.js` (below). The package minifies that file without its comment
headers, so their notices live here rather than in the shipped bundle.

## Bundled into `pptx.js`

`pptx.js` is generated from `src/pptx.src.js` by `npm run build:renderer`, which
bundles these with it:

| library | version | licence |
| --- | --- | --- |
| uuid | 10.x | MIT |
| tslib | 2.7.x | 0BSD |
| Apache ECharts + ZRender | 5.5.x | Apache License 2.0 |

The bundled renderer carries no `eval`, so the extension runs under the MV3 CSP
(`script-src 'self'`).
