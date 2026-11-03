# Privacy

## File access

The extension fetches a supported file only after its link is hovered or
keyboard-focused long enough to open a preview. It uses
`credentials: "same-origin"`: browser cookies are sent only with same-origin
requests. A cross-origin file can be read only if its server permits the request
with CORS. Files are fetched directly, not through a proxy or third-party
service.

## Storage and collection

Preview data is held in memory only. The LRU cache retains at most 30 entries
and 64 MiB, counting rendered HTML, media, and text as well as downloaded data.
The extension writes nothing to disk. It has no telemetry, analytics, or
remote code.

## Permissions

- `scripting` is used by the background script to inject the PowerPoint renderer
  into the tab's isolated world only when a `.pptx` link is hovered for preview.
- `http://*/*` and `https://*/*` host permissions allow the extension to inspect
  matching links and fetch their files. The browser's same-origin credential
  behavior still applies.

Other preview kinds load about 100 KB of scripts. The PowerPoint renderer is
about 685 KB and is injected only on demand.

For privacy questions, contact the maintainer through
[GitHub](https://github.com/PhonezM/hover-file-preview).
