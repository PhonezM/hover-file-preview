# Releasing

1. Update the version in `manifest.json` and `package.json`; update the
   changelog for the release.
2. If `src/pptx.src.js` changed, rebuild the committed renderer:
   `npm run build:renderer`.
3. Run `bash build.sh` and `npm run size`. Run the test commands before release.
4. Commit the release changes and create the `vX.Y.Z` tag. Push the tag to
   trigger the release workflow.
5. CI attaches the Firefox and Chrome packages to the GitHub release. When
   `WEB_EXT_API_KEY` and `WEB_EXT_API_SECRET` are configured, CI also signs the
   Firefox package for AMO and attaches the signed XPI. Upload the Chrome package
   to the Chrome Web Store manually.

Do not move or replace a published tag. If a release needs correction, publish
a new version and tag. The `1.0.1` re-cut is a one-off exception; do not repeat
that process for later releases.
