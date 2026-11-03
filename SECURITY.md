# Security Policy

## Supported versions

Security fixes are provided for the `1.0.x` release line.

## Report a vulnerability

Report vulnerabilities privately through GitHub Security Advisories using
[Report a vulnerability](https://github.com/PhonezM/hover-file-preview/security/advisories/new).
Please do not report an unpatched vulnerability in a public issue. Include the
affected extension version, browser and version, impact, and steps to reproduce.

The maintainer targets an acknowledgement within 7 days. Reports are reviewed
and handled through coordinated disclosure; please allow time for a fix before
publishing details. There is no bug bounty program.

## Scope

Reports relevant to this extension include vulnerabilities in:

- The OOXML parsers for `.docx`, `.xlsx`, `.xlsm`, and `.pptx` files.
- HTML generated for the preview card from file contents.
- ZIP parsing, decompression, and archive resource limits.
- Background-script injection of the PowerPoint renderer.
