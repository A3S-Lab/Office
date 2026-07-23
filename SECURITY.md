# Security Policy

## Supported versions

Security fixes are provided for the latest released minor version. A3S Office
is currently pre-1.0, so applications should keep the package updated and test
release notes before deploying.

## Reporting a vulnerability

Do not open a public issue for a suspected vulnerability.

Use GitHub's private vulnerability reporting for `A3S-Lab/Office`. If that
feature is unavailable, contact an A3S Lab maintainer through GitHub without
including sensitive details and request a private channel. Include:

- the affected version;
- a minimal reproduction;
- the expected security boundary;
- the observed impact; and
- any known workaround.

We will acknowledge a complete report within five business days and coordinate
disclosure after a fix is available.

## Integration responsibilities

A3S Office runs untrusted document content in a browser application. Hosts must:

- use a restrictive Content Security Policy suitable for their deployment;
- validate uploaded file type and size on the server;
- store and serve user files with appropriate authorization;
- keep dependencies and browser runtimes updated;
- treat editor HTML and imported metadata as untrusted outside the editor; and
- avoid exposing AI, upload, or persistence endpoints without authentication.

The library does not provide authentication, collaborative authorization,
malware scanning, or durable storage.
