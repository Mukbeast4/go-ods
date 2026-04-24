# Security Policy

## Supported Versions

Only the latest minor release line receives security fixes. Older versions
may be updated at the maintainers' discretion.

| Version | Supported |
| ------- | --------- |
| 0.2.x   | yes       |
| < 0.2   | no        |

## Reporting a Vulnerability

**Please do not report security vulnerabilities through public GitHub issues,
discussions, or pull requests.**

Instead, use GitHub's private vulnerability reporting feature:

1. Go to <https://github.com/Mukbeast4/go-ods/security/advisories/new>.
2. Fill in the details of the issue (description, affected versions,
   reproduction steps, and impact).
3. Submit the report.

If private reporting is unavailable, you may contact the maintainer listed in
the repository metadata directly.

### What to include

A good report contains:

- A clear description of the vulnerability and its impact.
- The affected version(s) of `go-ods`.
- Steps to reproduce the issue, ideally with a minimal code sample or a
  sample `.ods` file that triggers the behavior.
- Any known workarounds.

### What to expect

- We will acknowledge receipt of your report within **7 days**.
- We will provide an initial assessment and an expected timeline for a fix
  within **14 days**.
- We will keep you informed of progress toward a fix and a coordinated
  release.
- Once a fix is released, we will publish a security advisory crediting the
  reporter (unless anonymity is requested).

## Scope

This policy covers vulnerabilities in the `go-ods` library itself, for
example:

- Parser issues that could cause denial of service when reading untrusted
  `.ods` input (excessive memory allocation, infinite loops, panics).
- Path traversal, zip-slip, or similar issues in the file-handling code.
- Data integrity issues in the writer that produce malformed output
  consumable by other applications.

Issues in third-party dependencies, in the Go standard library, or in
applications that use `go-ods` are out of scope and should be reported to
their respective maintainers.
