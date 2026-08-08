---
name: Receipt QR compatibility
description: Legacy receipt QR payloads may contain a scan path instead of only the receipt code.
---

Receipt QR handling must normalize both the raw code and path payloads such as `/scan/CODE` or `/recours/scan/CODE`, including URL-encoded variants, before routing.

**Why:** Older printed receipts encoded a generic `/scan/CODE` path, while the recours scanner expected only a bare code and produced a not-found URL.

**How to apply:** Keep the browser scanner and the server route tolerant of both legacy and current QR payload formats so already-printed receipts remain usable.