---
name: Imported receipt PDFs
description: Durable rules for recognizing already printed and sold receipt PDFs
---

Imported receipt PDFs are treated as a source of already-issued receipts, not as a request to generate replacement receipts. Existing QR payloads must be decoded from the PDF and preserved exactly so the printed paper remains scannable.

**Why:** Generating new QR values would make the physical receipts unusable, while a partial import could validate only part of a sold batch.

**How to apply:** Parse and validate every page, visible receipt number, QR payload, and number/QR association before one atomic database commit. Existing matching number+QR pairs are safe idempotent duplicates; conflicts or unreadable QR codes must abort the entire import. Imported receipts start unused and are consumed only by the existing one-time scan flow.