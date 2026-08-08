---
name: Student relevé portal
description: Student access to relevés uses matricule lookup, protected HTML preview, and a separate REL receipt order.
---

The student relevé portal must never expose the official DOCX as a download. Students authenticate the view with their matricule, see a nominative watermarked image rendered from the official template with the official Dean name visible, and consume a dedicated REL receipt to submit one order per relevé. The DÉCANAT can close the public portal globally.

**Why:** A web page cannot technically prevent an external phone photo or every browser capture, so the system must prevent official-file extraction and make any reproduction traceable instead of promising impossible protection.

**How to apply:** Render the filled DOCX server-side to an image, never return the source document to the browser, keep the University and Faculty header on one line in generated output, keep REL receipts separate from bulletin receipts, consume them atomically when creating a `soumise` order, and let the DÉCANAT update the order status or disable student access.