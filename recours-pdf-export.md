---
name: Recours PDF export
description: The recours PDF export must tolerate user-entered text and unavailable attachments.
---

The recours PDF exporter escapes all user-entered values before passing them to ReportLab, retries without attachments if an attachment causes generation to fail, and writes large output to a temporary file instead of keeping it all in memory.

**Why:** Production data can contain characters that ReportLab interprets as markup, persistent uploads may be missing or invalid, and large PDFs can exceed worker memory or the default request timeout; none should turn a download into HTTP 500.

**How to apply:** Keep PDF routes defensive, stream large exports from temporary files, allow enough production request time, test both filtered and unfiltered exports, and republish after export fixes.