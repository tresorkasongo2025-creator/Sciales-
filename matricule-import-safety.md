---
name: Official matricule import safety
description: Rules for matching official student matricules during deliberation-grid imports.
---

Only an exact official promotion list may be selected automatically. Name matching can propose a unique full-name or name/post-name match, but every row remains editable during review. Ambiguous, missing, and reused matricules must be resolved before persistence; the matricule from the Excel grid is never a fallback.

**Why:** A plausible name match or a grid-provided matricule can silently attach results to the wrong student, while duplicate official names must remain visible for DECANAT review.

**How to apply:** Preserve the draft-before-save flow, validate every selected matricule against the official promotion list, reject duplicates, and keep academic session in the re-import identity.