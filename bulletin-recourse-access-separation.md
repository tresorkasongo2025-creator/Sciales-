---
name: Bulletin and recours access separation
description: Rules governing four independent bulletin stages and their payment receipts.
---

The four result stages are independent paid accesses: initial deliberation, first-session recours, second-session deliberation, and second-session recours. Existing sessions are initial by default; importing a new grille must be additive and must never replace, delete, or reset a stage that already contains paid access.

Receipt roles are mutually exclusive: ordinary `B-` receipts unlock only initial sessions, `RS-` receipts submit a first-session recours only, `RR-` receipts unlock only first-session recours results, `S2-` receipts unlock only second-session deliberations, and `R2-` receipts unlock only second-session recours results. Once a valid receipt unlocks its matching bulletin, access remains available even though the receipt is marked used.

**Why:** Students who are absent from a later grille must not be forced to pay for that stage, while students present in the grille need a separate payment to consult the new result. `B2-` already exists as a historical ordinary receipt lot, so `S2-` is used for the new second-session receipt to avoid ambiguity.

**How to apply:** Preserve all four stage values in imports, portal labels, dashboard tracking, payment reports, scan filtering, backups, restores, and future payment or bulletin lookup logic. The payment report supports grouped filters and exact stage filters for 1ère session, 2ème session, recours 1ère session, and recours 2ème session. Treat missing legacy type fields as `initial`/`bulletin`, except for recognizable `RS-`, `RR-`, `S2-`, and `R2-` receipt numbers during migration. A bulletin scan must carry the target bulletin ID so the receipt is checked against the exact displayed stage, not only the matricule.