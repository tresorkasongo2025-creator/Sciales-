---
name: Six receipt categories
description: Business rules for bulletin and recours receipt categories, prefixes, amounts, and stage isolation.
---

The payment system has six independent receipt categories:

- Results session 1: `bulletin`, `B`/historical `B1`/`B2`, 5000 FC.
- Recours submission session 1: `recours`, `RS-`, 10 USD.
- Results of recours session 1: `resultat_recours`, `RR-`, 5000 FC.
- Results session 2: `session_2`, `S2-`, 5000 FC.
- Recours submission session 2: `recours_session_2_soumission`, `RS2-`, 10 USD.
- Results of recours session 2: `recours_session_2`, `R2-`, 5000 FC.

**Why:** A payment receipt must unlock only the exact bulletin stage or recours action for which it was issued. Historical `B2-` receipts are ordinary bulletin receipts and must never be reclassified as `S2-`.

**How to apply:** Validate the printed prefix against the stored receipt type before consuming a receipt. Keep `RS`/`RS2` limited to recours submission and `RR`/`R2`/`S2`/`B` limited to bulletin result consultation. New imports, backups, restores, and PDF generators must preserve the category and canonical amount. The database `type_recu` column must be at least 40 characters; after changing this schema in development, republish so production receives the additive widening.