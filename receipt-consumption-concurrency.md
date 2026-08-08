---
name: Receipt consumption concurrency
description: Receipt assignment must remain atomic across PostgreSQL and local SQLite.
---

Receipt consumption must combine a database row lock where supported with a conditional update on `utilise = false` before creating the linked payment record.

**Why:** SQLite does not apply `FOR UPDATE`, so a read-then-write check alone can let concurrent operators attribute one receipt twice.

**How to apply:** Any future flow that consumes a receipt—including manual bulletin payments and recovery tools—must reserve it atomically and roll back the reservation if the complete transaction fails.