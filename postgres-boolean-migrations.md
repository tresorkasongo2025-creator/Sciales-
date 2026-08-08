---
name: PostgreSQL boolean migration defaults
description: Boolean columns added by startup migrations must use PostgreSQL boolean literals.
---

Use `TRUE` and `FALSE` for PostgreSQL boolean defaults in schema migrations instead of numeric `1` and `0`.

**Why:** PostgreSQL rejects integer default expressions for boolean columns, even though SQLite accepts them.

**How to apply:** When adding boolean columns through the existing startup migration path, use `DEFAULT TRUE` or `DEFAULT FALSE` and keep the model default compatible with both databases.