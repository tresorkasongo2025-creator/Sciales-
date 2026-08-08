---
name: Persistent uploaded file storage
description: Uploaded media must survive Autoscale republishes and remain included in backups.
---

Production Autoscale storage is not a durable home for user-uploaded files. The application
stores uploaded media in the persistent PostgreSQL database and keeps the local static copy
only as a compatibility fallback; published deployment is required before this protection
applies to production.

**Why:** Files previously saved only under `static/` disappeared after republishes even though
their filenames remained in PostgreSQL.

**How to apply:** Keep new upload routes, file-serving URLs, migration, and backup/restore
support aligned with the database-backed file store. Never delete the local fallback until
the persistent copy is verified.