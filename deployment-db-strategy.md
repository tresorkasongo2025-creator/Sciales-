---
name: Deployment DB strategy
description: Why published file storage cannot be used as the permanent database, history of data loss, and the required managed database strategy
---

The app originally used Neon PostgreSQL for production. After the previous credential problem, the application was hardened to require a PostgreSQL `DATABASE_URL` whenever it runs in deployment. The current environment has a working PostgreSQL connection.

**Important platform constraint:** Replit's published-app filesystem is not the permanent data store. SQLite files and uploaded files kept only on the published filesystem must never be treated as permanent storage.

**Data loss incident:** An earlier deployment used SQLite on a published filesystem, so a republish wiped the production database. The user lost production data.

**Current protection:** Deployment is now autoscale-backed with managed PostgreSQL. The application refuses to start in deployment without PostgreSQL instead of silently falling back to SQLite. `_auto_seed_if_empty()` remains only a guarded legacy seed path and does not run when `DATABASE_URL` is present.

**Permanent data rule:** Keep all production records in managed PostgreSQL. The current code stores uploaded assets in the `file_assets` database table and includes them in backups, so republishing does not depend on local filesystem files.

**How to apply:** Never reintroduce SQLite or replace `DATABASE_URL` during publishing. Before a major migration or restore, make a backup and verify database row counts afterward. Additive migrations are acceptable; destructive resets are not.

**Critical rule:** After every `deployConfig(deploymentTarget: "vm", ...)` call, remind the user to REPUBLISH immediately so the change takes effect on the live app. The live deployment keeps its old target until the user clicks Publish.

**PaiementAudit cascade bug:** `paiement_audits.bulletin_id` is `nullable=False`. When deleting a BulletinSession, the delete session route must explicitly delete PaiementAudit rows first (before deleting BulletinData via cascade). This was fixed in `decanat_bulletins_supprimer_session`.
