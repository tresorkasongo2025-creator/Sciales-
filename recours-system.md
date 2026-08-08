---
name: Recours system
description: Full appeals (recours) feature — model, routes, templates, and DECANAT management view.
---

## Model
`Recours` table in `app.py` (after `ScanLog`): nom/postnom/prenom/sexe/telephone/promotion/filiere, `reclamations_json` (JSON list of `{id, texte, detail}`), `recu_id` FK → `recus_paiement.id`, `recu_numero`, `date_soumission`, `statut`, `ip_soumission`.

## Claims
12 claim types defined in `_RECOURS_CLAIMS` list of tuples `(id, label, detail_field_name_or_None)`.
Promotions: `_RECOURS_PROMOTIONS`; Filières: `_RECOURS_FILIERES`.

## Routes
- `GET /recours` → `recours_form()` — public form
- `POST /recours/valider` → validates, stores draft in Flask `session['recours_draft']`, redirects to scanner
- `GET /recours/scanner` → `recours_scanner()` mode='scan' — QR scan page (jsQR library)
- `GET /recours/scan/<code>` → validates receipt (unused RecuPaiement), shows confirmation (mode='confirm')
- `POST /recours/scan/<code>` → creates Recours record, marks receipt utilisé, clears session, redirects to /recours/succes
- `GET /recours/succes` → success page (mode='success')
- `GET /decanat/recours` → DECANAT list grouped by promotion, with orphaned-receipt repair section
- `GET /decanat/recours/pdf?promo=X` → PDF via `_generer_recours_pdf()` (ReportLab, one page-break per promotion)
- `POST /decanat/recours/lier-bulletin` → links an orphaned receipt to a bulletin by matricule and marks bulletin as paid

## Templates
- `templates/recours_form.html` — public form
- `templates/recours_scanner.html` — scanner + confirm + error + success (mode variable)
- `templates/decanat_recours.html` — DECANAT view (uses `fromjson` Jinja2 filter, NOT `from_json`)

## Navigation
`base.html`: "Recours" link added after "Bulletins".
`decanat_dashboard.html`: "Gestion des Recours" card added (red border).

**Why:** db.create_all() at startup auto-creates the `recours` table — no manual migration needed.

**How to apply:** Any future claim type additions go in `_RECOURS_CLAIMS` in app.py; the form template iterates them dynamically.
