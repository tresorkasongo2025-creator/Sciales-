---
name: Recours promotion timers
description: Recours deadlines are configurable per promotion and close all related filières when expired.
---

Recours timers are configured per promotion with allowed durations of 1, 24, 48, or 72 hours. Expiration is enforced server-side and closes every filière in that promotion; the public form only mirrors the current state.

**Why:** A client-side countdown alone could be bypassed, and the requested deadline applies to the whole promotion rather than one filière.

**How to apply:** Keep the DÉCANAT as the only place that starts, restarts, or stops a timer; validate the promotion state both when the form is submitted and when the QR workflow continues.