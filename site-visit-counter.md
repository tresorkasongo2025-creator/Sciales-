---
name: Site visit counter
description: Public homepage visits are persisted and summarized for the DÉCANAT.
---

The site visit counter records one visit per request to the public homepage and stores only the timestamp, with no IP address or personal data.

**Why:** The DÉCANAT needs a persistent traffic total that survives restarts and publication without turning the counter into a personal tracking system.

**How to apply:** Keep homepage counting separate from static assets and expose the total plus today's count in the DÉCANAT dashboard.