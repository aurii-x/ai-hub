# AURI Decision Log

Format: Date — Decision — Why — Alternatives considered

---

**2026-06-21 — Renamed santhoshOS to AURI (v0.1)**
Why: the project moved past a personal infra experiment, multi-AI chat, persistence, storage, and security are all working, so it earned a real name. Current state is being treated as "Foundation Complete."
Alternatives considered: keep santhoshOS — rejected, reads as an internal codename rather than something that could carry a landing page or be used outside one person's laptop.

**Setup phase — Used the official LobeHub database deployment instead of the simple LobeChat container**
Why: the simple deployment connected to Postgres but created zero tables, meaning chats were not actually persisted.
Alternatives considered: manually wire Postgres into the simple container — rejected, would mean reverse-engineering Lobe's own schema instead of using the supported path.

**Setup phase — Postgres for the persistence layer**
Why: needed real database-backed storage for chats, agents, and sessions rather than relying on browser storage, which disappears on cache clear or device change.
Alternatives considered: browser-local storage only — rejected, not durable, not portable across devices.

**Setup phase — RustFS for object storage**
Why: S3-compatible, self-hosted object store for files and the future knowledge base, keeps storage under user control instead of a third-party bucket.
Alternatives considered: none formally evaluated yet; revisit if RustFS proves limiting once the knowledge base (Phase 4) is built.

**Setup phase — One active value per secret in `.env`**
Why: duplicate variables in `.env` caused Docker Compose to silently pick the wrong value, which broke RustFS and auth. Made it a hard rule to grep and confirm a single active value per key before starting the stack.

**Setup phase — Stopped the old simple stack before starting the database deployment**
Why: the old Postgres container was still holding port 5432, which blocked the new database-backed stack from starting. Fixed by tearing down the old stack first.

**Setup phase — Rebuilt `bucket.config.json` as a real file**
Why: RustFS init failed because Docker had created `bucket.config.json` as a directory (the source file was missing at mount time). Fixed by deleting the directory and writing the actual JSON policy file.

**Setup phase — Gemini 2.5 Pro / Flash as the first model provider**
Why: older experimental model names had been retired by Google; current models needed to be used. Pro for deeper work, Flash for routine use, multi-provider support (OpenAI, Claude, OpenRouter) deferred to Phase 5.
