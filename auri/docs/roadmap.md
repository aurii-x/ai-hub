# AURI Roadmap

**Project:** AURI — Personal Intelligence Layer
**Formerly:** santhoshOS
**Current status:** Foundation Complete — v0.1, ~35% to v1.0

---

## Phase tracker

| Phase | Name | Status |
|---|---|---|
| 1 | Infrastructure | Complete |
| 2 | Persistence Verification | Pending |
| 3 | Custom Agents | Pending |
| 4 | Knowledge Base | Pending |
| 5 | Multi-Provider AI | Pending |
| 6 | Branding | Pending |
| 7 | Integrations | Pending |
| 8 | Memory Layer | Pending |

---

## Version timeline

| Version | Milestone |
|---|---|
| v0.1 (today) | Multi-AI chat, database, persistence |
| v0.2 | Knowledge base |
| v0.3 | Custom agents |
| v0.4 | Memory layer |
| v0.5 | Workflow integrations |
| v0.6 | Personal dashboard |
| v1.0 | Digital second brain |

---

## Phase detail

### Phase 1 — Infrastructure (complete)
Git repository, Docker, Docker Compose, GitHub connected, AI layer (Gemini 2.5 Pro / Flash) wired through LobeChat, Postgres persistence verified, RustFS object storage configured, secrets and `.env` handling locked down.

### Phase 2 — Persistence Verification (next)
Goal: prove memory survives a restart.

Test:
1. Create a chat
2. `docker compose down`
3. `docker compose up -d`
4. Confirm the chat still exists

### Phase 3 — Custom Agents
Goal: build the first agents.

- **Resume OS Agent** — resume optimization, job matching, career planning
- **PM Agent** — product strategy, roadmaps, PRDs, market analysis
- **Research Agent** — summarize PDFs, research topics, compare technologies

### Phase 4 — Knowledge Base
RustFS is currently an empty store. Need: PDF upload, notes, documents, personal knowledge. Eventually: Resume OS documents, career notes, PM coursework, research papers.

### Phase 5 — Multi-Provider AI
Currently Gemini only. Add ChatGPT API, Claude API, Copilot (optional). Once done, AURI becomes provider-independent.

### Phase 6 — Branding
auri.ai (optional), landing page, logo, architecture diagram. Hosted on GitHub Pages at `https://<username>.github.io/auri`.

### Phase 7 — Integrations
Notion, Todoist, Gmail, Google Calendar.

### Phase 8 — Memory Layer (important)
Right now Postgres only stores LobeChat's own app data. Eventually AURI Memory should hold preferences, projects, decisions, goals, career data, and learning history — i.e. memory about *Santhosh*, not just about chat sessions.

---

## Near-term plan

1. Verify persistence (~10 min) — create chat, restart Docker, confirm chat exists
2. Build 3 agents (~20 min) — Resume OS, PM, Research
3. Build AURI landing page (~30–60 min) — host on GitHub Pages
