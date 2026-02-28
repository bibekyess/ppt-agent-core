# Agent-Driven PowerPoint Automation Plan

## 1) Product intent

Build a backend service that allows agents to execute presentation edits through explicit commands such as:

- `open_presentation("ppt-1.pptx")`
- `update_slide(slide=2, operations=[...])`
- `preview_slide(slide=2)`
- `save_as("ppt-1.v2.pptx")`

And provide a client UI where humans can **watch work-in-progress** and inspect diffs/history.

---

## 2) Core design principles

1. **Structured operations over free-form scripts**
   - Agents should call typed operations (`add_text`, `replace_text`, `set_theme`) instead of sending arbitrary Python.
2. **Deterministic and replayable jobs**
   - Every request should become an operation log that can be replayed/debugged.
3. **Observable by default**
   - Emit stage-level events and snapshots for each long-running task.
4. **Safe editing boundaries**
   - Validate files, slide indexes, shape IDs, and operation schemas before mutating decks.

---

## 3) Recommended system architecture

### A. API service layer

- Framework: FastAPI (Python) for HTTP + WebSocket support.
- Responsibilities:
  - Authentication/authorization
  - Job creation and status APIs
  - File management (local/S3/etc.)
  - Streaming updates to clients

### B. Job queue / worker layer

- Queue options: Redis + RQ/Celery, or a simple DB-backed queue to start.
- Responsibilities:
  - Execute edit requests serially per presentation (avoid file corruption)
  - Acquire file locks
  - Retry transient failures
  - Emit operation-level progress events

### C. PowerPoint execution engine

- Keep a dedicated module around `win32com` for real Office compatibility on Windows.
- Add an optional abstraction boundary for alternate engines later (e.g., `python-pptx` for non-visual batch operations).
- Suggested interfaces:
  - `PresentationSession.open(path)`
  - `PresentationSession.apply(operations)`
  - `PresentationSession.snapshot(slide)`
  - `PresentationSession.save(path)`

### D. Event streaming and observability

- WebSocket channel keyed by `job_id`.
- Event types:
  - `job.started`
  - `operation.started`
  - `operation.completed`
  - `snapshot.ready`
  - `job.completed`
  - `job.failed`
- Persist events for post-mortem/debug history.

### E. Front-end monitor (optional first iteration)

- A lightweight web app showing:
  - Job timeline
  - Current operation
  - Per-slide snapshots before/after
  - Final downloadable deck

---

## 4) Operation schema (agent contract)

Define a strict JSON schema for edits. Example:

```json
{
  "presentation_id": "ppt-1",
  "target": { "slide": 2 },
  "operations": [
    {
      "type": "replace_text",
      "shape_selector": { "shape_name": "Title 1" },
      "value": "Q3 Business Review"
    },
    {
      "type": "set_fill_color",
      "shape_selector": { "shape_id": 17 },
      "value": "#1F4E79"
    }
  ]
}
```

Key recommendation: version this schema (e.g., `schema_version: "1.0"`) so agent and server can evolve safely.

---

## 5) API endpoints (starter)

- `POST /presentations/upload`
- `POST /jobs` (submit edit job)
- `GET /jobs/{job_id}` (status + summary)
- `GET /jobs/{job_id}/events` (event stream fallback)
- `WS /jobs/{job_id}/stream` (live updates)
- `GET /presentations/{id}/slides/{n}/preview`
- `POST /presentations/{id}/save-as`

---

## 6) Execution flow for your target scenario

User/agent says:

> “Open `ppt-1.pptx`, modify slide 2 with these changes, and let my clients watch live progress.”

Runtime flow:

1. Agent sends structured request to `POST /jobs`.
2. API validates request and creates `job_id`.
3. Worker opens `ppt-1.pptx` (with lock) via `win32com` engine.
4. Worker applies each operation in sequence.
5. After each operation, worker emits progress event + optional slide snapshot.
6. Clients subscribed to WebSocket see timeline and preview updates.
7. Worker saves as new version, emits completion, unlocks file.

---

## 7) Incremental roadmap

### Phase 1: Stabilize local engine

- Refactor current script into reusable service classes.
- Add operation dispatcher + validators.
- Add tests for schema validation and command planning.

### Phase 2: Expose as server

- Build FastAPI endpoints and job state machine.
- Add persistent storage for jobs/events/artifacts.
- Add file-locking and idempotency keys.

### Phase 3: Live client visibility

- Add WebSocket event streaming.
- Generate snapshots (PNG) after key operations.
- Build simple monitoring dashboard.

### Phase 4: Multi-agent and governance

- Add RBAC, tenant isolation, audit trails.
- Add policy checks (e.g., forbid deleting locked slides).
- Add “dry-run” mode to preview planned edits before apply.

---

## 8) Risks and mitigations

- **Windows/Office runtime dependency**: run workers on dedicated Windows hosts.
- **Concurrent edits**: enforce one active mutating job per presentation.
- **Unbounded agent behavior**: allow only whitelisted operations + strict schema validation.
- **Visual mismatch**: include snapshot-based verification after each critical edit.

---

## 9) Suggested immediate next steps in this repo

1. Extract current COM logic into `engine/` module with explicit methods.
2. Introduce a `schemas.py` with pydantic operation models.
3. Add a minimal FastAPI app with `POST /jobs` and in-memory queue.
4. Implement one complete path: open file -> edit slide text -> save -> return output path.
5. Add event logger and one preview endpoint.

This gives you an end-to-end vertical slice quickly, then you can scale reliability and UX.
