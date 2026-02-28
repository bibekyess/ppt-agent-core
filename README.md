# ppt-agent-core

A starter project for automating Microsoft PowerPoint using Python and exposing it as an agent-friendly API.

## What is implemented now

This repository now includes a minimal **Phase 1 vertical slice**:

- Typed operation schema (`replace_text`) for agent edit requests.
- In-memory async job manager with job states: `queued`, `running`, `completed`, `failed`.
- FastAPI endpoints:
  - `POST /jobs`
  - `GET /jobs/{job_id}`
  - `GET /health`
- Engine abstraction:
  - Windows: `win32com` PowerPoint engine (real PPT edits)
  - Non-Windows: mock engine that writes a JSON artifact for development/testing

## Run API locally

```bash
uv run uvicorn automate_ppt.api:app --reload
```

## Example request

```bash
curl -X POST http://127.0.0.1:8000/jobs \
  -H 'content-type: application/json' \
  -d '{
    "schema_version": "1.0",
    "presentation_path": "./samples/ppt-1.pptx",
    "target": {"slide": 2},
    "operations": [
      {
        "type": "replace_text",
        "shape_selector": {"shape_name": "Title 1"},
        "value": "Updated by agent"
      }
    ]
  }'
```

## Notes

- Real `.pptx` editing via `win32com` only works on Windows with Microsoft PowerPoint installed.
- The mock engine is for API/workflow validation in non-Windows environments.
