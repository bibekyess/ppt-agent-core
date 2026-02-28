# ppt-agent-core

A starter project for automating Microsoft PowerPoint using Python and exposing it as an agent-friendly API.

## What is implemented now

This repository includes a practical Phase-1 API with preflight validation:

- Typed operation schema (`replace_text`) for agent edit requests.
- In-memory async job manager with job states: `queued`, `running`, `completed`, `failed`.
- FastAPI endpoints:
  - `GET /health`
  - `GET /presentations/inspect`
  - `POST /jobs/plan`
  - `POST /jobs`
  - `GET /jobs/{job_id}`
- Engine abstraction:
  - Windows: `win32com` PowerPoint engine (real PPT edits + slide inspection)
  - Non-Windows: mock engine (JSON artifacts for development/testing)

## Current supported edit operation

- `replace_text`

## Planned next edit operation (create slide)

Creating slides via API is the next iteration. The target operation shape is:

- `add_slide`
  - create a new slide (append or at index)
  - optional layout (`blank`, `title`, `title_and_content`)
  - optional starter text blocks

Example **planned** payload (not implemented yet):

```json
{
  "schema_version": "1.0",
  "presentation_path": "./samples/ppt-1.pptx",
  "target": { "slide": 2 },
  "operations": [
    {
      "type": "add_slide",
      "position": "after_target",
      "layout": "title_and_content",
      "title": "New Slide from Agent",
      "content": "Generated content goes here"
    }
  ]
}
```

> Note: until `add_slide` is implemented in schemas/planner/engine, API requests must use currently supported operations.


## PowerPoint session behavior

- On Windows, inspect and edit operations now keep PowerPoint and the target presentation open.
- If the file is already open in PowerPoint, the engine attaches to that open presentation and reuses it.
- If the file is not open, the engine opens it once and keeps it open for subsequent operations.

## Why this helps with "shape not found"

Before editing, you can now inspect a slide and run planning validation:

1. Inspect shapes on the target slide (names/ids/text previews).
2. Run `/jobs/plan` using your intended operations.
3. Submit `/jobs` only when plan returns `"valid": true`.

The `shape_selector` now supports:

- `shape_id`
- exact `shape_name`
- `contains_text` (find shape by current text content)

## Run API locally

```bash
uv run uvicorn automate_ppt.api:app --reload
```

## Inspect a slide first

```bash
curl "http://127.0.0.1:8000/presentations/inspect?presentation_path=./samples/ppt-1.pptx&slide=2"
```

## Validate your request before executing

```bash
curl -X POST http://127.0.0.1:8000/jobs/plan \
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

## Execute job (replace text)

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
        "shape_selector": {"contains_text": "Old title"},
        "value": "Updated by agent"
      }
    ]
  }'
```

## Notes

- Real `.pptx` editing/inspection via `win32com` only works on Windows with Microsoft PowerPoint installed.
- In non-Windows environments, inspection/editing uses mock data for workflow validation.
