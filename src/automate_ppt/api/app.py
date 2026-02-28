from __future__ import annotations

from fastapi import FastAPI, HTTPException, Query

from automate_ppt.job_manager import JobManager
from automate_ppt.schemas import (
    EditJobRequest,
    EditJobResponse,
    JobPlanResponse,
    JobStatusResponse,
    SlideInspectionResponse,
)

app = FastAPI(title="ppt-agent-core API", version="0.2.0")
manager = JobManager()


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/presentations/inspect", response_model=SlideInspectionResponse)
def inspect_slide(
    presentation_path: str = Query(..., description="Path to .pptx file"),
    slide: int = Query(..., ge=1),
) -> SlideInspectionResponse:
    try:
        result = manager.inspect_slide(presentation_path, slide)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return SlideInspectionResponse.model_validate(result)


@app.post("/jobs/plan", response_model=JobPlanResponse)
def plan_job(request: EditJobRequest) -> JobPlanResponse:
    try:
        return manager.plan(request)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/jobs", response_model=EditJobResponse)
def create_job(request: EditJobRequest) -> EditJobResponse:
    state = manager.submit(request)
    return EditJobResponse(job_id=state.job_id, status=state.status)


@app.get("/jobs/{job_id}", response_model=JobStatusResponse)
def get_job(job_id: str) -> JobStatusResponse:
    state = manager.get(job_id)
    if state is None:
        raise HTTPException(status_code=404, detail="Job not found")

    return JobStatusResponse(
        job_id=state.job_id,
        status=state.status,
        error=state.error,
        output_path=state.output_path,
    )
