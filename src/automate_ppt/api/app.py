from __future__ import annotations

from fastapi import FastAPI, HTTPException

from automate_ppt.job_manager import JobManager
from automate_ppt.schemas import EditJobRequest, EditJobResponse, JobStatusResponse

app = FastAPI(title="ppt-agent-core API", version="0.1.0")
manager = JobManager()


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


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
