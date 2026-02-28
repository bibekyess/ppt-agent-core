from __future__ import annotations

import os
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any

from automate_ppt.engine.mock_engine import MockPresentationEngine
from automate_ppt.engine.win32_engine import Win32PresentationEngine
from automate_ppt.planner import plan_job
from automate_ppt.schemas import EditJobRequest, JobPlanResponse


@dataclass
class JobState:
    job_id: str
    status: str
    error: str | None = None
    output_path: str | None = None


class JobManager:
    def __init__(self) -> None:
        self._jobs: dict[str, JobState] = {}
        self._lock = threading.Lock()
        self._executor = ThreadPoolExecutor(max_workers=2)
        self._engine = self._select_engine()
        self._engine_lock = threading.Lock()

    def submit(self, request: EditJobRequest) -> JobState:
        job_id = str(uuid.uuid4())
        with self._lock:
            state = JobState(job_id=job_id, status="queued")
            self._jobs[job_id] = state
        self._executor.submit(self._run_job, job_id, request)
        return state

    def get(self, job_id: str) -> JobState | None:
        with self._lock:
            return self._jobs.get(job_id)

    def inspect_slide(self, presentation_path: str, slide_number: int) -> dict[str, Any]:
        with self._engine_lock:
            self._engine.open(presentation_path)
            return self._engine.inspect_slide(slide_number)

    def plan(self, request: EditJobRequest) -> JobPlanResponse:
        slide_info = self.inspect_slide(request.presentation_path, request.target.slide)
        return plan_job(request, slide_info)

    def _set_status(self, job_id: str, *, status: str, error: str | None = None, output_path: str | None = None) -> None:
        with self._lock:
            state = self._jobs[job_id]
            state.status = status
            state.error = error
            state.output_path = output_path

    def _run_job(self, job_id: str, request: EditJobRequest) -> None:
        self._set_status(job_id, status="running")
        output_path = request.output_path or self._default_output_path(request.presentation_path)

        try:
            with self._engine_lock:
                self._engine.open(request.presentation_path)
                preflight = plan_job(request, self._engine.inspect_slide(request.target.slide))
                if not preflight.valid:
                    raise ValueError(
                        "Preflight validation failed: "
                        + "; ".join(preflight.suggestions + preflight.warnings)
                    )
                self._engine.apply(request.target.slide, request.operations)
                final_path = self._engine.save(output_path)
            self._set_status(job_id, status="completed", output_path=final_path)
        except Exception as exc:
            self._set_status(job_id, status="failed", error=str(exc))

    @staticmethod
    def _default_output_path(input_path: str) -> str:
        base, ext = os.path.splitext(input_path)
        ext = ext or ".json"
        return f"{base}.edited{ext}"

    @staticmethod
    def _select_engine():
        if os.name == "nt":
            return Win32PresentationEngine()
        return MockPresentationEngine()
