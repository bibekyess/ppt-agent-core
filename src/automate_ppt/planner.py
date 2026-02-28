from __future__ import annotations

import difflib
from typing import Any

from automate_ppt.schemas import EditJobRequest, JobPlanResponse, ReplaceTextOperation


def plan_job(request: EditJobRequest, slide_info: dict[str, Any]) -> JobPlanResponse:
    shapes = slide_info.get("shapes", [])
    shape_names = [str(s.get("shape_name")) for s in shapes if s.get("shape_name")]

    warnings: list[str] = []
    suggestions: list[str] = []
    valid = True

    for op in request.operations:
        if isinstance(op, ReplaceTextOperation):
            if _operation_matches_shape(op, shapes):
                continue

            valid = False
            selector = op.shape_selector
            if selector.shape_name:
                close = difflib.get_close_matches(selector.shape_name, shape_names, n=3, cutoff=0.3)
                if close:
                    suggestions.append(
                        f"replace_text: selector '{selector.shape_name}' not found; close matches: {', '.join(close)}"
                    )
                else:
                    warnings.append(
                        f"replace_text: selector '{selector.shape_name}' not found on slide {request.target.slide}."
                    )
            else:
                warnings.append(
                    "replace_text: selector did not match any shape. Consider shape_id, exact shape_name, or contains_text."
                )

    return JobPlanResponse(
        valid=valid,
        target_slide=request.target.slide,
        available_shapes=shape_names,
        warnings=warnings,
        suggestions=suggestions,
    )


def _operation_matches_shape(operation: ReplaceTextOperation, shapes: list[dict[str, Any]]) -> bool:
    selector = operation.shape_selector

    for shape in shapes:
        if selector.shape_id is not None and shape.get("shape_id") == selector.shape_id:
            return True

        if selector.shape_name is not None and shape.get("shape_name") == selector.shape_name:
            return True

        if selector.contains_text:
            text_preview = str(shape.get("text_preview") or "").lower()
            if selector.contains_text.lower() in text_preview:
                return True

    return False
