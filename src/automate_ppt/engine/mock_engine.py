from __future__ import annotations

import json
import os
from typing import Any

from automate_ppt.schemas import Operation


class MockPresentationEngine:
    """Fallback engine for non-Windows development/testing.

    It does not edit actual PPT files; it writes an execution log artifact.
    """

    def __init__(self) -> None:
        self._input_path: str | None = None
        self._operations: list[dict] = []

    def open(self, presentation_path: str) -> None:
        self._input_path = presentation_path

    def inspect_slide(self, slide_number: int) -> dict[str, Any]:
        return {
            "slide": slide_number,
            "shape_count": 2,
            "shapes": [
                {
                    "shape_id": 1,
                    "shape_name": "Title 1",
                    "has_text": True,
                    "text_preview": "Mock slide title",
                },
                {
                    "shape_id": 2,
                    "shape_name": "Content Placeholder 2",
                    "has_text": True,
                    "text_preview": "Mock slide content",
                },
            ],
            "note": "Mock inspection data. Use Windows+PowerPoint for real inspection.",
        }

    def apply(self, slide_number: int, operations: list[Operation]) -> None:
        for operation in operations:
            self._operations.append(
                {
                    "slide": slide_number,
                    "type": operation.type,
                    "payload": operation.model_dump(),
                }
            )

    def save(self, output_path: str) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        data = {
            "input": self._input_path,
            "operations": self._operations,
            "note": "Mock engine output. Run on Windows with win32com for real PPT edits.",
        }
        with open(output_path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)
        return os.path.abspath(output_path)

    def close(self) -> None:
        return
