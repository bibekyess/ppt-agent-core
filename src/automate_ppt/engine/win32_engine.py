from __future__ import annotations

import os
from typing import Any

from automate_ppt.schemas import Operation, ReplaceTextOperation


class Win32PresentationEngine:
    def __init__(self) -> None:
        self._ppt: Any | None = None
        self._presentation: Any | None = None

    def _ensure_runtime(self):
        try:
            import win32com.client  # type: ignore
        except Exception as exc:  # pragma: no cover - platform dependent
            raise RuntimeError(
                "win32com is required and only works on Windows with PowerPoint installed."
            ) from exc
        return win32com.client

    def open(self, presentation_path: str) -> None:
        if not os.path.exists(presentation_path):
            raise FileNotFoundError(f"Presentation not found: {presentation_path}")

        win32 = self._ensure_runtime()
        self._ppt = win32.Dispatch("PowerPoint.Application")
        self._ppt.Visible = True
        self._presentation = self._ppt.Presentations.Open(os.path.abspath(presentation_path))

    def apply(self, slide_number: int, operations: list[Operation]) -> None:
        if self._presentation is None:
            raise RuntimeError("No presentation is open.")

        slide = self._presentation.Slides(slide_number)
        for operation in operations:
            self._apply_operation(slide, operation)

    def save(self, output_path: str) -> str:
        if self._presentation is None:
            raise RuntimeError("No presentation is open.")

        final_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(final_path), exist_ok=True)
        self._presentation.SaveAs(final_path)
        return final_path

    def close(self) -> None:
        if self._presentation is not None:
            self._presentation.Close()
            self._presentation = None
        if self._ppt is not None:
            self._ppt.Quit()
            self._ppt = None

    def _apply_operation(self, slide: Any, operation: Operation) -> None:
        if isinstance(operation, ReplaceTextOperation):
            shape = self._find_shape(slide, operation)
            shape.TextFrame.TextRange.Text = operation.value
            return

        raise ValueError(f"Unsupported operation type: {operation.type}")

    def _find_shape(self, slide: Any, operation: ReplaceTextOperation) -> Any:
        selector = operation.shape_selector
        if selector.shape_id is not None:
            return slide.Shapes(selector.shape_id)

        if selector.shape_name is not None:
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                if shape.Name == selector.shape_name:
                    return shape

        raise ValueError("Shape not found for selector.")
