from __future__ import annotations

import difflib
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
        abs_path = os.path.abspath(presentation_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Presentation not found: {presentation_path}")

        win32 = self._ensure_runtime()

        try:
            self._ppt = win32.GetActiveObject("PowerPoint.Application")
        except Exception:
            self._ppt = win32.Dispatch("PowerPoint.Application")

        self._ppt.Visible = True

        existing = self._find_open_presentation(abs_path)
        if existing is not None:
            self._presentation = existing
            return

        self._presentation = self._ppt.Presentations.Open(abs_path)

    def inspect_slide(self, slide_number: int) -> dict[str, Any]:
        if self._presentation is None:
            raise RuntimeError("No presentation is open.")

        slide = self._presentation.Slides(slide_number)
        shapes: list[dict[str, Any]] = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            text = self._shape_text(shape)
            shapes.append(
                {
                    "shape_id": shape.Id,
                    "shape_name": shape.Name,
                    "has_text": bool(text),
                    "text_preview": text[:200] if text else None,
                }
            )

        return {"slide": slide_number, "shape_count": len(shapes), "shapes": shapes}

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
        # Intentionally no-op: keep PowerPoint/presentation open across operations.
        return

    def _find_open_presentation(self, abs_path: str) -> Any | None:
        if self._ppt is None:
            return None

        basename = os.path.basename(abs_path)
        for idx in range(1, self._ppt.Presentations.Count + 1):
            presentation = self._ppt.Presentations(idx)
            full_name = ""
            try:
                full_name = os.path.abspath(str(presentation.FullName))
            except Exception:
                pass

            if full_name and os.path.normcase(full_name) == os.path.normcase(abs_path):
                return presentation

            if str(presentation.Name) == basename:
                return presentation

        return None

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

        all_shapes = [slide.Shapes(i) for i in range(1, slide.Shapes.Count + 1)]

        if selector.shape_name is not None:
            for shape in all_shapes:
                if shape.Name == selector.shape_name:
                    return shape

        if selector.contains_text:
            needle = selector.contains_text.lower()
            for shape in all_shapes:
                hay = (self._shape_text(shape) or "").lower()
                if needle in hay:
                    return shape

        requested_name = selector.shape_name or ""
        available_names = [shape.Name for shape in all_shapes]
        suggestions = difflib.get_close_matches(requested_name, available_names, n=3, cutoff=0.3)
        msg = "Shape not found for selector."
        if suggestions:
            msg += f" Did you mean one of: {', '.join(suggestions)}"
        raise ValueError(msg)

    @staticmethod
    def _shape_text(shape: Any) -> str | None:
        try:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                return str(shape.TextFrame.TextRange.Text)
        except Exception:
            return None
        return None
