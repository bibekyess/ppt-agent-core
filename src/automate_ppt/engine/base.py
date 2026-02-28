from __future__ import annotations

from typing import Any, Protocol

from automate_ppt.schemas import Operation


class PresentationEngine(Protocol):
    def open(self, presentation_path: str) -> None: ...

    def inspect_slide(self, slide_number: int) -> dict[str, Any]: ...

    def apply(self, slide_number: int, operations: list[Operation]) -> None: ...

    def save(self, output_path: str) -> str: ...

    def close(self) -> None: ...
