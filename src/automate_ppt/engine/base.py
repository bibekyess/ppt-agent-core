from __future__ import annotations

from typing import Protocol

from automate_ppt.schemas import Operation


class PresentationEngine(Protocol):
    def open(self, presentation_path: str) -> None: ...

    def apply(self, slide_number: int, operations: list[Operation]) -> None: ...

    def save(self, output_path: str) -> str: ...

    def close(self) -> None: ...
