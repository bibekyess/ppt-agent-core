from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field


class ShapeSelector(BaseModel):
    shape_id: int | None = None
    shape_name: str | None = None


class ReplaceTextOperation(BaseModel):
    type: Literal["replace_text"]
    shape_selector: ShapeSelector
    value: str


Operation = Annotated[Union[ReplaceTextOperation], Field(discriminator="type")]


class Target(BaseModel):
    slide: int = Field(ge=1)


class EditJobRequest(BaseModel):
    schema_version: str = "1.0"
    presentation_path: str
    output_path: str | None = None
    target: Target
    operations: list[Operation] = Field(min_length=1)


class EditJobResponse(BaseModel):
    job_id: str
    status: str


class JobStatusResponse(BaseModel):
    job_id: str
    status: str
    error: str | None = None
    output_path: str | None = None
