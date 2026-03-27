import uuid
from datetime import datetime
from pydantic import BaseModel


class ReportTypeResponse(BaseModel):
    id: uuid.UUID
    name: str
    label: str
    sort_order: int

    class Config:
        from_attributes = True


class ReportResponse(BaseModel):
    id: uuid.UUID
    name: str
    type_label: str        # Resolved from report_type.label
    type_name: str         # Resolved from report_type.name
    original_filename: str
    file_size_bytes: int | None
    uploaded_at: datetime
    uploaded_by_name: str  # Resolved from uploader.full_name
    is_deleted: bool

    # Structured-storage fields
    sequence_id: int | None = None
    file_type: str | None = None           # "input" or "output"
    stored_filename: str | None = None     # e.g. "20032026122900_report.xls"
    folder_path: str | None = None         # e.g. "uploads/reports/3/inputs/20032026122900_report.xls"

    class Config:
        from_attributes = True


class BulkActionRequest(BaseModel):
    report_ids: list[uuid.UUID]
