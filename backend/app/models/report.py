import uuid
from datetime import datetime
from sqlalchemy import String, Boolean, DateTime, BigInteger, Integer, ForeignKey, func
from sqlalchemy.orm import Mapped, mapped_column, relationship
from sqlalchemy.dialects.postgresql import UUID
from app.database import Base


class Report(Base):
    __tablename__ = "reports"

    id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), primary_key=True, default=uuid.uuid4
    )
    name: Mapped[str] = mapped_column(String(300), nullable=False, index=True)
    type_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("report_types.id", ondelete="RESTRICT"), nullable=False
    )
    file_url: Mapped[str] = mapped_column(String, nullable=False)
    original_filename: Mapped[str] = mapped_column(String(300), nullable=False)
    file_size_bytes: Mapped[int | None] = mapped_column(BigInteger, nullable=True)
    uploaded_by: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False
    )
    uploaded_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), nullable=False, index=True
    )
    is_deleted: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False, index=True)
    deleted_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)

    # ── New structured-storage fields ──────────────────────────────────────────
    sequence_id: Mapped[int | None] = mapped_column(
        Integer, ForeignKey("report_sequences.id", ondelete="SET NULL"), nullable=True, index=True
    )
    file_type: Mapped[str | None] = mapped_column(
        String(10), nullable=True
    )  # "input" or "output"
    stored_filename: Mapped[str | None] = mapped_column(
        String(400), nullable=True
    )  # e.g. "20032026122900_report.xls"
    folder_path: Mapped[str | None] = mapped_column(
        String(700), nullable=True
    )  # e.g. "uploads/reports/3/inputs/20032026122900_report.xls"
    # ────────────────────────────────────────────────────────────────────────────

    # Relationships
    report_type = relationship("ReportType", lazy="joined")
    uploader = relationship("User", lazy="joined")
    sequence = relationship("ReportSequence", back_populates="reports", lazy="joined")
