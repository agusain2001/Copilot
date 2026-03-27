from datetime import datetime
from sqlalchemy import Integer, DateTime, func
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.database import Base


class ReportSequence(Base):
    """
    Counter table — each row represents a numbered report group folder.
    The integer primary key IS the folder number (1, 2, 3, ...).
    Input and output files that belong to the same job share the same sequence id.
    """
    __tablename__ = "report_sequences"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), nullable=False
    )

    # Relationship back to all reports in this group
    reports = relationship("Report", back_populates="sequence", lazy="dynamic")
