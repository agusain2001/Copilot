"""Add report_sequences table and new columns to reports

Revision ID: 002_add_report_sequence_fields
Revises: 001_seed_report_types
Create Date: 2026-03-20
"""
from alembic import op
import sqlalchemy as sa

revision = "002_add_report_sequence_fields"
down_revision = "001_seed_report_types"
branch_labels = None
depends_on = None


def upgrade():
    # 1. Create the report_sequences counter table
    op.create_table(
        "report_sequences",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            server_default=sa.text("now()"),
            nullable=False,
        ),
        sa.PrimaryKeyConstraint("id"),
    )

    # 2. Add new columns to the reports table
    op.add_column("reports", sa.Column("sequence_id", sa.Integer(), nullable=True))
    op.add_column("reports", sa.Column("file_type", sa.String(length=10), nullable=True))
    op.add_column("reports", sa.Column("stored_filename", sa.String(length=400), nullable=True))
    op.add_column("reports", sa.Column("folder_path", sa.String(length=700), nullable=True))

    # 3. Add FK from reports.sequence_id -> report_sequences.id
    op.create_foreign_key(
        "fk_reports_sequence_id",
        "reports",
        "report_sequences",
        ["sequence_id"],
        ["id"],
        ondelete="SET NULL",
    )

    # 4. Index on reports.sequence_id for fast group lookups
    op.create_index("ix_reports_sequence_id", "reports", ["sequence_id"])


def downgrade():
    op.drop_index("ix_reports_sequence_id", table_name="reports")
    op.drop_constraint("fk_reports_sequence_id", "reports", type_="foreignkey")
    op.drop_column("reports", "folder_path")
    op.drop_column("reports", "stored_filename")
    op.drop_column("reports", "file_type")
    op.drop_column("reports", "sequence_id")
    op.drop_table("report_sequences")
