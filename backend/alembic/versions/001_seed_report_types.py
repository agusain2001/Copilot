"""Seed initial report types

Revision ID: 001_seed_report_types
Revises:
Create Date: 2026-03-20
"""
from alembic import op
import sqlalchemy as sa
import uuid

revision = "001_seed_report_types"
down_revision = None
branch_labels = None
depends_on = None


def upgrade():
    op.execute("""
        CREATE TABLE IF NOT EXISTS report_types (
            id UUID PRIMARY KEY,
            name VARCHAR(50) UNIQUE NOT NULL,
            label VARCHAR(100) NOT NULL,
            sort_order INT NOT NULL DEFAULT 0,
            is_active BOOLEAN NOT NULL DEFAULT TRUE
        )
    """)

    report_types = [
        (str(uuid.uuid4()), "alpha", "Alpha", 1),
        (str(uuid.uuid4()), "beta", "Beta", 2),
        (str(uuid.uuid4()), "gamma", "Gamma", 3),
        (str(uuid.uuid4()), "theta", "Theta", 4),
    ]

    for rt_id, name, label, order in report_types:
        op.execute(f"""
            INSERT INTO report_types (id, name, label, sort_order, is_active)
            VALUES ('{rt_id}', '{name}', '{label}', {order}, TRUE)
            ON CONFLICT (name) DO NOTHING
        """)


def downgrade():
    op.execute("DELETE FROM report_types WHERE name IN ('alpha', 'beta', 'gamma', 'theta')")
