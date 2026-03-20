"""
Seed script: creates tables and seeds initial data.
Run: python seed.py
"""
import asyncio
import uuid
from app.database import engine, AsyncSessionLocal, Base
from app.models import *  # noqa
from app.models.report_type import ReportType
from app.models.user import User
from app.utils.auth import hash_password
from sqlalchemy import select


async def seed():
    # Create all tables
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)

    async with AsyncSessionLocal() as db:
        # Seed report types
        types = [
            ("alpha", "Alpha", 1),
            ("beta", "Beta", 2),
            ("gamma", "Gamma", 3),
            ("theta", "Theta", 4),
        ]
        for name, label, order in types:
            existing = await db.execute(select(ReportType).where(ReportType.name == name))
            if not existing.scalar_one_or_none():
                db.add(ReportType(id=uuid.uuid4(), name=name, label=label, sort_order=order))

        # Seed admin user (optional demo user)
        existing_user = await db.execute(select(User).where(User.username == "admin"))
        if not existing_user.scalar_one_or_none():
            admin = User(
                id=uuid.uuid4(),
                username="admin",
                full_name="Admin User",
                employee_id="EMP-0001",
                email="admin@fccs.com",
                phone="+91 XXXXX XXXXX",
                password_hash=hash_password("Admin@12345"),
                is_active=True,
            )
            db.add(admin)

        await db.commit()
        print("[OK] Database seeded successfully!")
        print("   Default user: username=admin, password=Admin@12345")


if __name__ == "__main__":
    asyncio.run(seed())
