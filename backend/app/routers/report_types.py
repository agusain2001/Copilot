from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from app.database import get_db
from app.models.report_type import ReportType
from app.schemas.report import ReportTypeResponse
from app.dependencies import get_current_user
from app.models.user import User

router = APIRouter(prefix="/api/report-types", tags=["Report Types"])


@router.get("", response_model=list[ReportTypeResponse])
async def list_report_types(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(ReportType).where(ReportType.is_active == True).order_by(ReportType.sort_order)
    )
    return result.scalars().all()
