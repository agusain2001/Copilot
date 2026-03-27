from app.routers.auth import router as auth_router
from app.routers.users import router as users_router
from app.routers.reports import router as reports_router
from app.routers.report_types import router as report_types_router
from app.routers.processing import router as processing_router

__all__ = ["auth_router", "users_router", "reports_router", "report_types_router", "processing_router"]
