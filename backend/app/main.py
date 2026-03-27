from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from app.config import settings
from app.routers import auth_router, users_router, reports_router, report_types_router, processing_router


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Create upload directories on startup
    uploads = Path(settings.UPLOAD_DIR)
    uploads.mkdir(parents=True, exist_ok=True)
    (uploads / "reports").mkdir(exist_ok=True)
    (uploads / "photos").mkdir(exist_ok=True)

    # Mount static files now that directory is guaranteed to exist
    app.mount("/uploads", StaticFiles(directory=str(uploads)), name="uploads")
    yield


app = FastAPI(
    title="FCCS Copilot API",
    description="Foreign Currency Conversion System — Report Management Platform",
    version="1.0.0",
    lifespan=lifespan,
)

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.BACKEND_CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include all routers
app.include_router(auth_router)
app.include_router(users_router)
app.include_router(reports_router)
app.include_router(report_types_router)
app.include_router(processing_router)


@app.get("/")
async def root():
    return {"message": "FCCS Copilot API is running", "version": "1.0.0"}


@app.get("/health")
async def health():
    return {"status": "healthy"}
