import os
import uuid
import aiofiles
from pathlib import Path
from fastapi import UploadFile, HTTPException
from app.config import settings

ALLOWED_REPORT_EXTENSIONS = {".csv", ".xlsx", ".xls"}
ALLOWED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}


async def save_upload_file(file: UploadFile, subfolder: str = "") -> tuple[str, int]:
    """Save an uploaded file. Returns (file_url, file_size_bytes)."""
    upload_dir = Path(settings.UPLOAD_DIR) / subfolder
    upload_dir.mkdir(parents=True, exist_ok=True)

    ext = Path(file.filename).suffix.lower()
    unique_name = f"{uuid.uuid4()}{ext}"
    file_path = upload_dir / unique_name

    content = await file.read()
    file_size = len(content)

    max_bytes = settings.MAX_FILE_SIZE_MB * 1024 * 1024
    if file_size > max_bytes:
        raise HTTPException(status_code=413, detail=f"File too large. Max {settings.MAX_FILE_SIZE_MB}MB.")

    async with aiofiles.open(file_path, "wb") as f:
        await f.write(content)

    return str(file_path), file_size


def delete_file(file_url: str) -> None:
    """Delete a file from disk if it exists."""
    try:
        path = Path(file_url)
        if path.exists():
            path.unlink()
    except Exception:
        pass
