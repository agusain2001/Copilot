import os
import aiofiles
from pathlib import Path
from datetime import datetime, timezone
from fastapi import UploadFile, HTTPException
from app.config import settings

ALLOWED_REPORT_EXTENSIONS = {".csv", ".xlsx", ".xls"}
ALLOWED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}


async def save_upload_file(file: UploadFile, subfolder: str = "") -> tuple[str, int]:
    """Save an uploaded file (legacy helper). Returns (file_url, file_size_bytes)."""
    upload_dir = Path(settings.UPLOAD_DIR) / subfolder
    upload_dir.mkdir(parents=True, exist_ok=True)

    import uuid as _uuid
    ext = Path(file.filename).suffix.lower()
    unique_name = f"{_uuid.uuid4()}{ext}"
    file_path = upload_dir / unique_name

    content = await file.read()
    file_size = len(content)

    max_bytes = settings.MAX_FILE_SIZE_MB * 1024 * 1024
    if file_size > max_bytes:
        raise HTTPException(status_code=413, detail=f"File too large. Max {settings.MAX_FILE_SIZE_MB}MB.")

    async with aiofiles.open(file_path, "wb") as f:
        await f.write(content)

    return str(file_path), file_size


async def save_report_file(
    file: UploadFile,
    sequence_id: int,
    file_type: str,  # "input" or "output"
) -> tuple[str, str, str, int]:
    """
    Save a report file using the structured path convention.

    Path structure:
        uploads/reports/<sequence_id>/inputs/<ddmmyyyyhhmmss>_<original_filename>
        uploads/reports/<sequence_id>/outputs/<ddmmyyyyhhmmss>_<original_filename>

    Returns:
        (full_file_path, stored_filename, folder_path, file_size_bytes)
    """
    if file_type not in ("input", "output"):
        raise HTTPException(
            status_code=400,
            detail="file_type must be 'input' or 'output'"
        )

    # Subfolder is "inputs" or "outputs"
    subfolder = f"{file_type}s"

    # Build timestamp prefix: ddmmyyyyhhmmss
    now = datetime.now()  # Use local system time instead of UTC
    timestamp = now.strftime("%d%m%Y%H%M%S")

    original_name = Path(file.filename).name
    stored_filename = f"{timestamp}_{original_name}"

    # Full directory: uploads/reports/<sequence_id>/inputs/
    upload_dir = Path(settings.UPLOAD_DIR) / "reports" / str(sequence_id) / subfolder
    upload_dir.mkdir(parents=True, exist_ok=True)

    file_path = upload_dir / stored_filename

    # Read and size-check
    content = await file.read()
    file_size = len(content)

    max_bytes = settings.MAX_FILE_SIZE_MB * 1024 * 1024
    if file_size > max_bytes:
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Max {settings.MAX_FILE_SIZE_MB}MB."
        )

    async with aiofiles.open(file_path, "wb") as f:
        await f.write(content)

    # Relative folder_path for DB storage (portable across OS)
    folder_path = str(Path("uploads") / "reports" / str(sequence_id) / subfolder / stored_filename)

    return str(file_path), stored_filename, folder_path, file_size


def delete_file(file_url: str) -> None:
    """Delete a file from disk if it exists."""
    try:
        path = Path(file_url)
        if path.exists():
            path.unlink()
    except Exception:
        pass
