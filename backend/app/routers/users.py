from fastapi import APIRouter, Depends, HTTPException, status, UploadFile, File
from fastapi.responses import FileResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from datetime import datetime, timezone
from pathlib import Path
from app.database import get_db
from app.models.user import User
from app.models.user_preference import UserPreference
from app.schemas.user import (
    UserResponse,
    UserUpdate,
    PasswordChangeRequest,
    UserPreferenceResponse,
    UserPreferenceUpdate,
)
from app.utils.auth import verify_password, hash_password
from app.utils.files import save_upload_file, ALLOWED_IMAGE_EXTENSIONS
from app.dependencies import get_current_user

router = APIRouter(prefix="/api/users", tags=["Users"])


@router.get("/me", response_model=UserResponse)
async def get_me(current_user: User = Depends(get_current_user)):
    return current_user


@router.put("/me", response_model=UserResponse)
async def update_me(
    payload: UserUpdate,
    current_user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    if payload.full_name is not None:
        current_user.full_name = payload.full_name
    if payload.email is not None:
        current_user.email = payload.email
    if payload.phone is not None:
        current_user.phone = payload.phone

    await db.commit()
    await db.refresh(current_user)
    return current_user


@router.put("/me/password")
async def change_password(
    payload: PasswordChangeRequest,
    current_user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    if not verify_password(payload.current_password, current_user.password_hash):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Current password is incorrect")

    if payload.new_password != payload.confirm_password:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="New passwords do not match")

    if len(payload.new_password) < 8:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Password must be at least 8 characters")

    current_user.password_hash = hash_password(payload.new_password)
    current_user.password_changed_at = datetime.now(timezone.utc)
    await db.commit()

    return {"message": "Password changed successfully"}


@router.post("/me/photo")
async def upload_photo(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_IMAGE_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Only .jpg, .jpeg, .png files are allowed")

    file_url, _ = await save_upload_file(file, subfolder="photos")
    current_user.profile_photo_url = file_url
    await db.commit()

    return {"profile_photo_url": file_url}


@router.get("/me/photo")
async def get_photo(current_user: User = Depends(get_current_user)):
    if not current_user.profile_photo_url:
        raise HTTPException(status_code=404, detail="No profile photo set")
    path = Path(current_user.profile_photo_url)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Photo file not found")
    return FileResponse(path)


@router.get("/me/preferences", response_model=UserPreferenceResponse)
async def get_preferences(
    current_user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(
        select(UserPreference).where(UserPreference.user_id == current_user.id)
    )
    pref = result.scalar_one_or_none()
    if not pref:
        return UserPreferenceResponse(theme="dark")
    return pref


@router.put("/me/preferences", response_model=UserPreferenceResponse)
async def update_preferences(
    payload: UserPreferenceUpdate,
    current_user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    if payload.theme not in ("dark", "light"):
        raise HTTPException(status_code=400, detail="Theme must be 'dark' or 'light'")

    result = await db.execute(
        select(UserPreference).where(UserPreference.user_id == current_user.id)
    )
    pref = result.scalar_one_or_none()

    if pref:
        pref.theme = payload.theme
    else:
        pref = UserPreference(user_id=current_user.id, theme=payload.theme)
        db.add(pref)

    await db.commit()
    await db.refresh(pref)
    return pref
