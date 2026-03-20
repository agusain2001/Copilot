import uuid
from datetime import datetime
from pydantic import BaseModel, EmailStr


class UserBase(BaseModel):
    full_name: str
    email: EmailStr
    phone: str | None = None


class UserCreate(UserBase):
    username: str
    employee_id: str
    password: str


class UserUpdate(BaseModel):
    full_name: str | None = None
    email: EmailStr | None = None
    phone: str | None = None


class PasswordChangeRequest(BaseModel):
    current_password: str
    new_password: str
    confirm_password: str


class UserResponse(UserBase):
    id: uuid.UUID
    username: str
    employee_id: str
    profile_photo_url: str | None
    password_changed_at: datetime | None
    is_active: bool
    created_at: datetime

    class Config:
        from_attributes = True


class UserPreferenceResponse(BaseModel):
    theme: str

    class Config:
        from_attributes = True


class UserPreferenceUpdate(BaseModel):
    theme: str  # "dark" | "light"
