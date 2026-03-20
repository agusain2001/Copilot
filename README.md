# FCCS Copilot

> **Foreign Currency Conversion System — Report Management Platform**

A full-stack web application for managing, uploading, filtering, downloading, and organizing financial exchange reports. Built with **React + Vite** on the frontend and **FastAPI + PostgreSQL** on the backend.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Tech Stack](#tech-stack)
- [Project Structure](#project-structure)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Backend Setup](#backend-setup)
  - [Frontend Setup](#frontend-setup)
- [Environment Variables](#environment-variables)
- [API Reference](#api-reference)
- [Default Credentials](#default-credentials)
- [Authentication Flow](#authentication-flow)
- [Report Types](#report-types)

---

## Overview

FCCS Copilot is an internal tool designed for teams working with foreign currency conversion reports. It provides a clean, secure interface to:

- Upload Excel / CSV reports categorized by type (Alpha, Beta, Gamma, Theta)
- Filter, search, and sort reports by name, type, or date
- Download individual reports or bulk-export multiple reports as a ZIP
- Soft-delete reports with a 5-second undo window
- Manage user profile, password, and theme preferences

---

## Features

| Feature | Description |
|---|---|
| 🔐 JWT Authentication | Stateless access + refresh token system |
| 📁 Report Upload | Upload `.csv`, `.xlsx`, `.xls` files (up to 50 MB) |
| 🗂 Report Categorization | Classify reports as Alpha, Beta, Gamma, or Theta |
| 🔍 Search & Filter | Filter by name, report type, and upload date |
| 📦 Bulk Actions | Bulk delete, bulk restore, and bulk export as ZIP |
| ↩️ Undo Delete | 5-second undo toast notification after deletion |
| 👤 User Profile | Edit profile info, change password, upload photo |
| 🌗 Theme Toggle | Dark / Light mode with server-persisted preference |
| 🛡 Protected Routes | All report routes require a valid JWT |
| 📊 Auto-seeded DB | Default report types and admin user seeded on first run |

---

## Tech Stack

### Frontend
| Technology | Version | Purpose |
|---|---|---|
| React | 19.x | UI framework |
| Vite | 8.x | Build tool & dev server |
| React Router DOM | 7.x | Client-side routing |
| Axios | 1.x | HTTP client |
| React Icons | 5.x | Icon library |
| React DatePicker | 9.x | Date filter UI |
| date-fns | 4.x | Date formatting |
| TypeScript (dev) | 5.x | Type checking |

### Backend
| Technology | Version | Purpose |
|---|---|---|
| FastAPI | 0.115 | API framework |
| Uvicorn | 0.30 | ASGI server |
| SQLAlchemy (async) | 2.0 | ORM |
| asyncpg | 0.30 | Async PostgreSQL driver |
| Alembic | 1.13 | Database migrations |
| Pydantic v2 | 2.10 | Data validation & settings |
| python-jose | 3.3 | JWT encoding/decoding |
| passlib + bcrypt | 1.7 / 4.0 | Password hashing |
| aiofiles | 24.1 | Async file I/O |
| PostgreSQL | 15+ | Relational database |

---

## Project Structure

```
FCCS/
├── backend/
│   ├── app/
│   │   ├── main.py              # FastAPI app entrypoint
│   │   ├── config.py            # Settings (from .env)
│   │   ├── database.py          # Async SQLAlchemy engine
│   │   ├── dependencies.py      # JWT dependency injection
│   │   ├── models/              # SQLAlchemy ORM models
│   │   │   ├── user.py
│   │   │   ├── report.py
│   │   │   ├── report_type.py
│   │   │   ├── user_preference.py
│   │   │   └── exchange_rate.py
│   │   ├── routers/             # FastAPI route handlers
│   │   │   ├── auth.py          # /api/auth
│   │   │   ├── reports.py       # /api/reports
│   │   │   ├── users.py         # /api/users
│   │   │   └── report_types.py  # /api/report-types
│   │   ├── schemas/             # Pydantic request/response models
│   │   └── utils/
│   │       ├── auth.py          # Password hashing, JWT helpers
│   │       └── files.py         # File upload/delete helpers
│   ├── alembic/                 # Database migrations
│   ├── alembic.ini
│   ├── seed.py                  # Seeds DB with default data
│   ├── requirements.txt
│   └── .env                     # (not committed) env vars
│
├── frontend/
│   ├── src/
│   │   ├── App.jsx              # Route definitions
│   │   ├── main.jsx             # React entry point
│   │   ├── pages/
│   │   │   ├── LoginPage.jsx    # Login UI
│   │   │   └── ReportsPage.jsx  # Main reports dashboard
│   │   ├── components/
│   │   │   ├── Navbar.jsx       # Top navigation bar
│   │   │   ├── AddReportModal.jsx
│   │   │   ├── DeleteConfirmModal.jsx
│   │   │   ├── DownloadConfirmModal.jsx
│   │   │   ├── EditProfileModal.jsx
│   │   │   ├── ProtectedRoute.jsx
│   │   │   └── Toast.jsx        # Undo-toast notification
│   │   ├── context/
│   │   │   ├── AuthContext.jsx  # Global auth state
│   │   │   └── ThemeContext.jsx # Dark/light theme state
│   │   └── services/
│   │       └── api.js           # Axios instance + interceptors
│   ├── public/
│   ├── index.html
│   ├── vite.config.js
│   └── package.json
│
├── .gitignore
└── README.md
```

---

## Getting Started

### Prerequisites

- **Node.js** 18+ and **npm**
- **Python** 3.11+
- **PostgreSQL** 15+ running locally or remotely
- **Git**

---

### Backend Setup

```bash
# 1. Navigate to the backend directory
cd backend

# 2. Create and activate a virtual environment
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Create the .env file (see Environment Variables section below)
copy .env.example .env    # Windows
# OR
cp .env.example .env      # macOS / Linux

# 5. Seed the database (creates tables + default data)
python seed.py

# 6. Start the development server
uvicorn app.main:app --reload --port 8000
```

The API will be available at: `http://localhost:8000`  
Interactive API docs: `http://localhost:8000/docs`

---

### Frontend Setup

```bash
# 1. Navigate to the frontend directory
cd frontend

# 2. Install dependencies
npm install

# 3. Start the development server
npm run dev
```

The frontend will be available at: `http://localhost:5173`

---

## Environment Variables

Create a `.env` file inside the `backend/` directory with the following variables:

```env
# PostgreSQL async connection string
DATABASE_URL=postgresql+asyncpg://username:password@localhost:5432/fccs_db

# JWT secret key (change this in production!)
SECRET_KEY=your-super-secret-key-change-in-production

# JWT settings (optional — defaults shown)
ALGORITHM=HS256
ACCESS_TOKEN_EXPIRE_MINUTES=30
REFRESH_TOKEN_EXPIRE_DAYS=7

# File upload settings (optional — defaults shown)
UPLOAD_DIR=uploads
MAX_FILE_SIZE_MB=50
```

> ⚠️ **Never commit your `.env` file to version control.** It is listed in `.gitignore`.

---

## API Reference

### Auth — `/api/auth`

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/auth/login` | Login with username + password, returns JWT tokens |
| `POST` | `/api/auth/refresh` | Exchange a refresh token for a new access token |
| `POST` | `/api/auth/logout` | Stateless logout (client discards tokens) |

### Reports — `/api/reports`

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/api/reports` | List reports (supports `search`, `type_name`, `date_from`, `date_to`, `sort_by`) |
| `POST` | `/api/reports` | Upload a new report (multipart form: `name`, `type_name`, `file`) |
| `GET` | `/api/reports/{id}` | Get a single report by ID |
| `DELETE` | `/api/reports/{id}` | Soft-delete a report |
| `POST` | `/api/reports/{id}/restore` | Restore a soft-deleted report |
| `GET` | `/api/reports/{id}/download` | Download a report file |
| `POST` | `/api/reports/bulk-delete` | Soft-delete multiple reports |
| `POST` | `/api/reports/bulk-restore` | Restore multiple soft-deleted reports |
| `POST` | `/api/reports/bulk-export` | Download multiple reports as a `.zip` file |

### Users — `/api/users`

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/api/users/me` | Get current user's profile |
| `PUT` | `/api/users/me` | Update profile (name, email, phone) |
| `PUT` | `/api/users/me/password` | Change password |
| `POST` | `/api/users/me/photo` | Upload profile photo |
| `GET` | `/api/users/me/photo` | Retrieve profile photo |
| `GET` | `/api/users/me/preferences` | Get theme preferences |
| `PUT` | `/api/users/me/preferences` | Update theme (`dark` or `light`) |

### Report Types — `/api/report-types`

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/api/report-types` | List all active report types |

---

## Default Credentials

After running `python seed.py`, the following default admin account is created:

| Field | Value |
|---|---|
| Username | `admin` |
| Password | `Admin@12345` |
| Email | `admin@fccs.com` |
| Employee ID | `EMP-0001` |

> 🔐 **Change the admin password immediately after your first login in a production environment.**

---

## Authentication Flow

```
1. Client  →  POST /api/auth/login  →  { access_token, refresh_token }
2. Client stores tokens (localStorage / sessionStorage based on "Remember Me")
3. Client attaches: Authorization: Bearer <access_token> on every request
4. On 401 Unauthorized, the Axios interceptor automatically:
      - calls POST /api/auth/refresh with the refresh_token
      - retries the original request with the new access_token
5. On logout, tokens are cleared from storage
```

---

## Report Types

The system comes pre-seeded with 4 report type categories:

| Name | Label | Description |
|---|---|---|
| `alpha` | Alpha | Type A reports |
| `beta` | Beta | Type B reports |
| `gamma` | Gamma | Type C reports |
| `theta` | Theta | Type D reports |

Accepted file formats for reports: `.csv`, `.xlsx`, `.xls`

---

## License

This project is proprietary and intended for internal use within the FCCS organization.
