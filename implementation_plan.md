# FCCS Copilot — Full Implementation Plan

FCCS (Foreign Currency Conversion System) is a multi-shop, multi-country report management platform. Shops across different countries use different currencies (INR, USD, AED, etc.) and headquarters needs consolidated reports with currency conversion. This plan covers the full frontend and backend build, with currency conversion designed as a future-ready module.

## User Review Required

> [!IMPORTANT]
> **Database**: Using the provided Neon PostgreSQL connection string. All tables will be created via Alembic migrations.

> [!IMPORTANT]
> **Report Types**: Based on the screenshots, report types are **Alpha, Beta, Gamma, Theta**. Please confirm if these are the final types or if they should be dynamic/configurable.

> [!WARNING]
> **File Storage**: Reports (CSV/Excel) will be stored on the **local filesystem** (`uploads/` directory) for now. If cloud storage (S3, GCS) is needed, please specify.

> [!IMPORTANT]
> **Currency Conversion**: The currency module is designed as a placeholder. The core CRUD + reporting features will be built first. Conversion logic will be wired in later once the exchange rate source is decided.

---

## Proposed Changes

### 1. Project Scaffolding

#### [NEW] Project Root Structure
```
G:\FCCS\
├── backend/           ← FastAPI application
│   ├── app/
│   │   ├── __init__.py
│   │   ├── main.py
│   │   ├── config.py
│   │   ├── database.py
│   │   ├── models/
│   │   ├── schemas/
│   │   ├── routers/
│   │   ├── services/
│   │   └── utils/
│   ├── alembic/
│   ├── alembic.ini
│   ├── requirements.txt
│   └── uploads/        ← File storage directory
└── frontend/          ← React + Vite application
    ├── src/
    │   ├── assets/
    │   ├── components/
    │   ├── pages/
    │   ├── context/
    │   ├── services/
    │   ├── styles/
    │   ├── App.jsx
    │   └── main.jsx
    ├── public/
    ├── index.html
    ├── vite.config.js
    └── package.json
```

---

### 2. Backend — Database Schema

#### Entity Relationship Diagram

```mermaid
erDiagram
    users ||--o{ reports : uploads
    users ||--o| user_preferences : has
    report_types ||--o{ reports : categorizes
    shops ||--o{ reports : "future: belongs_to"
    exchange_rates }o--|| shops : "future: converts"

    users {
        uuid id PK
        varchar username UK
        varchar full_name
        varchar employee_id UK
        varchar email UK
        varchar phone
        text password_hash
        text profile_photo_url
        timestamp password_changed_at
        boolean is_active
        timestamp created_at
        timestamp updated_at
    }

    reports {
        uuid id PK
        varchar name
        uuid type_id FK
        text file_url
        varchar original_filename
        bigint file_size_bytes
        uuid uploaded_by FK
        timestamp uploaded_at
        boolean is_deleted
        timestamp deleted_at
    }

    report_types {
        uuid id PK
        varchar name UK
        varchar label
        int sort_order
        boolean is_active
    }

    user_preferences {
        uuid id PK
        uuid user_id FK-UK
        varchar theme "light or dark"
        timestamp updated_at
    }

    shops {
        uuid id PK
        varchar name
        varchar country
        varchar currency_code
        boolean is_active
        timestamp created_at
    }

    exchange_rates {
        uuid id PK
        varchar from_currency
        varchar to_currency
        decimal rate
        date effective_date
    }
```

#### [NEW] [models/user.py](file:///G:/FCCS/backend/app/models/user.py)

| Column | Type | Notes |
|---|---|---|
| `id` | UUID (PK) | Auto-generated |
| `username` | VARCHAR(100) | Unique, for login |
| `full_name` | VARCHAR(200) | Display name |
| `employee_id` | VARCHAR(50) | Unique, e.g. "EMP-1024" |
| `email` | VARCHAR(200) | Unique |
| `phone` | VARCHAR(20) | Optional |
| `password_hash` | TEXT | bcrypt hash |
| `profile_photo_url` | TEXT | File path or URL |
| `password_changed_at` | TIMESTAMP | Tracks "Last changed X days ago" |
| `is_active` | BOOLEAN | Default true |
| `created_at` | TIMESTAMP | Auto |
| `updated_at` | TIMESTAMP | Auto |

#### [NEW] [models/report.py](file:///G:/FCCS/backend/app/models/report.py)

| Column | Type | Notes |
|---|---|---|
| `id` | UUID (PK) | Auto-generated |
| `name` | VARCHAR(300) | Report display name |
| `type_id` | UUID (FK → report_types.id) | Links to report_types |
| `file_url` | TEXT | Path to uploaded file |
| `original_filename` | VARCHAR(300) | Original file name |
| `file_size_bytes` | BIGINT | For display (optional) |
| `uploaded_at` | TIMESTAMP | Auto |
| `uploaded_by` | UUID (FK → users.id) | Who uploaded |
| `is_deleted` | BOOLEAN | Default false — soft-delete for undo |
| `deleted_at` | TIMESTAMP | Null until soft-deleted |

#### [NEW] [models/report_type.py](file:///G:/FCCS/backend/app/models/report_type.py)

| Column | Type | Notes |
|---|---|---|
| `id` | UUID (PK) | Auto-generated |
| `name` | VARCHAR(50) | Unique key: alpha, beta, gamma, theta |
| `label` | VARCHAR(100) | Display: "Alpha", "Beta", etc. |
| `sort_order` | INT | For dropdown ordering |
| `is_active` | BOOLEAN | Default true |

Seeded with: Alpha, Beta, Gamma, Theta

#### [NEW] [models/user_preference.py](file:///G:/FCCS/backend/app/models/user_preference.py)

| Column | Type | Notes |
|---|---|---|
| `id` | UUID (PK) | Auto-generated |
| `user_id` | UUID (FK → users.id) | Unique — one row per user |
| `theme` | VARCHAR(10) | `"dark"` or `"light"`, default `"dark"` |
| `updated_at` | TIMESTAMP | Auto |

#### [NEW] [models/shop.py](file:///G:/FCCS/backend/app/models/shop.py) — Future-ready

| Column | Type | Notes |
|---|---|---|
| `id` | UUID (PK) | Auto-generated |
| `name` | VARCHAR(200) | Shop name |
| `country` | VARCHAR(100) | Country name |
| `currency_code` | VARCHAR(10) | ISO code (INR, USD, AED) |
| `is_active` | BOOLEAN | Default true |
| `created_at` | TIMESTAMP | Auto |

#### [NEW] [models/exchange_rate.py](file:///G:/FCCS/backend/app/models/exchange_rate.py) — Future-ready

| Column | Type | Notes |
|---|---|---|
| `id` | UUID (PK) | Auto-generated |
| `from_currency` | VARCHAR(10) | e.g. INR |
| `to_currency` | VARCHAR(10) | e.g. USD |
| `rate` | DECIMAL(18,6) | Conversion rate |
| `effective_date` | DATE | When this rate applies |

---

### 3. Backend — Authentication

#### [NEW] [routers/auth.py](file:///G:/FCCS/backend/app/routers/auth.py)

| Endpoint | Method | Description |
|---|---|---|
| `/api/auth/login` | POST | Login with username + password → returns JWT |
| `/api/auth/logout` | POST | Blacklists the current token |
| `/api/auth/refresh` | POST | Refresh access token |

- **JWT** with `python-jose` — Access token (30 min) + Refresh token (7 days)
- Passwords hashed with `passlib[bcrypt]`
- Dependency: `get_current_user` extracts user from `Authorization: Bearer <token>` header

---

### 4. Backend — User Profile & Preferences

#### [NEW] [routers/users.py](file:///G:/FCCS/backend/app/routers/users.py)

| Endpoint | Method | Description |
|---|---|---|
| `/api/users/me` | GET | Get profile + `password_changed_at` (for "Last changed X days ago") |
| `/api/users/me` | PUT | Update full_name, email, phone |
| `/api/users/me/password` | PUT | Change password: requires `current_password`, `new_password`, `confirm_password`. Updates `password_changed_at`. |
| `/api/users/me/photo` | POST | Upload profile photo (jpg/png/jpeg) |
| `/api/users/me/preferences` | GET | Get user preferences (theme, etc.) |
| `/api/users/me/preferences` | PUT | Update preferences (e.g. `{ "theme": "light" }`) |

#### [NEW] [routers/report_types.py](file:///G:/FCCS/backend/app/routers/report_types.py)

| Endpoint | Method | Description |
|---|---|---|
| `/api/report-types` | GET | List all active report types (for dropdowns) |

---

### 5. Backend — Reports

#### [NEW] [routers/reports.py](file:///G:/FCCS/backend/app/routers/reports.py)

| Endpoint | Method | Description |
|---|---|---|
| `/api/reports` | GET | List reports (query params: `type`, `date_from`, `date_to`, `search`, `sort_by`) |
| `/api/reports` | POST | Upload new report (multipart: `name`, `type`, `file`) |
| `/api/reports/{id}` | GET | Get single report details |
| `/api/reports/{id}` | DELETE | Soft-delete a report (sets `is_deleted=true`) |
| `/api/reports/{id}/restore` | POST | Undo delete — restores a soft-deleted report |
| `/api/reports/bulk-delete` | POST | Soft-delete multiple reports by IDs |
| `/api/reports/bulk-restore` | POST | Undo bulk delete — restores multiple reports |
| `/api/reports/bulk-export` | POST | Download multiple reports as a single ZIP file |
| `/api/reports/{id}/download` | GET | Download a single report file |

---

### 6. Backend — Configuration & Utilities

#### [NEW] [config.py](file:///G:/FCCS/backend/app/config.py)
- `DATABASE_URL` from environment / `.env` file
- `SECRET_KEY`, `ALGORITHM`, `ACCESS_TOKEN_EXPIRE_MINUTES`
- `UPLOAD_DIR` for file storage path
- CORS origins configuration

#### [NEW] [database.py](file:///G:/FCCS/backend/app/database.py)
- SQLAlchemy async engine + session factory
- Neon DB connection with `sslmode=require`

#### [NEW] [main.py](file:///G:/FCCS/backend/app/main.py)
- FastAPI app instantiation
- CORS middleware (allow React dev server)
- Router includes for auth, users, reports
- Lifespan event to create upload directory

#### [NEW] [requirements.txt](file:///G:/FCCS/backend/requirements.txt)
```
fastapi==0.115.0
uvicorn[standard]==0.30.0
sqlalchemy[asyncio]==2.0.35
asyncpg==0.30.0
alembic==1.13.0
python-jose[cryptography]==3.3.0
passlib[bcrypt]==1.7.4
python-multipart==0.0.9
python-dotenv==1.0.1
pydantic[email]==2.9.0
aiofiles==24.1.0
```

---

### 7. Frontend — Design System & Core

#### [NEW] React + Vite Project
- Initialize with `npx -y create-vite@latest ./ --template react`
- Install: `react-router-dom`, `axios`, `react-icons`, `react-datepicker`

#### [NEW] [src/styles/index.css](file:///G:/FCCS/frontend/src/styles/index.css)
Dual-theme system using **CSS custom properties** on `[data-theme]`:

| Token | Dark Theme | Light Theme |
|---|---|---|
| `--bg-primary` | `#1a1a1a` | `#f5f5f5` |
| `--bg-secondary` | `#242424` | `#ffffff` |
| `--bg-card` | `#2a2a2a` | `#ffffff` |
| `--bg-table-row` | `#1e1e1e` | `#fafafa` |
| `--bg-table-row-alt` | `#262626` | `#f0f0f0` |
| `--bg-input` | `#333333` | `#f5f5f5` |
| `--bg-modal-overlay` | `rgba(0,0,0,0.7)` | `rgba(0,0,0,0.4)` |
| `--text-primary` | `#ffffff` | `#1a1a1a` |
| `--text-secondary` | `#999999` | `#666666` |
| `--border-color` | `#3a3a3a` | `#e0e0e0` |
| `--accent-cyan` | `#0891b2` | `#0891b2` |
| `--accent-gold` | `#b8860b` | `#b8860b` |
| `--accent-danger` | `#dc2626` | `#dc2626` |
| `--btn-primary-bg` | `#ffffff` | `#1a1a1a` |
| `--btn-primary-text` | `#1a1a1a` | `#ffffff` |
| `--toast-bg` | `#333333` | `#1a1a1a` |

- Default theme: **dark** (matches screenshots)
- Theme applied via `<html data-theme="dark">` attribute
- All components use `var(--token)` — zero hard-coded colors
- Login page accent curves (teal + gold) remain consistent in both themes

#### [NEW] [src/context/ThemeContext.jsx](file:///G:/FCCS/frontend/src/context/ThemeContext.jsx)
- Reads theme from `GET /api/users/me/preferences` on login
- `toggleTheme()` switches between `"dark"` and `"light"`
- Persists choice via `PUT /api/users/me/preferences`
- Sets `data-theme` attribute on `<html>`

#### [NEW] [src/components/Navbar.jsx](file:///G:/FCCS/frontend/src/components/Navbar.jsx)
- FCCS Copilot logo (left)
- Theme toggle: 🌙 moon icon (dark) ↔ ☀️ sun icon (light)
- User avatar + initials badge → dropdown

#### [NEW] [src/context/AuthContext.jsx](file:///G:/FCCS/frontend/src/context/AuthContext.jsx)
- Stores JWT + user info
- `login()`, `logout()`, `isAuthenticated`
- `ProtectedRoute` wrapper component

#### [NEW] [src/services/api.js](file:///G:/FCCS/frontend/src/services/api.js)
- Axios instance with base URL + interceptor for auth token
- Auto-redirect to login on 401

---

### 8. Frontend — Login Page

#### [NEW] [src/pages/LoginPage.jsx](file:///G:/FCCS/frontend/src/pages/LoginPage.jsx)
Matching screenshot exactly:
- Full-screen dark background with decorative teal + golden curved arcs
- Centered login card with subtle dashed border
- "Log in" heading + "Please login to continue to your account."
- Username input, Password input (with eye toggle)
- "Keep me logged in" checkbox
- "Log in" button (full-width, white)
- FCCS Copilot logo in top-left corner

---

### 9. Frontend — Reports Dashboard

#### [NEW] [src/pages/ReportsPage.jsx](file:///G:/FCCS/frontend/src/pages/ReportsPage.jsx)
- **Header**: "All Reports" title
- **Filter pills**: "Sort By Type" + "Sort By Date" (cyan outlined buttons with filter icon)
- **Action bar**: "+ Add Report" button (top-right), Search icon
- **Data table** columns: Checkbox | Report Name | Type | Uploaded Time | File
- **Empty state**: "No Reports Added Yet" with centered "+ Add Report" button
- **Filled state**: List of reports with checkboxes
- **Bulk actions**: When items selected → show delete (🗑) + export (⬇) icons next to search
- **File column**: Shows filename with "click to view" label and file icon

#### [NEW] [src/components/AddReportModal.jsx](file:///G:/FCCS/frontend/src/components/AddReportModal.jsx)
Two-step upload flow in a dark modal:
- **Step 1 — File Upload**:
  - Title: "Upload New Report"
  - Subtitle: "Upload your report file. Supported formats: .csv, .xlsx, .xls"
  - Select file button **OR** drag-and-drop zone with cloud upload icon
  - Shows selected file name with green checkmark after selection
  - Validation error text if wrong file type (red text)
- **Step 2 — Report Details**:
  - Title: "Upload New Report"
  - Fields: Report Name (text input), Report Type (dropdown: Alpha/Beta/Gamma/Theta)
  - File already uploaded shown below (e.g. "Employeefile.csv" with file icon, clickable)
  - "Add Report" button (bottom-right, cyan/blue accent)
- On success → closes modal + shows success toast

#### [NEW] [src/components/DeleteConfirmModal.jsx](file:///G:/FCCS/frontend/src/components/DeleteConfirmModal.jsx)
Used for **both** single and bulk delete:
- **Single delete**: Title "Delete Report File", red warning icon
  - Message: "This action permanently removes the report files from your list. You won't be able to recover them."
  - Table showing Report Name + Type of the single report
  - Buttons: "Cancel" (ghost) + "Delete" (red)
- **Bulk delete**: Title "Delete Report Files" (plural), same layout
  - Table listing all selected reports (Name + Type)
  - Buttons: "Cancel" (ghost) + "Delete" (red)

#### [NEW] [src/components/DownloadConfirmModal.jsx](file:///G:/FCCS/frontend/src/components/DownloadConfirmModal.jsx)
Used for **both** single and bulk download:
- **Single download**: Title "Download File", blue download icon
  - Message: "The selected report will be downloaded to your device."
  - Table showing Report Name + Type
  - "Download" button (full-width, white)
- **Bulk download**: Title "Download Files" (plural)
  - Message: "The selected report will be downloaded to your device as ZIP file."
  - Table listing all selected reports (Name + Type)
  - "Download" button (full-width, white)

#### [NEW] [src/components/ToastNotification.jsx](file:///G:/FCCS/frontend/src/components/ToastNotification.jsx)
Bottom-center snackbar/toast bar:
- **After single delete**: `Report: Employee Info File Has been deleted.` + "Undo" button + ✕ close
- **After bulk delete**: `3 Reports Has been deleted.` + "Undo" button + ✕ close
- **After upload success**: `Report: Employeefile.csv was saved` + ✕ close
- Dark background, white text, "Undo" button styled as a pill
- Auto-dismisses after ~5 seconds, or user clicks ✕
- "Undo" calls `POST /api/reports/{id}/restore` or `POST /api/reports/bulk-restore`

#### [NEW] [src/components/DatePicker.jsx](file:///G:/FCCS/frontend/src/components/DatePicker.jsx)
- Calendar picker with month navigation (matching dark theme from screenshot)

#### [NEW] [src/components/TypeFilter.jsx](file:///G:/FCCS/frontend/src/components/TypeFilter.jsx)
- Dropdown showing: Alpha, Beta, Gamma, Theta

---

### 10. Frontend — User Profile

#### [NEW] [src/components/UserDropdown.jsx](file:///G:/FCCS/frontend/src/components/UserDropdown.jsx)
- Shows user avatar + name + email
- "View Profile →" link → opens edit profile modal
- "Log Out ↵" button

#### [NEW] [src/components/EditProfileModal.jsx](file:///G:/FCCS/frontend/src/components/EditProfileModal.jsx)
Matching screenshots — centered overlay modal:
- **Header**: "Edit Profile Information" + subtitle "Update your personal details and contact information." + close (✕) button
- **Basic Information** section:
  - Profile photo circle with camera icon overlay + "Upload a File (.jpg, .png, .jpeg)"
  - Full Name (editable), Employee ID (read-only, with copy-to-clipboard icon)
  - Email Address (editable), Phone Number (editable)
- **Password Management** section (two states):
  - **Collapsed (default)**:
    - Title: "Password Management" + subtitle "Update your account password to keep your profile secure."
    - Current Password field (masked, read-only, with eye toggle)
    - "Last changed 45 days ago" text below
    - "Change Password →" button (right side)
  - **Expanded** (after clicking "Change Password →"):
    - Same current password display
    - Two new fields appear below: "Create New Password" + "Confirm New Password"
    - "Save Password" button (bottom-right)
    - Backend validates: current password correct, new == confirm, meets strength rules
- **Footer**: "Log Out" button (red, left) + "Save Changes" button (white, right)

---

## Verification Plan

### Automated Tests

1. **Backend unit tests** (pytest):
   ```bash
   cd G:\FCCS\backend
   pytest tests/ -v
   ```
   - Test auth endpoints (login, token validation, password change)
   - Test report CRUD endpoints
   - Test file upload and download

2. **Frontend build check**:
   ```bash
   cd G:\FCCS\frontend
   npm run build
   ```
   - Ensures no compilation errors

### Browser-Based Verification

1. Start backend: `cd G:\FCCS\backend && uvicorn app.main:app --reload --port 8000`
2. Start frontend: `cd G:\FCCS\frontend && npm run dev`
3. **Login flow**: Navigate to login page → enter credentials → verify redirect to dashboard
4. **Reports CRUD**: Add report (2-step modal) → verify table updates → select all → bulk delete → undo
5. **Profile modal**: Click avatar → View Profile → edit fields → Save Changes
6. **Password change**: Click "Change Password →" → enter new password → Save Password
7. **Theme toggle**: Switch dark ↔ light → verify all pages update → reload page → verify preference persists
8. **Sort & Filter**: Test Sort By Type, Sort By Date, and Search functionality
9. **Download**: Single file download + bulk ZIP download via confirmation modal
10. **Delete + Undo**: Single delete → click Undo → verify restore. Bulk delete → Undo → verify restore.

### Manual Verification
- User visually confirms both **dark** and **light** themes match expected aesthetics
- User confirms CSV/Excel upload and download works correctly
- User confirms password change flow works end-to-end
