# FCCS Project — Implementation Tasks

## Phase 1: Project Setup
- [ ] Initialize React frontend with Vite
- [ ] Initialize FastAPI backend project structure
- [ ] Configure PostgreSQL connection (Neon DB)
- [ ] Define database models (users, reports, report_types, user_preferences, shops, exchange_rates)
- [ ] Run Alembic migrations & seed report_types

## Phase 2: Backend — Authentication & User Management
- [ ] User model (with password_changed_at for "Last changed X days ago")
- [ ] Auth endpoints: POST `/api/auth/login`, POST `/api/auth/logout`
- [ ] JWT token management (access + refresh tokens)
- [ ] Profile endpoints: GET/PUT `/api/users/me`, PUT `/api/users/me/password` (current + new + confirm)
- [ ] Profile photo upload endpoint
- [ ] User preferences endpoints: GET/PUT `/api/users/me/preferences` (theme)

## Phase 3: Backend — Reports Module
- [ ] Report model (id, name, type, uploaded_time, file_url, uploaded_by, is_deleted, deleted_at)
- [ ] CRUD endpoints: GET/POST/DELETE `/api/reports`
- [ ] File upload (CSV/Excel) with storage
- [ ] Soft-delete + restore endpoints (single & bulk) for undo support
- [ ] Sort by type, sort by date filters
- [ ] Search by report name
- [ ] Bulk delete & bulk export (ZIP) endpoints

## Phase 4: Backend — Shops & Currency (Future-Ready)
- [ ] Shop model (id, name, country, currency_code)
- [ ] Currency conversion utility placeholder
- [ ] Exchange rate model (from_currency, to_currency, rate, date)

## Phase 5: Frontend — Core Setup & Design System
- [ ] Vite + React project scaffolding
- [ ] Dual-theme design system: CSS custom properties for dark + light themes
- [ ] ThemeContext (toggle, persist via API)
- [ ] Routing setup (React Router)
- [ ] Axios API service layer
- [ ] Auth context & protected routes

## Phase 6: Frontend — Login Page
- [ ] Login page UI (dark bg, curved accent lines, centered card)
- [ ] FCCS Copilot logo & branding
- [ ] Username, password fields, "Keep me logged in", login button
- [ ] Login API integration & token storage

## Phase 7: Frontend — Reports Dashboard
- [ ] Top navbar (logo, dark/light toggle, user avatar dropdown)
- [ ] "All Reports" page with data table
- [ ] Empty state ("No Reports Added Yet")
- [ ] Two-step "+ Add Report" upload modal (drag-and-drop → report details)
- [ ] Sort By Type / Sort By Date filter pills
- [ ] Search bar
- [ ] Bulk select (checkboxes), delete (trash icon), export (download icon)
- [ ] Delete confirmation modal (single + bulk) with report name/type list
- [ ] Download confirmation modal (single + bulk) with report name/type list
- [ ] Toast/snackbar notifications (delete with Undo, upload success)
- [ ] Calendar date picker component
- [ ] Type dropdown (Alpha, Beta, Gamma, Theta)
- [ ] File column with click-to-view functionality

## Phase 8: Frontend — User Profile
- [ ] Avatar dropdown menu (name, email, View Profile, Log Out)
- [ ] Edit Profile modal (photo, full name, employee ID, email, phone)
- [ ] Password Management: collapsed state (current pw, "Last changed X days ago", Change Password →)
- [ ] Password Management: expanded state (Create New Password, Confirm New Password, Save Password)
- [ ] Save Changes / Log Out buttons
- [ ] Theme toggle (moon/sun icon) with persistence

## Phase 9: Verification & Polish
- [ ] Run backend tests (pytest)
- [ ] Run frontend build check
- [ ] Browser-based end-to-end flow verification
- [ ] Responsive design review
