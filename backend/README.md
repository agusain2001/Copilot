# FCCS Copilot - Backend

This is the backend for the FCCS Copilot Report Management Platform, built with **FastAPI** and **PostgreSQL**.

## Features
- JWT Authentication (Stateless access + refresh token system)
- Report Upload & Categorization (Alpha, Beta, Gamma, Theta)
- Soft-delete, restoring, and bulk export functionality
- Async PostgreSQL with SQLAlchemy and asyncpg

## Getting Started

### Prerequisites
- Python 3.11+
- PostgreSQL 15+

### Setup

1. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # macOS / Linux
   source venv/bin/activate
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure Environment Variables:
   Create a `.env` file with:
   ```env
   DATABASE_URL=postgresql+asyncpg://username:password@localhost:5432/fccs_db
   SECRET_KEY=your-super-secret-key-change-in-production
   ALGORITHM=HS256
   ACCESS_TOKEN_EXPIRE_MINUTES=30
   REFRESH_TOKEN_EXPIRE_DAYS=7
   UPLOAD_DIR=uploads
   MAX_FILE_SIZE_MB=50
   ```

4. Database Setup:
   ```bash
   python seed.py
   ```

5. Run the server:
   ```bash
   uvicorn app.main:app --reload --port 8000
   ```

Interactive API docs are available at `http://localhost:8000/docs`.

## Default Credentials
After running `seed.py`, an admin account is available:
- **Username:** `admin`
- **Password:** `Admin@12345`
- **Email:** `admin@fccs.com`
