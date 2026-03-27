# ICM Report Processing — Full-Stack Integration

Integrate the standalone AI IC matching engine ([ic_matching_v3.py](file:///G:/FCCS/AI/ic_matching_v3.py)) into the existing FCCS backend/frontend so users can upload 5 Excel files, trigger server-side processing, and download the enriched ICM output — all from the web UI.

## User Review Required

> [!IMPORTANT]
> **File naming convention**: The user's screenshot shows 5 expected files: `report_inputs.xlsx`, `icm_report.xlsx`, `parent_journal.xlsx`, `contribution_journal.xlsx`, `plugaccount_journal.xlsx`. The current AI script expects `Intercompany Balances IC Matching Report (1).xlsx` and `Journal Report (1/2/4).xlsx`. The new integration will use the **user's file names** as the canonical mapping.

> [!WARNING]
> **openpyxl dependency**: The backend currently does not have `openpyxl` in [requirements.txt](file:///G:/FCCS/backend/requirements.txt). It will be added. This is a non-breaking addition.

---

## Proposed Changes

### 1. AI Processing Module (Backend Integration)

Move the IC matching logic from a standalone script into a clean, importable Python module under `backend/app/`.

#### [NEW] [ic_processor.py](file:///G:/FCCS/backend/app/ic_processor.py)

A refactored version of `AI/ic_matching_v3.py` designed to be called from the backend:

- **`process_icm_report(icm_path, journal_paths, output_path) → str`** — Main entry point
  - Accepts file paths for the ICM report and a dict of journal files `{"parent": path, "contribution": path, "plugaccount": path}`
  - Calls the existing 6-module pipeline: read ICM headers → read journals → match → build columns → write output
  - Returns the output file path
- **Journal label mapping**: Maps `parent_journal.xlsx` → Journal 1 (Parent Input), `contribution_journal.xlsx` → Journal 2 (Contribution Input), `plugaccount_journal.xlsx` → Journal 4 (Plug Account)
- **Variance calculation** (as described in transcript):
  - **Per journal block**: Variance 1 = Σ(Series1-Entity) − Σ(Series2-Partner); Variance 2 = Σ(Series1-Partner) − Σ(Series2-Entity); Total = Variance1 + Variance2
  - **Plug Account block**: Has only one value column; Total column combines all 3 journal totals (Parent Total + Contribution Total + Plug Account Total)
- **Plug Account Mapping**: Parses `report_inputs.xlsx` to identify the designated Plug Account code and a list of Eliminating Account codes. For the Plug Account Journal ONLY, any transaction hitting an Eliminating Account is remapped to the Plug Account so it aggregates correctly in the final output.
- All print statements replaced with Python `logging`

---

### 2. Backend API — New Processing Router

#### [NEW] [processing.py](file:///G:/FCCS/backend/app/routers/processing.py)

New FastAPI router at `/api/processing` with two endpoints:

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/processing/run` | `POST` | Accept 5 uploaded files, validate, run IC matching, save output, return result |
| `/api/processing/{sequence_id}/output` | `GET` | Download the generated output file |

**`POST /api/processing/run`** flow:
1. Accept `multipart/form-data` with 5 files: `icm_report`, `parent_journal`, `contribution_journal`, `plugaccount_journal`, `report_inputs` (plus `name` and `type_name` form fields)
2. Validate all files are `.xlsx`
3. Create a new `ReportSequence` (gets auto-increment ID)
4. Save all 5 input files to `uploads/reports/<seq_id>/inputs/`
5. Call `ic_processor.process_icm_report()` with the saved paths
6. Save output file to `uploads/reports/<seq_id>/outputs/`
7. Create `Report` DB records for each input and the output
8. Return JSON response with `sequence_id`, output `report_id`, processing status, and download URL

#### [MODIFY] [__init__.py](file:///G:/FCCS/backend/app/routers/__init__.py)

Register the new `processing_router`.

#### [MODIFY] [main.py](file:///G:/FCCS/backend/app/main.py)

Include `processing_router` in the app.

#### [MODIFY] [requirements.txt](file:///G:/FCCS/backend/requirements.txt)

Add `openpyxl>=3.1.0` dependency.

---

### 3. Frontend — Processing Page

#### [NEW] [ProcessingPage.jsx](file:///G:/FCCS/frontend/src/pages/ProcessingPage.jsx)

A new page at route `/processing` with:

- **Multi-file upload form**: 5 distinct drop zones (or a single zone with file-type labels), each for one of the required files:
  - ICM Report (`icm_report.xlsx`)
  - Parent Journal (`parent_journal.xlsx`)
  - Contribution Journal (`contribution_journal.xlsx`)
  - Plug Account Journal (`plugaccount_journal.xlsx`)
  - Report Inputs (`report_inputs.xlsx`)
- **Validation**: All 5 files must be uploaded before the "Process" button enables
- **Processing state**: Shows a spinner/progress indicator while the backend processes
- **Result display**: Shows success message with download button for the output file
- **Error handling**: Shows clear error messages if processing fails

#### [NEW] [ProcessingPage.css](file:///G:/FCCS/frontend/src/pages/ProcessingPage.css)

Styling for the processing page — dark theme matching existing design system with:
- Glassmorphism card for the upload area
- Animated file upload zones with drag-and-drop
- Processing spinner with progress animation
- Success/error state styling

#### [MODIFY] [App.jsx](file:///G:/FCCS/frontend/src/App.jsx)

Add route: `<Route path="/processing" element={<ProtectedRoute><ProcessingPage /></ProtectedRoute>} />`

#### [MODIFY] [Navbar.jsx](file:///G:/FCCS/frontend/src/components/Navbar.jsx)

Add "Processing" navigation link alongside existing nav items.

---

### 4. Output Column Attributes

Based on the user's screenshot, the output Excel file will contain these column groups per journal:

| Group | Columns |
|-------|---------|
| **Identity** | Entity, Partner |
| **Per Journal (×3)** | S1-Entity account cols, S2-Partner account cols, **Variance 1**, S1-Partner account cols, S2-Entity account cols, **Variance 2**, **Total** |
| **Grand Total** | Sum of all 3 journal Totals |

The existing `ic_matching_v3.py` already implements this layout. The refactored module will preserve this exact output format.

---

## Verification Plan

### Automated Tests

1. **Backend API test** — Run via `curl` or browser:
   ```
   # Start backend
   cd G:\FCCS\backend
   python -m uvicorn app.main:app --reload --port 8000

   # Test processing endpoint (requires auth token)
   curl -X POST http://localhost:8000/api/processing/run \
     -H "Authorization: Bearer <token>" \
     -F "name=test_run" \
     -F "type_name=alpha" \
     -F "icm_report=@icm_report.xlsx" \
     -F "parent_journal=@parent_journal.xlsx" \
     -F "contribution_journal=@contribution_journal.xlsx" \
     -F "plugaccount_journal=@plugaccount_journal.xlsx" \
     -F "report_inputs=@report_inputs.xlsx"
   ```

### Manual Verification

1. **Start both servers** and navigate to `http://localhost:5173/processing`
2. Upload all 5 Excel files using the UI
3. Click "Process" and wait for completion
4. Download the output file and verify:
   - Column structure matches the ICM report format
   - 3 journal blocks (Parent, Contribution, Plug Account)
   - Variance 1 & 2 calculated correctly per block
   - Grand Total = sum of all 3 journal Totals
5. Verify the output also appears in the Reports list page
