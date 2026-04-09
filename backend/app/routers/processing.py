"""
Processing router for IC matching.
"""

import os
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.database import get_db
from app.dependencies import get_current_user
from app.models.user import User
from app.models.report import Report
from app.models.report_type import ReportType
from app.models.report_sequence import ReportSequence
from app.config import settings

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/processing", tags=["Processing"])


async def _save_temp_upload(file: UploadFile, dest_dir: str) -> tuple[str, int]:
    os.makedirs(dest_dir, exist_ok=True)
    file_path = os.path.join(dest_dir, file.filename)
    content = await file.read()
    file_size = len(content)

    max_bytes = settings.MAX_FILE_SIZE_MB * 1024 * 1024
    if file_size > max_bytes:
        raise HTTPException(
            status_code=413,
            detail=f"File '{file.filename}' too large. Max {settings.MAX_FILE_SIZE_MB}MB.",
        )

    with open(file_path, "wb") as handle:
        handle.write(content)
    return file_path, file_size


def _validate_extension(file: UploadFile, label: str, allowed: set[str]):
    ext = Path(file.filename).suffix.lower()
    if ext not in allowed:
        allowed_text = ", ".join(sorted(allowed))
        raise HTTPException(
            status_code=400,
            detail=f"{label} must be one of [{allowed_text}] (got '{ext}')",
        )


def _validate_xlsx(file: UploadFile, label: str):
    _validate_extension(file, label, {".xlsx"})


def _validate_xlsx_or_csv(file: UploadFile, label: str):
    _validate_extension(file, label, {".xlsx", ".csv"})


@router.post("/run")
async def run_processing(
    name: str = Form(..., description="Name for this processing run"),
    type_name: str = Form("alpha", description="Report type name"),
    icm_report: UploadFile = File(...),
    parent_journal: UploadFile = File(...),
    contribution_journal: Optional[UploadFile] = File(None),
    plugaccount_journal: UploadFile = File(...),
    entity_with_currency: UploadFile = File(...),
    exchange_rates: UploadFile = File(...),
    ownership_structure: UploadFile = File(...),
    report_inputs: Optional[UploadFile] = File(None),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """
    Required uploads:
    - icm_report (.xlsx)
    - parent_journal (.xlsx)
    - plugaccount_journal (.xlsx)
    - entity_with_currency (.csv/.xlsx)
    - exchange_rates (.xlsx)
    - ownership_structure (.xlsx)

    Optional uploads:
    - contribution_journal (.xlsx) (deprecated and ignored by QAR flow)
    - report_inputs (.xlsx)
    """
    _validate_xlsx(icm_report, "ICM Report")
    _validate_xlsx(parent_journal, "Parent Journal")
    _validate_xlsx(plugaccount_journal, "Plug Account Journal")
    _validate_xlsx_or_csv(entity_with_currency, "Entity With Currency")
    _validate_xlsx(exchange_rates, "Exchange Rates")
    _validate_xlsx(ownership_structure, "Ownership Structure")
    if contribution_journal:
        _validate_xlsx(contribution_journal, "Contribution Journal")
    if report_inputs:
        _validate_xlsx(report_inputs, "Report Inputs")

    rt_result = await db.execute(
        select(ReportType).where(ReportType.name == type_name, ReportType.is_active == True)
    )
    report_type = rt_result.scalar_one_or_none()
    if not report_type:
        fallback_result = await db.execute(
            select(ReportType).where(ReportType.is_active == True).order_by(ReportType.sort_order)
        )
        report_type = fallback_result.scalars().first()
    if not report_type:
        raise HTTPException(
            status_code=500,
            detail="No active report type found in the database. Please seed report types.",
        )

    seq = ReportSequence()
    db.add(seq)
    await db.flush()
    sequence_id = seq.id

    input_dir = os.path.join(settings.UPLOAD_DIR, "reports", str(sequence_id), "inputs")
    files_map = {
        "icm_report": icm_report,
        "parent_journal": parent_journal,
        "plugaccount_journal": plugaccount_journal,
        "entity_with_currency": entity_with_currency,
        "exchange_rates": exchange_rates,
        "ownership_structure": ownership_structure,
    }
    if contribution_journal:
        files_map["contribution_journal"] = contribution_journal
    if report_inputs:
        files_map["report_inputs"] = report_inputs

    saved_paths = {}
    input_reports = []
    for file_key, upload_file in files_map.items():
        path, size = await _save_temp_upload(upload_file, input_dir)
        saved_paths[file_key] = path

        now_ts = datetime.now().strftime("%d%m%Y%H%M%S")
        stored_name = f"{now_ts}_{upload_file.filename}"
        folder_path = str(Path("uploads") / "reports" / str(sequence_id) / "inputs" / upload_file.filename)
        report = Report(
            name=f"{name} - {file_key}",
            type_id=report_type.id,
            file_url=path,
            original_filename=upload_file.filename,
            file_size_bytes=size,
            uploaded_by=current_user.id,
            sequence_id=sequence_id,
            file_type="input",
            stored_filename=stored_name,
            folder_path=folder_path,
        )
        db.add(report)
        input_reports.append(report)

    from app.ic_processor import process_icm_report

    output_dir = os.path.join(settings.UPLOAD_DIR, "reports", str(sequence_id), "outputs")
    os.makedirs(output_dir, exist_ok=True)
    output_filename = f"ICM_Output_{sequence_id}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    journal_paths = {
        "parent_journal": saved_paths["parent_journal"],
        "plugaccount_journal": saved_paths["plugaccount_journal"],
    }
    if saved_paths.get("contribution_journal"):
        journal_paths["contribution_journal"] = saved_paths["contribution_journal"]

    lookup_paths = {
        "entity_with_currency": saved_paths["entity_with_currency"],
        "exchange_rates": saved_paths["exchange_rates"],
        "ownership_structure": saved_paths["ownership_structure"],
    }

    try:
        process_icm_report(
            icm_path=saved_paths["icm_report"],
            journal_paths=journal_paths,
            output_path=output_path,
            report_inputs_path=saved_paths.get("report_inputs"),
            lookup_paths=lookup_paths,
        )
    except Exception as exc:
        logger.exception("IC processing failed for sequence %d", sequence_id)
        await db.rollback()
        raise HTTPException(
            status_code=500,
            detail=f"Processing failed: {str(exc)}",
        )

    output_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
    now_ts = datetime.now().strftime("%d%m%Y%H%M%S")
    output_stored_name = f"{now_ts}_{output_filename}"
    output_folder_path = str(Path("uploads") / "reports" / str(sequence_id) / "outputs" / output_filename)
    output_report = Report(
        name=f"{name} - output",
        type_id=report_type.id,
        file_url=output_path,
        original_filename=output_filename,
        file_size_bytes=output_size,
        uploaded_by=current_user.id,
        sequence_id=sequence_id,
        file_type="output",
        stored_filename=output_stored_name,
        folder_path=output_folder_path,
    )
    db.add(output_report)

    await db.commit()
    await db.refresh(output_report)

    return {
        "status": "success",
        "sequence_id": sequence_id,
        "output_report_id": str(output_report.id),
        "output_filename": output_filename,
        "download_url": f"/api/processing/{sequence_id}/output",
        "input_files_count": len(input_reports),
    }


@router.get("/{sequence_id}/output")
async def download_output(
    sequence_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Report).where(
            Report.sequence_id == sequence_id,
            Report.file_type == "output",
            Report.is_deleted == False,
        )
    )
    report = result.scalar_one_or_none()
    if not report:
        raise HTTPException(status_code=404, detail="Output file not found for this sequence")

    path = Path(report.file_url)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Output file not found on disk")

    return FileResponse(
        path,
        filename=report.original_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
