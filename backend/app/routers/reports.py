import uuid
import zipfile
import io
from pathlib import Path
from datetime import datetime, timezone
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form, Query, status
from fastapi.responses import FileResponse, StreamingResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, and_
from app.database import get_db
from app.models.report import Report
from app.models.report_type import ReportType
from app.models.report_sequence import ReportSequence
from app.schemas.report import ReportResponse, BulkActionRequest
from app.utils.files import save_report_file, delete_file, ALLOWED_REPORT_EXTENSIONS
from app.dependencies import get_current_user
from app.models.user import User

router = APIRouter(prefix="/api/reports", tags=["Reports"])


def report_to_response(r: Report) -> ReportResponse:
    return ReportResponse(
        id=r.id,
        name=r.name,
        type_label=r.report_type.label if r.report_type else "",
        type_name=r.report_type.name if r.report_type else "",
        original_filename=r.original_filename,
        file_size_bytes=r.file_size_bytes,
        uploaded_at=r.uploaded_at,
        uploaded_by_name=r.uploader.full_name if r.uploader else "",
        is_deleted=r.is_deleted,
        # Structured-storage fields
        sequence_id=r.sequence_id,
        file_type=r.file_type,
        stored_filename=r.stored_filename,
        folder_path=r.folder_path,
    )


@router.get("", response_model=list[ReportResponse])
async def list_reports(
    type_name: str | None = Query(None),
    date_from: datetime | None = Query(None),
    date_to: datetime | None = Query(None),
    search: str | None = Query(None),
    sort_by: str = Query("date_desc"),
    sequence_id: int | None = Query(None),
    file_type: str | None = Query(None),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    filters = [Report.is_deleted == False]

    if type_name:
        result = await db.execute(select(ReportType).where(ReportType.name == type_name))
        rt = result.scalar_one_or_none()
        if rt:
            filters.append(Report.type_id == rt.id)

    if date_from:
        filters.append(Report.uploaded_at >= date_from)
    if date_to:
        filters.append(Report.uploaded_at <= date_to)
    if search:
        filters.append(Report.name.ilike(f"%{search}%"))
    if sequence_id is not None:
        filters.append(Report.sequence_id == sequence_id)
    if file_type:
        filters.append(Report.file_type == file_type)

    query = select(Report).where(and_(*filters))

    if sort_by == "date_asc":
        query = query.order_by(Report.uploaded_at.asc())
    elif sort_by == "type":
        query = query.order_by(Report.type_id)
    else:
        query = query.order_by(Report.uploaded_at.desc())

    result = await db.execute(query)
    reports = result.scalars().all()
    return [report_to_response(r) for r in reports]


@router.post("", response_model=ReportResponse, status_code=status.HTTP_201_CREATED)
async def create_report(
    name: str = Form(...),
    type_name: str = Form(...),
    file: UploadFile = File(...),
    file_type: str = Form("input", description="Must be 'input' or 'output'"),
    sequence_id: int | None = Form(
        None,
        description="Existing sequence/group ID. Leave empty to create a new group folder.",
    ),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """
    Upload a report file.

    - **file_type**: `input` or `output`
    - **sequence_id**: optional integer. If omitted, a new group folder (N+1) is created.
      Pass the same `sequence_id` to add an output file to the same folder as its input.

    The file is saved to:
        `uploads/reports/<sequence_id>/inputs/<ddmmyyyyhhmmss>_<filename>`
        `uploads/reports/<sequence_id>/outputs/<ddmmyyyyhhmmss>_<filename>`
    """
    # Validate file_type
    if file_type not in ("input", "output"):
        raise HTTPException(status_code=400, detail="file_type must be 'input' or 'output'")

    # Validate extension
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_REPORT_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Only .csv, .xlsx, .xls files are allowed")

    # Validate report type
    rt_result = await db.execute(
        select(ReportType).where(ReportType.name == type_name, ReportType.is_active == True)
    )
    report_type = rt_result.scalar_one_or_none()
    if not report_type:
        raise HTTPException(status_code=400, detail=f"Report type '{type_name}' not found")

    # Resolve or create sequence
    if sequence_id is not None:
        # Verify the sequence exists
        seq_result = await db.execute(
            select(ReportSequence).where(ReportSequence.id == sequence_id)
        )
        seq = seq_result.scalar_one_or_none()
        if not seq:
            raise HTTPException(
                status_code=404,
                detail=f"Sequence group {sequence_id} not found. "
                       "Leave sequence_id empty to create a new group.",
            )
    else:
        # Create a new sequence entry — DB auto-assigns the next integer ID
        seq = ReportSequence()
        db.add(seq)
        await db.flush()   # flush to get the generated ID without full commit
        sequence_id = seq.id

    # Save the file with the structured path
    full_path, stored_filename, folder_path, file_size = await save_report_file(
        file, sequence_id=sequence_id, file_type=file_type
    )

    # Persist report record
    report = Report(
        name=name,
        type_id=report_type.id,
        file_url=full_path,
        original_filename=file.filename,
        file_size_bytes=file_size,
        uploaded_by=current_user.id,
        sequence_id=sequence_id,
        file_type=file_type,
        stored_filename=stored_filename,
        folder_path=folder_path,
    )
    db.add(report)
    await db.commit()
    await db.refresh(report)

    # Reload with relationships
    result = await db.execute(select(Report).where(Report.id == report.id))
    report = result.scalar_one()
    return report_to_response(report)


@router.get("/{report_id}", response_model=ReportResponse)
async def get_report(
    report_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(select(Report).where(Report.id == report_id, Report.is_deleted == False))
    report = result.scalar_one_or_none()
    if not report:
        raise HTTPException(status_code=404, detail="Report not found")
    return report_to_response(report)


@router.delete("/{report_id}")
async def delete_report(
    report_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(select(Report).where(Report.id == report_id, Report.is_deleted == False))
    report = result.scalar_one_or_none()
    if not report:
        raise HTTPException(status_code=404, detail="Report not found")

    report.is_deleted = True
    report.deleted_at = datetime.now(timezone.utc)
    await db.commit()
    return {"message": "Report deleted", "report_id": str(report_id)}


@router.post("/{report_id}/restore")
async def restore_report(
    report_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(select(Report).where(Report.id == report_id, Report.is_deleted == True))
    report = result.scalar_one_or_none()
    if not report:
        raise HTTPException(status_code=404, detail="Deleted report not found")

    report.is_deleted = False
    report.deleted_at = None
    await db.commit()
    return {"message": "Report restored", "report_id": str(report_id)}


@router.post("/bulk-delete")
async def bulk_delete(
    payload: BulkActionRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Report).where(Report.id.in_(payload.report_ids), Report.is_deleted == False)
    )
    reports = result.scalars().all()

    now = datetime.now(timezone.utc)
    for r in reports:
        r.is_deleted = True
        r.deleted_at = now

    await db.commit()
    return {"message": f"{len(reports)} report(s) deleted", "count": len(reports)}


@router.post("/bulk-restore")
async def bulk_restore(
    payload: BulkActionRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Report).where(Report.id.in_(payload.report_ids), Report.is_deleted == True)
    )
    reports = result.scalars().all()

    for r in reports:
        r.is_deleted = False
        r.deleted_at = None

    await db.commit()
    return {"message": f"{len(reports)} report(s) restored", "count": len(reports)}


@router.post("/bulk-export")
async def bulk_export(
    payload: BulkActionRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Report).where(Report.id.in_(payload.report_ids), Report.is_deleted == False)
    )
    reports = result.scalars().all()

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in reports:
            path = Path(r.file_url)
            if path.exists():
                zf.write(path, arcname=r.original_filename)

    zip_buffer.seek(0)
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=reports_export.zip"},
    )


@router.get("/{report_id}/download")
async def download_report(
    report_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(select(Report).where(Report.id == report_id, Report.is_deleted == False))
    report = result.scalar_one_or_none()
    if not report:
        raise HTTPException(status_code=404, detail="Report not found")

    path = Path(report.file_url)
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found on disk")

    return FileResponse(
        path,
        filename=report.original_filename,
        media_type="application/octet-stream",
    )
