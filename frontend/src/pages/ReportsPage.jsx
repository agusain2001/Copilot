import { useState, useEffect, useRef, useCallback } from 'react';
import DatePicker from 'react-datepicker';
import 'react-datepicker/dist/react-datepicker.css';
import Navbar from '../components/Navbar';
import Toast from '../components/Toast';
import AddReportModal from '../components/AddReportModal';
import DeleteConfirmModal from '../components/DeleteConfirmModal';
import DownloadConfirmModal from '../components/DownloadConfirmModal';
import api from '../services/api';
import { FiPlus, FiSearch, FiFilter, FiTrash2, FiDownload, FiFile } from 'react-icons/fi';
import './ReportsPage.css';

export default function ReportsPage() {
  const [reports, setReports] = useState([]);
  const [selected, setSelected] = useState(new Set());
  const [search, setSearch] = useState('');
  const [selectedType, setSelectedType] = useState(null);
  const [selectedDate, setSelectedDate] = useState(null);
  const [showTypeDropdown, setShowTypeDropdown] = useState(false);
  const [showAdd, setShowAdd] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null);   // array of reports
  const [downloadTarget, setDownloadTarget] = useState(null); // array of reports
  const [toast, setToast] = useState(null);
  const [lastDeletedIds, setLastDeletedIds] = useState([]);
  const [loading, setLoading] = useState(true);
  const toastTimer = useRef(null);

  const fetchReports = useCallback(async () => {
    setLoading(true);
    try {
      const params = {};
      if (search) params.search = search;
      if (selectedType) params.type = selectedType;
      if (selectedDate) {
        // format YYYY-MM-DD
        const iso = selectedDate.toISOString().split('T')[0];
        params.date_from = iso;
        params.date_to = iso;
      }
      const res = await api.get('/api/reports', { params });
      setReports(res.data);
    } catch {}
    finally { setLoading(false); }
  }, [search, selectedType, selectedDate]);

  useEffect(() => { fetchReports(); }, [fetchReports]);

  // Click outside to close type dropdown
  const typeRef = useRef();
  useEffect(() => {
    const handleClickOutside = (e) => {
      if (typeRef.current && !typeRef.current.contains(e.target)) setShowTypeDropdown(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const showToast = (msg, undoIds = []) => {
    clearTimeout(toastTimer.current);
    setLastDeletedIds(undoIds);
    setToast({ msg, hasUndo: undoIds.length > 0 });
    toastTimer.current = setTimeout(() => setToast(null), 5000);
  };

  // ── Selection ──
  const toggleAll = (e) => {
    if (e.target.checked) setSelected(new Set(reports.map((r) => r.id)));
    else setSelected(new Set());
  };
  const toggleOne = (id) => {
    setSelected((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  };

  // ── Delete ──
  const confirmDelete = async () => {
    const ids = deleteTarget.map((r) => r.id);
    try {
      if (ids.length === 1) {
        await api.delete(`/api/reports/${ids[0]}`);
      } else {
        await api.post('/api/reports/bulk-delete', { report_ids: ids });
      }
      setLastDeletedIds(ids);
      setReports((prev) => prev.filter((r) => !ids.includes(r.id)));
      setSelected(new Set());
      const msg = ids.length === 1
        ? `Report: ${deleteTarget[0].name} Has been deleted.`
        : `${ids.length} Reports Has been deleted.`;
      showToast(msg, ids);
    } catch {}
    setDeleteTarget(null);
  };

  const handleUndo = async () => {
    try {
      if (lastDeletedIds.length === 1) {
        await api.post(`/api/reports/${lastDeletedIds[0]}/restore`);
      } else {
        await api.post('/api/reports/bulk-restore', { report_ids: lastDeletedIds });
      }
      fetchReports();
    } catch {}
    setToast(null);
  };

  // ── Download ──
  const confirmDownload = async () => {
    try {
      if (downloadTarget.length === 1) {
        const res = await api.get(`/api/reports/${downloadTarget[0].id}/download`, { responseType: 'blob' });
        const url = URL.createObjectURL(res.data);
        const a = document.createElement('a');
        a.href = url; a.download = downloadTarget[0].original_filename || 'report';
        a.click(); URL.revokeObjectURL(url);
      } else {
        const res = await api.post('/api/reports/bulk-export',
          { report_ids: downloadTarget.map((r) => r.id) },
          { responseType: 'blob' }
        );
        const url = URL.createObjectURL(res.data);
        const a = document.createElement('a');
        a.href = url; a.download = 'reports_export.zip';
        a.click(); URL.revokeObjectURL(url);
      }
    } catch {}
    setDownloadTarget(null);
  };

  const selectedReports = reports.filter((r) => selected.has(r.id));

  const formatTime = (ts) => {
    const d = new Date(ts);
    return d.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', hour12: false })
      + ` (${d.getDate()}${getOrdinal(d.getDate())} ${d.toLocaleString('en-IN', { month: 'long' })},${d.getFullYear()})`;
  };

  const getOrdinal = (n) => ['th','st','nd','rd'][n%10 > 3 ? 0 : (n%100-n%10 !== 10 ? n%10 : 0)] || 'th';

  return (
    <div className="reports-page">
      <Navbar />

      <div className="reports-content">
        {/* Header row */}
        <div className="reports-header">
          <h1 className="page-title">All Reports</h1>
          <button id="add-report-btn" className="btn-primary" onClick={() => setShowAdd(true)}>
            <FiPlus size={14} /> Add Report
          </button>
        </div>

        {/* Controls row */}
        <div className="controls-row">
          <div className="filter-pills">
            <div style={{ position: 'relative' }} ref={typeRef}>
              <button
                className={`pill-btn ${selectedType ? 'active' : ''}`}
                onClick={() => setShowTypeDropdown(!showTypeDropdown)}
              >
                <FiFilter size={11} /> {selectedType ? selectedType.charAt(0).toUpperCase() + selectedType.slice(1) : 'Sort By Type'}
              </button>
              {showTypeDropdown && (
                <div className="filter-dropdown menu-animate">
                  {['alpha', 'beta', 'gamma', 'theta'].map(t => (
                    <div
                      key={t}
                      className={`filter-dropdown-item ${selectedType === t ? 'active' : ''}`}
                      onClick={() => {
                        setSelectedType(selectedType === t ? null : t);
                        setShowTypeDropdown(false);
                      }}
                    >
                      {t.charAt(0).toUpperCase() + t.slice(1)}
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div style={{ position: 'relative' }}>
              <DatePicker
                selected={selectedDate}
                onChange={(date) => setSelectedDate(date)}
                customInput={
                  <button className={`pill-btn ${selectedDate ? 'active' : ''}`}>
                    <FiFilter size={11} /> {selectedDate ? selectedDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) : 'Sort By Date'}
                  </button>
                }
                popperPlacement="bottom-start"
                className="dark-calendar"
              />
            </div>
          </div>

          <div className="right-controls">
            {selected.size > 0 && (
              <>
                <button
                  className="icon-action-btn danger"
                  title="Delete selected"
                  onClick={() => setDeleteTarget(selectedReports)}
                >
                  <FiTrash2 size={14} />
                </button>
                <button
                  className="icon-action-btn"
                  title="Download selected"
                  onClick={() => setDownloadTarget(selectedReports)}
                >
                  <FiDownload size={14} />
                </button>
              </>
            )}
            <div className="search-wrap">
              <FiSearch size={13} className="search-icon" />
              <input
                className="search-input"
                placeholder="Search"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>
          </div>
        </div>

        {/* Table */}
        <div className="table-wrap">
          {loading ? (
            <div className="empty-state"><p>Loading...</p></div>
          ) : reports.length === 0 ? (
            <div className="empty-state">
              <div style={{ fontSize: 28, marginBottom: 10, opacity: 0.3 }}>📋</div>
              <p style={{ fontWeight: 600, marginBottom: 4 }}>No Reports Added Yet</p>
              <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginBottom: 16 }}>
                No reports found. Click on "Add Report" to upload an Excel file and begin organizing your data.
              </p>
              <button className="btn-primary" onClick={() => setShowAdd(true)}>
                <FiPlus size={13} /> Add Report
              </button>
            </div>
          ) : (
            <table className="data-table">
              <thead>
                <tr>
                  <th style={{ width: 36 }}>
                    <input
                      type="checkbox"
                      checked={selected.size === reports.length && reports.length > 0}
                      onChange={toggleAll}
                    />
                  </th>
                  <th>Report Name</th>
                  <th>Type</th>
                  <th>Uploaded Time</th>
                  <th>File</th>
                </tr>
              </thead>
              <tbody>
                {reports.map((r) => (
                  <tr key={r.id} className={selected.has(r.id) ? 'selected-row' : ''}>
                    <td>
                      <input
                        type="checkbox"
                        checked={selected.has(r.id)}
                        onChange={() => toggleOne(r.id)}
                      />
                    </td>
                    <td>{r.name}</td>
                    <td style={{ color: 'var(--accent-cyan)' }}>{r.type_label}</td>
                    <td style={{ color: 'var(--text-secondary)' }}>{formatTime(r.uploaded_at)}</td>
                    <td>
                      <button
                        className="file-pill"
                        onClick={() => setDownloadTarget([r])}
                        title="Click to download"
                      >
                        <FiFile size={12} />
                        <span>{r.original_filename}</span>
                        <span className="file-pill-hint">click to view</span>
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {/* Modals */}
      {showAdd && (
        <AddReportModal
          onClose={() => setShowAdd(false)}
          onSuccess={(msg) => { fetchReports(); showToast(msg); }}
        />
      )}
      {deleteTarget && (
        <DeleteConfirmModal
          reports={deleteTarget}
          onConfirm={confirmDelete}
          onClose={() => setDeleteTarget(null)}
        />
      )}
      {downloadTarget && (
        <DownloadConfirmModal
          reports={downloadTarget}
          onConfirm={confirmDownload}
          onClose={() => setDownloadTarget(null)}
        />
      )}

      {/* Toast */}
      {toast && (
        <Toast
          message={toast.msg}
          onUndo={toast.hasUndo ? handleUndo : undefined}
          onClose={() => setToast(null)}
        />
      )}
    </div>
  );
}
