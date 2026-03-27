import { useState, useRef, useCallback } from 'react';
import api from '../services/api';
import {
  FiUploadCloud,
  FiFileText,
  FiX,
  FiPlay,
  FiDownload,
  FiRefreshCw,
  FiCheckCircle,
  FiAlertCircle,
} from 'react-icons/fi';

const FILE_SLOTS = [
  { key: 'icm_report',           label: 'ICM Report',           hint: 'Intercompany Balances IC Matching Report', required: true  },
  { key: 'parent_journal',       label: 'Parent Journal',        hint: 'Journal Report 1 — Parent Input Data',    required: true  },
  { key: 'contribution_journal', label: 'Contribution Journal',  hint: 'Journal Report 2 — Contribution Data',    required: true  },
  { key: 'plugaccount_journal',  label: 'Plug Account Journal',  hint: 'Journal Report 4 — Plug Account Data',    required: true  },
  { key: 'report_inputs',        label: 'Report Inputs',         hint: 'Optional — archived for records only',    required: false },
];

export default function AddReportModal({ onClose, onSuccess }) {
  const [files, setFiles]       = useState({});
  const [runName, setRunName]   = useState('');
  const [processing, setProcessing] = useState(false);
  const [result, setResult]     = useState(null);
  const [dragKey, setDragKey]   = useState(null);
  const inputRefs = useRef({});

  // ── File handlers ────────────────────────────────────────────────
  const handleFile = useCallback((key, file) => {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'xlsx') {
      alert(`Only .xlsx files are allowed. "${file.name}" is .${ext}`);
      return;
    }
    setFiles((prev) => ({ ...prev, [key]: file }));
  }, []);

  const removeFile = useCallback((key) => {
    setFiles((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  }, []);

  const handleDrop = useCallback(
    (key) => (e) => {
      e.preventDefault();
      e.stopPropagation();
      handleFile(key, e.dataTransfer.files[0]);
    },
    [handleFile]
  );

  const handleDragOver  = (e) => { e.preventDefault(); e.stopPropagation(); };
  const handleDragEnter = (key) => (e) => { e.preventDefault(); setDragKey(key); };
  const handleDragLeave = (key) => (e) => { e.preventDefault(); if (dragKey === key) setDragKey(null); };

  // ── Submit ───────────────────────────────────────────────────────
  const requiredSlots     = FILE_SLOTS.filter((s) => s.required);
  const allRequiredFilled = requiredSlots.every((s) => files[s.key]);
  const filledCount       = Object.keys(files).length;

  const handleProcess = async () => {
    if (!allRequiredFilled) return;
    setProcessing(true);
    setResult(null);

    const formData = new FormData();
    formData.append('name', runName || `Processing Run ${new Date().toLocaleString()}`);
    formData.append('type_name', 'alpha');
    FILE_SLOTS.forEach((s) => {
      if (files[s.key]) formData.append(s.key, files[s.key]);
    });

    try {
      const res = await api.post('/api/processing/run', formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 300000,
      });
      setResult({
        status: 'success',
        sequenceId:     res.data.sequence_id,
        outputReportId: res.data.output_report_id,
        outputFilename: res.data.output_filename,
      });
      onSuccess?.('Processing complete — output report saved.');
    } catch (err) {
      setResult({
        status: 'error',
        error: err.response?.data?.detail || err.message || 'Processing failed',
      });
    } finally {
      setProcessing(false);
    }
  };

  // ── Download ─────────────────────────────────────────────────────
  const handleDownload = async () => {
    if (!result?.sequenceId) return;
    try {
      const res = await api.get(`/api/processing/${result.sequenceId}/output`, {
        responseType: 'blob',
      });
      const url = URL.createObjectURL(res.data);
      const a   = document.createElement('a');
      a.href     = url;
      a.download = result.outputFilename || 'ICM_Output.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      alert('Download failed: ' + (err.response?.data?.detail || err.message));
    }
  };

  // ── Reset ────────────────────────────────────────────────────────
  const handleReset = () => {
    setFiles({});
    setRunName('');
    setResult(null);
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="modal-box icm-modal-box"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="icm-modal-header">
          <div>
            <h2 className="icm-modal-title">ICM Report Processing</h2>
            <p className="icm-modal-subtitle">
              Upload the 4 required Excel files (and optionally a Report Inputs file) to run intercompany matching.
            </p>
          </div>
          <button className="modal-close-btn" onClick={onClose}><FiX /></button>
        </div>

        {/* Upload Card */}
        <div className="icm-upload-card">
          <div className="icm-card-title">Upload Files</div>
          <div className="icm-card-subtitle">
            Drag &amp; drop or click each zone to select the required .xlsx files.
          </div>

          {/* Run name */}
          <div className="icm-name-row">
            <label htmlFor="icm-run-name">Run Name (optional)</label>
            <input
              id="icm-run-name"
              type="text"
              className="input-field"
              placeholder="e.g. March 2026 IC Run"
              value={runName}
              onChange={(e) => setRunName(e.target.value)}
              disabled={processing}
            />
          </div>

          {/* File drop zones */}
          <div className="icm-file-grid">
            {FILE_SLOTS.map((slot) => {
              const file      = files[slot.key];
              const isDragOver = dragKey === slot.key;
              return (
                <div
                  key={slot.key}
                  className={`icm-drop-zone${isDragOver ? ' drag-over' : ''}${file ? ' has-file' : ''}`}
                  onClick={() => inputRefs.current[slot.key]?.click()}
                  onDrop={handleDrop(slot.key)}
                  onDragOver={handleDragOver}
                  onDragEnter={handleDragEnter(slot.key)}
                  onDragLeave={handleDragLeave(slot.key)}
                >
                  <input
                    type="file"
                    accept=".xlsx"
                    ref={(el) => (inputRefs.current[slot.key] = el)}
                    onChange={(e) => { handleFile(slot.key, e.target.files[0]); e.target.value = ''; }}
                  />
                  <div className="icm-dz-icon">
                    {file ? <FiFileText /> : <FiUploadCloud />}
                  </div>
                  <div className="icm-dz-label">{slot.label}</div>
                  {file ? (
                    <div className="icm-dz-file">
                      {file.name}
                      <button
                        className="icm-remove-btn"
                        onClick={(e) => { e.stopPropagation(); removeFile(slot.key); }}
                        title="Remove file"
                      >
                        <FiX />
                      </button>
                    </div>
                  ) : (
                    <div className="icm-dz-hint">{slot.hint}</div>
                  )}
                </div>
              );
            })}
          </div>

          {/* Actions */}
          <div className="icm-actions">
            <span className="icm-file-counter">
              <strong>{filledCount}</strong> / {FILE_SLOTS.length} files
              {' — '}
              {allRequiredFilled
                ? <span style={{ color: 'var(--accent-success)', fontWeight: 600 }}>✓ Ready to process</span>
                : <span>{requiredSlots.filter((s) => files[s.key]).length} / {requiredSlots.length} required</span>
              }
            </span>
            <button
              className="icm-process-btn"
              disabled={!allRequiredFilled || processing}
              onClick={handleProcess}
            >
              {processing ? (
                <>
                  <span className="icm-btn-spinner" />
                  Processing…
                </>
              ) : (
                <>
                  <FiPlay size={14} />
                  Process Files
                </>
              )}
            </button>
          </div>
        </div>

        {/* Processing spinner overlay */}
        {processing && (
          <div className="icm-processing-state">
            <div className="icm-spinner-wrap">
              <div className="icm-spinner" />
              <div className="icm-spinner-text">Running IC Matching…</div>
              <div className="icm-spinner-sub">This may take a minute depending on file size.</div>
            </div>
          </div>
        )}

        {/* Result */}
        {result && !processing && (
          <div className={`icm-result-card ${result.status}`}>
            <div className="icm-result-header">
              <div className={`icm-result-icon ${result.status}`}>
                {result.status === 'success' ? <FiCheckCircle /> : <FiAlertCircle />}
              </div>
              <div className="icm-result-title">
                {result.status === 'success' ? 'Processing Complete' : 'Processing Failed'}
              </div>
            </div>
            <div className="icm-result-details">
              {result.status === 'success' ? (
                <>
                  Output file <strong>{result.outputFilename}</strong> is ready for download.
                  <br />
                  Sequence ID: {result.sequenceId}
                </>
              ) : (
                <>Error: {result.error}</>
              )}
            </div>
            <div className="icm-result-actions">
              {result.status === 'success' && (
                <button className="icm-download-btn" onClick={handleDownload}>
                  <FiDownload size={14} /> Download Output
                </button>
              )}
              <button className="icm-reset-btn" onClick={handleReset}>
                <FiRefreshCw size={14} style={{ marginRight: 6 }} /> New Run
              </button>
              <button className="icm-reset-btn" onClick={onClose}>
                Close
              </button>
            </div>
          </div>
        )}
      </div>

      <style>{`
        /* ── Modal shell ── */
        .icm-modal-box {
          max-width: 780px;
          width: 95vw;
          max-height: 92vh;
          overflow-y: auto;
          padding: 32px;
        }

        /* ── Header ── */
        .icm-modal-header {
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 16px;
          margin-bottom: 24px;
        }
        .icm-modal-title {
          font-size: 22px;
          font-weight: 700;
          color: var(--text-primary);
          margin-bottom: 6px;
        }
        .icm-modal-subtitle {
          font-size: 13px;
          color: var(--accent-cyan);
        }

        /* ── Upload card ── */
        .icm-upload-card {
          background: var(--bg-card);
          border: 1px solid var(--border-color);
          border-radius: var(--radius-lg);
          padding: 24px;
          box-shadow: var(--shadow-sm);
        }
        .icm-card-title {
          font-size: 15px;
          font-weight: 600;
          color: var(--text-primary);
          margin-bottom: 6px;
        }
        .icm-card-subtitle {
          font-size: 12px;
          color: var(--text-secondary);
          margin-bottom: 20px;
        }

        /* ── Name row ── */
        .icm-name-row {
          display: flex;
          flex-direction: column;
          gap: 6px;
          margin-bottom: 20px;
        }
        .icm-name-row label {
          font-size: 12px;
          color: var(--text-secondary);
          font-weight: 500;
        }

        /* ── File grid ── */
        .icm-file-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 14px;
          margin-bottom: 24px;
        }
        .icm-file-grid .icm-drop-zone:nth-child(5) {
          grid-column: 1 / -1;
        }
        @media (max-width: 560px) {
          .icm-file-grid { grid-template-columns: 1fr; }
        }

        /* ── Drop zone ── */
        .icm-drop-zone {
          position: relative;
          border: 2px dashed var(--border-color);
          border-radius: var(--radius-md);
          padding: 22px 16px;
          text-align: center;
          cursor: pointer;
          transition: border-color 0.2s, background 0.2s, transform 0.15s;
          background: var(--bg-secondary);
          min-height: 120px;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          gap: 7px;
        }
        .icm-drop-zone input[type="file"] { display: none; }
        .icm-drop-zone:hover {
          border-color: var(--accent-cyan);
          background: rgba(8,145,178,0.04);
        }
        .icm-drop-zone.drag-over {
          border-color: var(--accent-cyan-light);
          background: rgba(8,145,178,0.08);
          transform: scale(1.01);
        }
        .icm-drop-zone.has-file {
          border-color: var(--accent-success);
          border-style: solid;
          background: rgba(22,163,74,0.04);
        }
        .icm-dz-icon {
          font-size: 24px;
          color: var(--text-muted);
          transition: color 0.2s;
        }
        .icm-drop-zone:hover .icm-dz-icon,
        .icm-drop-zone.drag-over .icm-dz-icon { color: var(--accent-cyan); }
        .icm-drop-zone.has-file .icm-dz-icon { color: var(--accent-success); }
        .icm-dz-label {
          font-size: 13px;
          font-weight: 600;
          color: var(--text-primary);
        }
        .icm-dz-hint {
          font-size: 11px;
          color: var(--text-muted);
        }
        .icm-dz-file {
          font-size: 11px;
          color: var(--accent-success);
          font-weight: 500;
          display: flex;
          align-items: center;
          gap: 5px;
          word-break: break-all;
        }
        .icm-remove-btn {
          background: none;
          border: none;
          color: var(--accent-danger);
          cursor: pointer;
          font-size: 13px;
          padding: 2px;
          display: inline-flex;
          align-items: center;
          transition: opacity 0.15s;
          flex-shrink: 0;
        }
        .icm-remove-btn:hover { opacity: 0.7; }

        /* ── Actions row ── */
        .icm-actions {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
          flex-wrap: wrap;
        }
        .icm-file-counter {
          font-size: 12px;
          color: var(--text-secondary);
        }
        .icm-file-counter strong { color: var(--text-primary); }

        /* ── Process button ── */
        .icm-process-btn {
          background: linear-gradient(135deg, #0891b2, #0e7490);
          color: #fff;
          border: none;
          padding: 10px 24px;
          border-radius: var(--radius-sm);
          font-size: 13px;
          font-weight: 600;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          gap: 8px;
          transition: opacity 0.2s, transform 0.15s;
        }
        .icm-process-btn:hover:not(:disabled) {
          opacity: 0.9;
          transform: translateY(-1px);
        }
        .icm-process-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
          transform: none;
        }

        /* ── Inline spinner inside button ── */
        .icm-btn-spinner {
          width: 14px;
          height: 14px;
          border: 2px solid rgba(255,255,255,0.35);
          border-top-color: #fff;
          border-radius: 50%;
          animation: icm-spin 0.7s linear infinite;
          display: inline-block;
        }
        @keyframes icm-spin { to { transform: rotate(360deg); } }

        /* ── Processing state ── */
        .icm-processing-state {
          margin-top: 20px;
          padding: 24px;
          border-radius: var(--radius-md);
          border: 1px solid var(--border-color);
          background: var(--bg-secondary);
          text-align: center;
          animation: icm-slide-up 0.3s ease;
        }
        .icm-spinner-wrap {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 12px;
        }
        .icm-spinner {
          width: 38px;
          height: 38px;
          border: 3px solid var(--border-color);
          border-top-color: var(--accent-cyan);
          border-radius: 50%;
          animation: icm-spin 0.8s linear infinite;
        }
        .icm-spinner-text {
          font-size: 14px;
          color: var(--text-secondary);
          font-weight: 500;
        }
        .icm-spinner-sub {
          font-size: 12px;
          color: var(--text-muted);
        }

        /* ── Result card ── */
        .icm-result-card {
          margin-top: 20px;
          padding: 20px 24px;
          border-radius: var(--radius-md);
          border: 1px solid var(--border-color);
          background: var(--bg-secondary);
          animation: icm-slide-up 0.3s ease;
        }
        .icm-result-card.success { border-color: var(--accent-success); }
        .icm-result-card.error   { border-color: var(--accent-danger); }
        .icm-result-header {
          display: flex;
          align-items: center;
          gap: 10px;
          margin-bottom: 10px;
        }
        .icm-result-icon { font-size: 26px; }
        .icm-result-icon.success { color: var(--accent-success); }
        .icm-result-icon.error   { color: var(--accent-danger); }
        .icm-result-title {
          font-size: 15px;
          font-weight: 600;
          color: var(--text-primary);
        }
        .icm-result-details {
          font-size: 13px;
          color: var(--text-secondary);
          margin-bottom: 14px;
          line-height: 1.6;
        }
        .icm-result-actions { display: flex; gap: 10px; flex-wrap: wrap; }

        .icm-download-btn {
          background: linear-gradient(135deg, #16a34a, #15803d);
          color: #fff;
          border: none;
          padding: 9px 20px;
          border-radius: var(--radius-sm);
          font-size: 13px;
          font-weight: 600;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          gap: 7px;
          transition: opacity 0.2s;
        }
        .icm-download-btn:hover { opacity: 0.9; }

        .icm-reset-btn {
          background: transparent;
          color: var(--text-secondary);
          border: 1px solid var(--border-color);
          padding: 9px 20px;
          border-radius: var(--radius-sm);
          font-size: 13px;
          cursor: pointer;
          transition: all 0.15s;
          display: inline-flex;
          align-items: center;
        }
        .icm-reset-btn:hover {
          border-color: var(--text-secondary);
          color: var(--text-primary);
        }

        @keyframes icm-slide-up {
          from { opacity: 0; transform: translateY(10px); }
          to   { opacity: 1; transform: translateY(0); }
        }
      `}</style>
    </div>
  );
}
