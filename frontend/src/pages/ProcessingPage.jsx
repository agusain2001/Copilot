import { useState, useRef, useCallback } from 'react';
import Navbar from '../components/Navbar';
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
import './ProcessingPage.css';

const FILE_SLOTS = [
  { key: 'icm_report',            label: 'ICM Report',             hint: 'Intercompany Balances IC Matching Report', required: true  },
  { key: 'parent_journal',        label: 'Parent Journal',         hint: 'Journal Report 1 — Parent Input Data',    required: true  },
  { key: 'contribution_journal',  label: 'Contribution Journal',   hint: 'Journal Report 2 — Contribution Data',    required: true  },
  { key: 'plugaccount_journal',   label: 'Plug Account Journal',   hint: 'Journal Report 4 — Plug Account Data',    required: true  },
  { key: 'report_inputs',         label: 'Report Inputs',          hint: 'Optional — archived for records only',     required: false },
];

export default function ProcessingPage() {
  const [files, setFiles] = useState({});
  const [runName, setRunName] = useState('');
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState(null);   // { status, sequenceId, outputReportId, outputFilename, error }
  const inputRefs = useRef({});

  // ── File handlers ───────────────────────────────────────────────
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
      const file = e.dataTransfer.files[0];
      handleFile(key, file);
    },
    [handleFile]
  );

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  // ── Drag state for visual feedback ──────────────────────────────
  const [dragKey, setDragKey] = useState(null);

  const handleDragEnter = (key) => (e) => {
    e.preventDefault();
    setDragKey(key);
  };
  const handleDragLeave = (key) => (e) => {
    e.preventDefault();
    if (dragKey === key) setDragKey(null);
  };

  // ── Submit ──────────────────────────────────────────────────────
  const requiredSlots = FILE_SLOTS.filter((s) => s.required);
  const allRequiredFilled = requiredSlots.every((s) => files[s.key]);

  const handleProcess = async () => {
    if (!allRequiredFilled) return;
    setProcessing(true);
    setResult(null);

    const formData = new FormData();
    formData.append('name', runName || `Processing Run ${new Date().toLocaleString()}`);
    formData.append('type_name', 'alpha');
    FILE_SLOTS.forEach((s) => {
      if (files[s.key]) {
        formData.append(s.key, files[s.key]);
      }
    });

    try {
      const res = await api.post('/api/processing/run', formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
        timeout: 300000,  // 5 min timeout for large files
      });
      setResult({
        status: 'success',
        sequenceId: res.data.sequence_id,
        outputReportId: res.data.output_report_id,
        outputFilename: res.data.output_filename,
      });
    } catch (err) {
      setResult({
        status: 'error',
        error: err.response?.data?.detail || err.message || 'Processing failed',
      });
    } finally {
      setProcessing(false);
    }
  };

  // ── Download ────────────────────────────────────────────────────
  const handleDownload = async () => {
    if (!result?.sequenceId) return;
    try {
      const res = await api.get(`/api/processing/${result.sequenceId}/output`, {
        responseType: 'blob',
      });
      const url = URL.createObjectURL(res.data);
      const a = document.createElement('a');
      a.href = url;
      a.download = result.outputFilename || 'ICM_Output.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      alert('Download failed: ' + (err.response?.data?.detail || err.message));
    }
  };

  // ── Reset ───────────────────────────────────────────────────────
  const handleReset = () => {
    setFiles({});
    setRunName('');
    setResult(null);
  };

  const filledCount = Object.keys(files).length;

  return (
    <div className="processing-page">
      <Navbar />
      <div className="processing-content">
        <div className="processing-header">
          <h1>ICM Report Processing</h1>
          <p>Upload the 4 required Excel files (and optionally a Report Inputs file) to run intercompany matching.</p>
        </div>

        <div className="upload-card">
          <div className="upload-card-title">Upload Files</div>
          <div className="upload-card-subtitle">
            Drag & drop or click each zone to select the required .xlsx files.
          </div>

          {/* Run name */}
          <div className="processing-name-row">
            <div className="form-group">
              <label htmlFor="run-name">Run Name (optional)</label>
              <input
                id="run-name"
                type="text"
                className="input-field"
                placeholder="e.g. March 2026 IC Run"
                value={runName}
                onChange={(e) => setRunName(e.target.value)}
                disabled={processing}
              />
            </div>
          </div>

          {/* File drop zones */}
          <div className="file-grid">
            {FILE_SLOTS.map((slot) => {
              const file = files[slot.key];
              const isDragOver = dragKey === slot.key;
              return (
                <div
                  key={slot.key}
                  className={`drop-zone ${isDragOver ? 'drag-over' : ''} ${file ? 'has-file' : ''}`}
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
                    onChange={(e) => {
                      handleFile(slot.key, e.target.files[0]);
                      e.target.value = '';
                    }}
                  />
                  <div className="drop-zone-icon">
                    {file ? <FiFileText /> : <FiUploadCloud />}
                  </div>
                  <div className="drop-zone-label">{slot.label}</div>
                  {file ? (
                    <div className="drop-zone-file">
                      {file.name}
                      <button
                        className="remove-file-btn"
                        onClick={(e) => {
                          e.stopPropagation();
                          removeFile(slot.key);
                        }}
                        title="Remove file"
                      >
                        <FiX />
                      </button>
                    </div>
                  ) : (
                    <div className="drop-zone-hint">{slot.hint}</div>
                  )}
                </div>
              );
            })}
          </div>

          {/* Actions */}
          <div className="processing-actions">
            <span className="file-counter">
              <strong>{filledCount}</strong> / {FILE_SLOTS.length} files
              {' — '}
              {allRequiredFilled
                ? <span style={{ color: 'var(--accent-success)', fontWeight: 600 }}>✓ Ready to process</span>
                : <span>{requiredSlots.filter((s) => files[s.key]).length} / {requiredSlots.length} required</span>
              }
            </span>
            <button
              className="process-btn"
              disabled={!allRequiredFilled || processing}
              onClick={handleProcess}
            >
              <FiPlay size={15} />
              {processing ? 'Processing…' : 'Process Files'}
            </button>
          </div>
        </div>

        {/* Processing spinner */}
        {processing && (
          <div className="processing-state">
            <div className="spinner-container">
              <div className="spinner" />
              <div className="spinner-text">Running IC Matching…</div>
              <div className="spinner-subtext">
                This may take a minute depending on file size.
              </div>
            </div>
          </div>
        )}

        {/* Result */}
        {result && !processing && (
          <div className={`result-card ${result.status}`}>
            <div className="result-header">
              <div className={`result-icon ${result.status}`}>
                {result.status === 'success' ? (
                  <FiCheckCircle />
                ) : (
                  <FiAlertCircle />
                )}
              </div>
              <div className="result-title">
                {result.status === 'success'
                  ? 'Processing Complete'
                  : 'Processing Failed'}
              </div>
            </div>
            <div className="result-details">
              {result.status === 'success' ? (
                <>
                  Output file <strong>{result.outputFilename}</strong> is ready
                  for download.
                  <br />
                  Sequence ID: {result.sequenceId}
                </>
              ) : (
                <>Error: {result.error}</>
              )}
            </div>
            <div className="result-actions">
              {result.status === 'success' && (
                <button className="download-btn" onClick={handleDownload}>
                  <FiDownload size={14} />
                  Download Output
                </button>
              )}
              <button className="reset-btn" onClick={handleReset}>
                <FiRefreshCw size={14} style={{ marginRight: 6 }} />
                New Run
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
