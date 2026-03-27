import { useState, useRef } from 'react';
import { FiUploadCloud, FiFile, FiCheck, FiX } from 'react-icons/fi';
import api from '../services/api';

const REPORT_TYPES = ['alpha', 'beta', 'gamma', 'theta'];

export default function AddReportModal({ onClose, onSuccess }) {
  const [file, setFile] = useState(null);
  const [fileError, setFileError] = useState('');
  const [reportName, setReportName] = useState('');
  const [reportType, setReportType] = useState('alpha');
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const fileInputRef = useRef();

  const validateFile = (f) => {
    // Screenshot says (Supports .xlsx files only) but let's allow .csv and .xls too
    const allowed = ['.csv', '.xlsx', '.xls'];
    const ext = '.' + f.name.split('.').pop().toLowerCase();
    if (!allowed.includes(ext)) {
      setFileError('Only .csv, .xlsx, .xls files are allowed');
      return false;
    }
    setFileError('');
    return true;
  };

  const handleFileDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f && validateFile(f)) {
      setFile(f);
      if (!reportName) setReportName(f.name.replace(/\.[^.]+$/, ''));
    }
  };

  const handleFileChange = (e) => {
    const f = e.target.files[0];
    if (f && validateFile(f)) {
      setFile(f);
      if (!reportName) setReportName(f.name.replace(/\.[^.]+$/, ''));
    }
  };

  const handleSubmit = async () => {
    if (!file) {
      setFileError('Please select a file');
      return;
    }
    if (!reportName.trim()) {
      setFileError('Please enter a file name');
      return;
    }

    setLoading(true);
    try {
      const formData = new FormData();
      formData.append('name', reportName);
      formData.append('type_name', reportType);
      formData.append('file_type', 'input');
      formData.append('file', file);
      await api.post('/api/reports', formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
      });
      onSuccess?.(`Report: ${file.name} was saved`);
      onClose();
    } catch (err) {
      setFileError(err.response?.data?.detail || 'Upload failed');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 650 }}>
        <div className="modal-header">
          <div>
            <h2>Upload New Report</h2>
            <p>Add your report details and upload an Excel file to include it in your reports list.</p>
          </div>
          <button className="modal-close-btn" onClick={onClose}><FiX /></button>
        </div>

        {/* Report File Identity Section */}
        <div className="section-container">
          <h3 className="section-title-sm">Report File Identity</h3>
          <div className="form-grid-2">
            <div className="form-group">
              <label>File Name</label>
              <input 
                className="input-field" 
                value={reportName} 
                onChange={(e) => setReportName(e.target.value)} 
                placeholder="Enter File Name" 
              />
            </div>
            <div className="form-group">
              <label>Report Type</label>
              <select 
                className="input-field" 
                value={reportType} 
                onChange={(e) => setReportType(e.target.value)}
              >
                <option value="" disabled>Select Report Type</option>
                {REPORT_TYPES.map((t) => (
                  <option key={t} value={t}>{t.charAt(0).toUpperCase() + t.slice(1)}</option>
                ))}
              </select>
            </div>
          </div>
        </div>

        {/* File Upload Section */}
        <div 
          className={`drop-zone ${dragging ? 'dragging' : ''} ${file ? 'has-file' : ''}`}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleFileDrop}
          onClick={() => fileInputRef.current.click()}
        >
          <div className="drop-zone-header">
            <span style={{ fontWeight: 600, color: 'var(--text-primary)' }}>File Upload</span>
            <span style={{ fontSize: 11, color: 'var(--text-secondary)' }}> (Supports .xlsx, .csv files only)</span>
          </div>

          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: 'none' }} onChange={handleFileChange} />
          
          <div className="drop-zone-content">
            {file ? (
              <div style={{ textAlign: 'center' }}>
                <div style={{ color: 'var(--accent-success)', fontSize: 36, marginBottom: 8 }}>
                  <FiCheck />
                </div>
                <div style={{ fontSize: 13, fontWeight: 500 }}>{file.name}</div>
                <div style={{ fontSize: 11, color: 'var(--text-secondary)', marginTop: 4 }}>
                  {(file.size / 1024).toFixed(1)} KB — Click to change
                </div>
              </div>
            ) : (
              <div style={{ textAlign: 'center' }}>
                <div style={{ color: 'var(--accent-cyan)', fontSize: 42, marginBottom: 8 }}>
                  <FiUploadCloud />
                </div>
                <div style={{ fontSize: 14, fontWeight: 500, color: 'var(--accent-cyan)', marginBottom: 2 }}>
                  Drag & drop your file here
                </div>
                <div style={{ fontSize: 12, color: 'var(--text-secondary)', textDecoration: 'underline', cursor: 'pointer' }}>
                  or click to browse from your device
                </div>
              </div>
            )}
          </div>
        </div>

        {fileError && (
          <p style={{ fontSize: 12, color: 'var(--accent-danger)', marginTop: 8 }}>{fileError}</p>
        )}

        {/* Footer */}
        <div style={{ marginTop: 20, textAlign: 'right' }}>
          <button className="btn-primary btn-submit-report" onClick={handleSubmit} disabled={loading}>
            {loading ? 'Adding...' : 'Add Report'}
          </button>
        </div>
      </div>

      <style>{`
        .section-container {
          border: 1px solid var(--border-color);
          border-radius: var(--radius-md);
          padding: 16px 20px;
          margin-bottom: 20px;
          background: transparent;
        }
        .section-title-sm {
          font-size: 13px;
          font-weight: 600;
          color: var(--text-primary);
          margin-bottom: 12px;
        }
        .drop-zone {
          border: 1px dashed var(--accent-cyan);
          border-radius: var(--radius-md);
          padding: 16px;
          cursor: pointer;
          transition: all 0.15s;
          position: relative;
          background: rgba(8,145,178,0.02);
        }
        .drop-zone:hover, .drop-zone.dragging {
          background: rgba(8,145,178,0.06);
        }
        .drop-zone.has-file {
          border-color: var(--accent-success);
          background: rgba(22,163,74,0.04);
        }
        .drop-zone-header {
          position: absolute;
          top: 16px;
          left: 16px;
          font-size: 13px;
        }
        .drop-zone-content {
          padding: 60px 20px 40px;
          display: flex;
          align-items: center;
          justify-content: center;
        }
        .btn-submit-report {
          background: #ffffff;
          color: #1a1a1a;
          font-weight: 600;
          padding: 10px 24px;
        }
        [data-theme=\"light\"] .btn-submit-report {
          background: #1a1a1a;
          color: #ffffff;
        }
      `}</style>
    </div>
  );
}
