import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import Navbar from '../components/Navbar';
import api from '../services/api';
import {
  FiDownload,
  FiRefreshCw,
  FiCheckCircle,
  FiAlertCircle,
  FiClock,
  FiArrowLeft,
} from 'react-icons/fi';
import './ProcessingPage.css';

export default function ProcessingPage() {
  const navigate = useNavigate();
  const [runs, setRuns]     = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fetchRuns = async () => {
      setLoading(true);
      try {
        const res = await api.get('/api/processing/runs');
        setRuns(res.data || []);
      } catch {
        // If endpoint doesn't exist yet, show empty state gracefully
        setRuns([]);
      } finally {
        setLoading(false);
      }
    };
    fetchRuns();
  }, []);

  const handleDownload = async (run) => {
    try {
      const res = await api.get(`/api/processing/${run.sequence_id}/output`, {
        responseType: 'blob',
      });
      const url = URL.createObjectURL(res.data);
      const a   = document.createElement('a');
      a.href     = url;
      a.download = run.output_filename || 'ICM_Output.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      alert('Download failed: ' + (err.response?.data?.detail || err.message));
    }
  };

  const formatTime = (ts) => {
    if (!ts) return '—';
    const d = new Date(ts);
    return d.toLocaleString('en-IN', {
      day: '2-digit', month: 'short', year: 'numeric',
      hour: '2-digit', minute: '2-digit', hour12: false,
    });
  };

  return (
    <div className="processing-page">
      <Navbar />
      <div className="processing-content">

        {/* Header */}
        <div className="processing-header">
          <h1>Processing History</h1>
          <p>
            View previous IC Matching runs. To start a new run, click{' '}
            <strong style={{ color: 'var(--accent-cyan)' }}>+ Add Report</strong> on the Reports page.
          </p>
        </div>

        {/* Back to Reports button */}
        <button
          className="back-to-reports-btn"
          onClick={() => navigate('/reports')}
        >
          <FiArrowLeft size={14} />
          Go to Reports &amp; Add Report
        </button>

        {/* Runs list */}
        <div className="runs-section">
          {loading ? (
            <div className="processing-state">
              <div className="spinner-container">
                <div className="spinner" />
                <div className="spinner-text">Loading history…</div>
              </div>
            </div>
          ) : runs.length === 0 ? (
            <div className="empty-runs">
              <div className="empty-runs-icon"><FiClock /></div>
              <div className="empty-runs-title">No Processing Runs Yet</div>
              <div className="empty-runs-sub">
                Use <strong>+ Add Report</strong> on the Reports page to upload your files and run IC Matching.
              </div>
              <button className="process-btn" onClick={() => navigate('/reports')}>
                <FiArrowLeft size={14} /> Go to Reports
              </button>
            </div>
          ) : (
            <div className="runs-list">
              {runs.map((run) => (
                <div key={run.sequence_id} className={`run-card ${run.status}`}>
                  <div className="run-card-header">
                    <div className={`run-status-icon ${run.status}`}>
                      {run.status === 'success'
                        ? <FiCheckCircle />
                        : <FiAlertCircle />}
                    </div>
                    <div className="run-card-info">
                      <div className="run-card-name">{run.name}</div>
                      <div className="run-card-time">{formatTime(run.created_at)}</div>
                    </div>
                    <div className="run-card-actions">
                      {run.status === 'success' && run.output_filename && (
                        <button className="download-btn" onClick={() => handleDownload(run)}>
                          <FiDownload size={13} /> Download
                        </button>
                      )}
                    </div>
                  </div>
                  {run.status === 'success' && run.output_filename && (
                    <div className="run-card-detail">
                      Output: <strong>{run.output_filename}</strong>
                      &nbsp;·&nbsp; Seq: {run.sequence_id}
                    </div>
                  )}
                  {run.status === 'error' && run.error && (
                    <div className="run-card-detail error-text">Error: {run.error}</div>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>

      </div>
    </div>
  );
}
