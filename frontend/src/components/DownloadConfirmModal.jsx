import { FiDownload } from 'react-icons/fi';

export default function DownloadConfirmModal({ reports, onConfirm, onClose }) {
  const isBulk = reports.length > 1;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 480 }}>
        <div style={{ textAlign: 'center', marginBottom: 16 }}>
          <div style={{
            width: 52, height: 52, borderRadius: '50%',
            background: 'rgba(8,145,178,0.12)', border: '2px solid var(--accent-cyan)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            margin: '0 auto 12px'
          }}>
            <FiDownload size={22} color="var(--accent-cyan)" />
          </div>
          <h2 style={{ fontSize: 16, fontWeight: 600 }}>
            {isBulk ? 'Download Files' : 'Download File'}
          </h2>
          <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginTop: 4 }}>
            {isBulk
              ? 'The selected reports will be downloaded to your device as a ZIP file.'
              : 'The selected report will be downloaded to your device.'}
          </p>
        </div>

        <table className="data-table" style={{ marginBottom: 20 }}>
          <thead>
            <tr>
              <th>Report Name</th>
              <th>Type</th>
            </tr>
          </thead>
          <tbody>
            {reports.map((r) => (
              <tr key={r.id}>
                <td>{r.name}</td>
                <td>{r.type_label}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <button className="btn-primary" style={{ width: '100%' }} onClick={onConfirm}>
          Download
        </button>
      </div>
    </div>
  );
}
