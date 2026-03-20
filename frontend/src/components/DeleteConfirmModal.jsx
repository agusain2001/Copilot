import { FiAlertCircle } from 'react-icons/fi';

export default function DeleteConfirmModal({ reports, onConfirm, onClose }) {
  const isBulk = reports.length > 1;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 480 }}>
        <div style={{ textAlign: 'center', marginBottom: 16 }}>
          <div style={{
            width: 52, height: 52, borderRadius: '50%',
            background: 'rgba(220,38,38,0.1)', border: '2px solid var(--accent-danger)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            margin: '0 auto 12px'
          }}>
            <FiAlertCircle size={22} color="var(--accent-danger)" />
          </div>
          <h2 style={{ fontSize: 16, fontWeight: 600 }}>
            {isBulk ? 'Delete Report Files' : 'Delete Report File'}
          </h2>
          <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginTop: 4 }}>
            This action permanently removes the report files from your list.
            You won't be able to recover them.
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

        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
          <button className="btn-ghost" onClick={onClose}>Cancel</button>
          <button className="btn-danger" onClick={onConfirm}>
            {isBulk ? `Delete (${reports.length})` : 'Delete'}
          </button>
        </div>
      </div>
    </div>
  );
}
