import './Toast.css';

export default function Toast({ message, onUndo, onClose }) {
  return (
    <div className="toast">
      <span className="toast-message">{message}</span>
      <div className="toast-actions">
        {onUndo && (
          <button className="toast-undo" onClick={onUndo}>Undo</button>
        )}
        <button className="toast-close" onClick={onClose}>✕</button>
      </div>
    </div>
  );
}
