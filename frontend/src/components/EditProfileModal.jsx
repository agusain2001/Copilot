import { useState, useRef } from 'react';
import { FiX, FiCamera, FiCopy, FiEye, FiEyeOff } from 'react-icons/fi';
import { useAuth } from '../context/AuthContext';
import api from '../services/api';
import { buildFileUrl } from '../config';
import './EditProfileModal.css';

export default function EditProfileModal({ onClose }) {
  const { user, updateUser, logout } = useAuth();
  const [fullName, setFullName] = useState(user?.full_name || '');
  const [email, setEmail] = useState(user?.email || '');
  const [phone, setPhone] = useState(user?.phone || '');
  const [saving, setSaving] = useState(false);
  const [saveMsg, setSaveMsg] = useState('');

  const [pwExpanded, setPwExpanded] = useState(false);
  const [currentPw, setCurrentPw] = useState('');
  const [showCurrentPw, setShowCurrentPw] = useState(false);
  const [newPw, setNewPw] = useState('');
  const [confirmPw, setConfirmPw] = useState('');
  const [pwError, setPwError] = useState('');
  const [pwSuccess, setPwSuccess] = useState('');

  const photoRef = useRef();

  const daysSincePasswordChange = () => {
    if (!user?.password_changed_at) return null;
    const diff = Date.now() - new Date(user.password_changed_at).getTime();
    return Math.floor(diff / (1000 * 60 * 60 * 24));
  };

  const handleSave = async () => {
    setSaving(true); setSaveMsg('');
    try {
      const res = await api.put('/api/users/me', { full_name: fullName, email, phone });
      updateUser(res.data);
      setSaveMsg('Changes saved!');
      setTimeout(() => setSaveMsg(''), 2500);
    } catch {
      setSaveMsg('Failed to save changes.');
    } finally {
      setSaving(false);
    }
  };

  const handlePhotoUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const fd = new FormData(); fd.append('file', file);
    try {
      const res = await api.post('/api/users/me/photo', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      updateUser({ profile_photo_url: res.data.profile_photo_url });
    } catch {}
  };

  const handleChangePassword = async () => {
    setPwError(''); setPwSuccess('');
    try {
      await api.put('/api/users/me/password', {
        current_password: currentPw,
        new_password: newPw,
        confirm_password: confirmPw,
      });
      setPwSuccess('Password changed successfully!');
      setPwExpanded(false);
      setCurrentPw(''); setNewPw(''); setConfirmPw('');
    } catch (err) {
      setPwError(err.response?.data?.detail || 'Password change failed');
    }
  };

  const initials = user
    ? user.full_name.split(' ').map((w) => w[0]).join('').slice(0, 2).toUpperCase()
    : 'U';

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box profile-modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2>Edit Profile Information</h2>
            <p>Update your personal details and contact information.</p>
          </div>
          <button className="modal-close-btn" onClick={onClose}><FiX /></button>
        </div>

        <p className="section-title">Basic Information</p>
        <div className="profile-photo-row">
          <div className="profile-photo-wrap" onClick={() => photoRef.current.click()}>
            {user?.profile_photo_url ? (
              <img src={buildFileUrl(user.profile_photo_url)} alt="profile" className="profile-img" />
            ) : (
              <div className="profile-initials">{initials}</div>
            )}
            <div className="photo-overlay"><FiCamera size={16} /></div>
            <input ref={photoRef} type="file" accept=".jpg,.jpeg,.png" style={{ display: 'none' }} onChange={handlePhotoUpload} />
          </div>
          <div>
            <div style={{ fontSize: 12, fontWeight: 500, color: 'var(--text-primary)' }}>Edit Profile Photo</div>
            <div style={{ fontSize: 11, color: 'var(--text-secondary)' }}>Upload a File (.jpg, .png, .jpeg)</div>
          </div>
        </div>

        <div className="form-grid-2">
          <div className="form-group">
            <label>Full Name</label>
            <input className="input-field" value={fullName} onChange={(e) => setFullName(e.target.value)} />
          </div>
          <div className="form-group">
            <label>Employee ID</label>
            <div style={{ position: 'relative' }}>
              <input className="input-field" value={user?.employee_id || ''} readOnly style={{ paddingRight: 36 }} />
              <button
                style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: 'var(--text-secondary)' }}
                onClick={() => navigator.clipboard.writeText(user?.employee_id || '')}
              >
                <FiCopy size={14} />
              </button>
            </div>
          </div>
          <div className="form-group">
            <label>Email Address</label>
            <input className="input-field" type="email" value={email} onChange={(e) => setEmail(e.target.value)} />
          </div>
          <div className="form-group">
            <label>Phone Number</label>
            <input className="input-field" value={phone} onChange={(e) => setPhone(e.target.value)} />
          </div>
        </div>

        <hr className="divider" />

        <div style={{ marginBottom: 16 }}>
          <p className="section-title" style={{ marginBottom: 4 }}>Password Management</p>
          <p style={{ fontSize: 11, color: 'var(--text-secondary)', marginBottom: 14 }}>
            Update your account password to keep your profile secure.
          </p>

          <div style={{ display: 'flex', alignItems: 'flex-start', gap: 12 }}>
            <div style={{ flex: 1 }}>
              <label style={{ fontSize: 12, color: 'var(--text-secondary)', display: 'block', marginBottom: 6 }}>
                Current Password{pwExpanded ? 's' : ''}
              </label>
              <div style={{ position: 'relative' }}>
                <input
                  className="input-field"
                  type={showCurrentPw ? 'text' : 'password'}
                  value={pwExpanded ? currentPw : '••••••••••'}
                  onChange={pwExpanded ? (e) => setCurrentPw(e.target.value) : undefined}
                  readOnly={!pwExpanded}
                  style={{ paddingRight: 36 }}
                />
                <button
                  style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: 'var(--text-muted)' }}
                  onClick={() => setShowCurrentPw(!showCurrentPw)}
                >
                  {showCurrentPw ? <FiEyeOff size={14} /> : <FiEye size={14} />}
                </button>
              </div>
              {daysSincePasswordChange() !== null && (
                <p style={{ fontSize: 11, color: 'var(--text-muted)', marginTop: 4 }}>
                  Last changed {daysSincePasswordChange()} days ago
                </p>
              )}
            </div>

            {!pwExpanded && (
              <button
                className="btn-ghost"
                style={{ marginTop: 20, whiteSpace: 'nowrap' }}
                onClick={() => setPwExpanded(true)}
              >
                Change Password →
              </button>
            )}
          </div>

          {pwExpanded && (
            <div className="form-grid-2" style={{ marginTop: 14 }}>
              <div className="form-group">
                <label>Create New Password</label>
                <input className="input-field" type="password" placeholder="Enter New Password" value={newPw} onChange={(e) => setNewPw(e.target.value)} />
              </div>
              <div className="form-group">
                <label>Confirm New Password</label>
                <input className="input-field" type="password" placeholder="Confirm New Password" value={confirmPw} onChange={(e) => setConfirmPw(e.target.value)} />
              </div>
              {pwError && <p style={{ fontSize: 12, color: 'var(--accent-danger)', gridColumn: 'span 2' }}>{pwError}</p>}
              {pwSuccess && <p style={{ fontSize: 12, color: 'var(--accent-success)', gridColumn: 'span 2' }}>{pwSuccess}</p>}
              <div style={{ gridColumn: 'span 2', textAlign: 'right' }}>
                <button className="btn-ghost" style={{ marginRight: 8 }} onClick={() => { setPwExpanded(false); setPwError(''); }}>Cancel</button>
                <button className="btn-primary" onClick={handleChangePassword}>Save Password</button>
              </div>
            </div>
          )}
        </div>

        <hr className="divider" />
        {saveMsg && <p style={{ fontSize: 12, color: saveMsg.includes('!') ? 'var(--accent-success)' : 'var(--accent-danger)', marginBottom: 10 }}>{saveMsg}</p>}
        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
          <button className="btn-danger" onClick={logout}>Log Out</button>
          <button className="btn-primary" onClick={handleSave} disabled={saving}>
            {saving ? 'Saving...' : 'Save Changes'}
          </button>
        </div>
      </div>
    </div>
  );
}
