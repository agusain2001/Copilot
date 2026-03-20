import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useTheme } from '../context/ThemeContext';
import { FiEye, FiEyeOff } from 'react-icons/fi';
import './LoginPage.css';

export default function LoginPage() {
  const { login } = useAuth();
  const { syncThemeFromServer } = useTheme();
  const navigate = useNavigate();

  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [rememberMe, setRememberMe] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);
    try {
      await login(username, password, rememberMe);
      await syncThemeFromServer();
      navigate('/reports');
    } catch (err) {
      setError(err.response?.data?.detail || 'Invalid username or password');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="login-page">
      {/* Decorative curved arcs */}
      <div className="arc arc-teal" />
      <div className="arc arc-gold" />

      {/* FCCS Logo */}
      <div className="login-logo">
        <div className="logo-icon-sm">
          <span className="logo-dot-a" />
          <span className="logo-dot-b" />
        </div>
        <span className="login-logo-text">
          <span style={{ color: '#f97316', fontWeight: 700 }}>FCCS</span>
          <span style={{ color: '#888', fontSize: '9px', letterSpacing: '1px', fontWeight: 500 }}>COPILOT</span>
        </span>
      </div>

      {/* Login Card */}
      <div className="login-card">
        <h1 className="login-title">Log in</h1>
        <p className="login-subtitle">Please login to continue to your account.</p>

        <form onSubmit={handleSubmit} className="login-form">
          <input
            id="login-username"
            type="text"
            className="input-field"
            placeholder="Enter Username"
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            required
            autoComplete="username"
          />

          <div className="password-wrapper">
            <input
              id="login-password"
              type={showPassword ? 'text' : 'password'}
              className="input-field"
              placeholder="Enter Password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              required
              autoComplete="current-password"
            />
            <button
              type="button"
              className="toggle-pw"
              onClick={() => setShowPassword(!showPassword)}
              tabIndex={-1}
            >
              {showPassword ? <FiEyeOff size={15} /> : <FiEye size={15} />}
            </button>
          </div>

          {error && <p className="login-error">{error}</p>}

          <label className="remember-me">
            <input
              type="checkbox"
              checked={rememberMe}
              onChange={(e) => setRememberMe(e.target.checked)}
            />
            <span>Keep me logged in</span>
          </label>

          <button
            id="login-submit"
            type="submit"
            className="login-btn"
            disabled={loading}
          >
            {loading ? 'Logging in...' : 'Log in'}
          </button>
        </form>
      </div>
    </div>
  );
}
