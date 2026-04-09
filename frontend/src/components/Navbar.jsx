import { useState, useRef, useEffect } from 'react';
import { NavLink } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useTheme } from '../context/ThemeContext';
import { FiMoon, FiSun, FiChevronDown, FiLogOut, FiUser } from 'react-icons/fi';
import EditProfileModal from './EditProfileModal';
import { buildFileUrl } from '../config';
import './Navbar.css';

export default function Navbar() {
  const { user, logout } = useAuth();
  const { theme, toggleTheme } = useTheme();
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [profileOpen, setProfileOpen] = useState(false);
  const dropRef = useRef(null);

  useEffect(() => {
    function handleClick(e) {
      if (dropRef.current && !dropRef.current.contains(e.target)) {
        setDropdownOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, []);

  const initials = user?.full_name
    ? user.full_name.split(' ').filter(w => w).map((w) => w[0]).join('').slice(0, 2).toUpperCase()
    : 'U';

  return (
    <>
      <nav className="navbar">
        <div className="navbar-logo">
          <div className="logo-icon">
            <div className="logo-dot logo-dot-1" />
            <div className="logo-dot logo-dot-2" />
            <div className="logo-dot logo-dot-3" />
          </div>
          <span className="logo-text">
            <span className="logo-fccs">FCCS</span>
            <span className="logo-copilot">COPILOT</span>
          </span>
        </div>

        <div className="navbar-links">
          <NavLink to="/reports" className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}>Reports</NavLink>
        </div>

        <div className="navbar-right">
          <button className="theme-toggle" onClick={toggleTheme} title="Toggle theme">
            {theme === 'dark' ? <FiSun size={16} /> : <FiMoon size={16} />}
          </button>

          <div className="user-menu" ref={dropRef}>
            <button className="avatar-btn" onClick={() => setDropdownOpen(!dropdownOpen)}>
              {user?.profile_photo_url ? (
                <img src={buildFileUrl(user.profile_photo_url)} alt="avatar" className="avatar-img" />
              ) : (
                <div className="avatar-initials">{initials}</div>
              )}
              <FiChevronDown size={12} className={`chevron ${dropdownOpen ? 'open' : ''}`} />
            </button>

            {dropdownOpen && (
              <div className="user-dropdown">
                <div className="dropdown-header">
                  <div className="avatar-initials-sm">{initials}</div>
                  <div>
                    <div className="dropdown-name">{user?.full_name}</div>
                    <div className="dropdown-email">{user?.email}</div>
                  </div>
                </div>
                <div className="dropdown-divider" />
                <button
                  className="dropdown-item"
                  onClick={() => { setDropdownOpen(false); setProfileOpen(true); }}
                >
                  <FiUser size={13} />
                  View Profile
                </button>
                <div className="dropdown-divider" />
                <button className="dropdown-item danger" onClick={logout}>
                  <FiLogOut size={13} />
                  Log Out
                </button>
              </div>
            )}
          </div>
        </div>
      </nav>

      {profileOpen && (
        <EditProfileModal onClose={() => setProfileOpen(false)} />
      )}
    </>
  );
}
