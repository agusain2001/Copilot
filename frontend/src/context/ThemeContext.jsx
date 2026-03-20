import { createContext, useContext, useState, useEffect } from 'react';
import api from '../services/api';

const ThemeContext = createContext(null);

export function ThemeProvider({ children }) {
  const [theme, setTheme] = useState(() => localStorage.getItem('theme') || 'dark');

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
  }, [theme]);

  const toggleTheme = async () => {
    const next = theme === 'dark' ? 'light' : 'dark';
    setTheme(next);
    // Persist to server if logged in
    const token = localStorage.getItem('access_token');
    if (token) {
      try {
        await api.put('/api/users/me/preferences', { theme: next });
      } catch {}
    }
  };

  const syncThemeFromServer = async () => {
    try {
      const res = await api.get('/api/users/me/preferences');
      setTheme(res.data.theme || 'dark');
    } catch {}
  };

  return (
    <ThemeContext.Provider value={{ theme, toggleTheme, syncThemeFromServer }}>
      {children}
    </ThemeContext.Provider>
  );
}

export function useTheme() {
  return useContext(ThemeContext);
}
