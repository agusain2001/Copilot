const rawApiBaseUrl = import.meta.env.VITE_API_BASE_URL?.trim();
const rawUploadsBaseUrl = import.meta.env.VITE_UPLOADS_BASE_URL?.trim();

export const API_BASE_URL = rawApiBaseUrl || '';
export const UPLOADS_BASE_URL =
  rawUploadsBaseUrl || '';

export function buildFileUrl(path) {
  if (!path) return '';

  if (/^https?:\/\//i.test(path)) return path;

  const base = UPLOADS_BASE_URL.replace(/\/+$/, '');
  const normalizedPath = path.startsWith('/') ? path : `/${path}`;

  return `${base}${normalizedPath}`;
}
