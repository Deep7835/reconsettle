// Thin fetch wrapper around the SettleOps FastAPI backend.
// All methods return parsed JSON or throw with a useful message.

const BASE = import.meta.env.VITE_API_BASE || "http://localhost:8000";
const AUTH_KEY = "settleops:auth";

// --- Token persistence helpers ---

export function getStoredAuth() {
  try {
    return JSON.parse(localStorage.getItem(AUTH_KEY) || "null");
  } catch {
    return null;
  }
}

export function setStoredAuth(value) {
  if (value) localStorage.setItem(AUTH_KEY, JSON.stringify(value));
  else localStorage.removeItem(AUTH_KEY);
}

export function clearAuth() {
  localStorage.removeItem(AUTH_KEY);
}

function authHeader() {
  const auth = getStoredAuth();
  return auth?.access_token ? { Authorization: `Bearer ${auth.access_token}` } : {};
}

async function request(path, { method = "GET", body, headers, isForm, skipAuth } = {}) {
  const opts = { method, headers: { ...(skipAuth ? {} : authHeader()), ...headers } };
  if (body !== undefined) {
    if (isForm) {
      opts.body = body;
    } else {
      opts.headers["Content-Type"] = "application/json";
      opts.body = JSON.stringify(body);
    }
  }
  const r = await fetch(`${BASE}${path}`, opts);

  // On 401 anywhere except the login call itself, clear the token and notify
  // the app so it can show the login screen.
  if (r.status === 401 && !skipAuth) {
    clearAuth();
    window.dispatchEvent(new CustomEvent("settleops:auth-expired"));
  }

  const text = await r.text();
  let parsed = null;
  try {
    parsed = text ? JSON.parse(text) : null;
  } catch {
    parsed = text;
  }
  if (!r.ok) {
    const detail = parsed?.detail || parsed?.message || text || r.statusText;
    const err = new Error(typeof detail === "string" ? detail : JSON.stringify(detail));
    err.status = r.status;
    err.body = parsed;
    throw err;
  }
  return parsed;
}

export const apiBase = BASE;

export const api = {
  // ---- auth ----
  login: (username, password) =>
    request("/api/auth/login", {
      method: "POST",
      body: { username, password },
      skipAuth: true,
    }),
  me: () => request("/api/auth/me"),
  logout: () => {
    clearAuth();
    window.dispatchEvent(new CustomEvent("settleops:auth-expired"));
  },

  // ---- meta ----
  health: () => request("/api/health", { skipAuth: true }),

  // ---- dashboard ----
  dashboard: () => request("/api/dashboard/stats"),

  // ---- entries ----
  listEntries: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/api/entries${qs ? `?${qs}` : ""}`);
  },
  createEntry: (body) => request("/api/entries", { method: "POST", body }),
  deleteEntry: (id) => request(`/api/entries/${id}`, { method: "DELETE" }),
  clearAllEntries: () => request("/api/entries", { method: "DELETE" }),
  entryTotals: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/api/entries/totals${qs ? `?${qs}` : ""}`);
  },
  entryGrouped: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/api/entries/grouped${qs ? `?${qs}` : ""}`);
  },
  exportEntriesUrl: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return `${BASE}/api/entries/export.xlsx${qs ? `?${qs}` : ""}`;
  },

  // ---- recon ----
  listRecon: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/api/recon${qs ? `?${qs}` : ""}`);
  },
  createRecon: (body) => request("/api/recon", { method: "POST", body }),
  updateRecon: (id, body) => request(`/api/recon/${id}`, { method: "PATCH", body }),
  deleteRecon: (id) => request(`/api/recon/${id}`, { method: "DELETE" }),
  clearAllRecon: () => request("/api/recon", { method: "DELETE" }),
  reconStats: () => request("/api/recon/stats"),
  uploadBank: (file) => {
    const fd = new FormData();
    fd.append("file", file);
    return request("/api/recon/upload/bank", { method: "POST", body: fd, isForm: true });
  },
  uploadSource: (file) => {
    const fd = new FormData();
    fd.append("file", file);
    return request("/api/recon/upload/source", { method: "POST", body: fd, isForm: true });
  },
  listBankUploads: () => request("/api/recon/uploads/bank"),
  listSourceUploads: () => request("/api/recon/uploads/source"),
  aggregated: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/api/recon/aggregated${qs ? `?${qs}` : ""}`);
  },

  // ---- merchants ----
  listMerchants: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/api/merchants${qs ? `?${qs}` : ""}`);
  },
  merchantSummary: () => request("/api/merchants/summary"),
  addCompany: (body) => request("/api/merchants/companies", { method: "POST", body }),

  // ---- gmail ----
  gmailConfig: () => request("/api/gmail/config"),
  updateGmailConfig: (body) => request("/api/gmail/config", { method: "PATCH", body }),
  gmailAuthStart: () => request("/api/gmail/auth/start"),
  gmailMessages: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/api/gmail/messages${qs ? `?${qs}` : ""}`);
  },
  gmailSync: () => request("/api/gmail/sync", { method: "POST" }),
  gmailRuns: (limit = 20) => request(`/api/gmail/runs?limit=${limit}`),

  // ---- telegram ----
  tgConfig: () => request("/api/telegram/config"),
  updateTgConfig: (body) => request("/api/telegram/config", { method: "PATCH", body }),
  tgMe: () => request("/api/telegram/me"),
  tgSync: () => request("/api/telegram/sync", { method: "POST" }),
  tgResetState: () => request("/api/telegram/state/reset", { method: "POST" }),
  tgRuns: (limit = 30) => request(`/api/telegram/runs?limit=${limit}`),
  tgSendMessage: (body) => request("/api/telegram/send/message", { method: "POST", body }),
  tgSendReport: (body = {}) => request("/api/telegram/send/report", { method: "POST", body }),

  // ---- telegram recipients ----
  tgRecipients: () => request("/api/telegram/recipients"),
  tgAddRecipient: (body) => request("/api/telegram/recipients", { method: "POST", body }),
  tgUpdateRecipient: (id, body) => request(`/api/telegram/recipients/${id}`, { method: "PATCH", body }),
  tgDeleteRecipient: (id) => request(`/api/telegram/recipients/${id}`, { method: "DELETE" }),
  tgSendReportTo: (id) => request(`/api/telegram/recipients/${id}/send/report`, { method: "POST" }),
};
