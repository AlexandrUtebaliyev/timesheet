// config.js — настройки приложения в localStorage. Ничего в репозиторий не попадает.

const KEYS = {
  scriptUrl: "timesheet.scriptUrl",
  pin: "timesheet.pin",
  theme: "timesheet.theme", // "auto" | "light" | "dark"
  deviceId: "timesheet.deviceId",
  normHoursOverride: "timesheet.normHoursOverride",
};

export function getScriptUrl() {
  return localStorage.getItem(KEYS.scriptUrl) || "";
}

export function setScriptUrl(url) {
  localStorage.setItem(KEYS.scriptUrl, url.trim());
}

export function getPin() {
  return localStorage.getItem(KEYS.pin) || "";
}

export function setPin(pin) {
  localStorage.setItem(KEYS.pin, pin.trim());
}

export function getTheme() {
  return localStorage.getItem(KEYS.theme) || "auto";
}

export function setTheme(theme) {
  localStorage.setItem(KEYS.theme, theme);
}

// Норма часов приходит с сервера (action=init), но пользователь может переопределить
// её локально в «Настройках» — это влияет только на подсказки/цвета/переработку в
// приложении на этом устройстве, саму Google-таблицу не меняет (см. DECISIONS.md).
export function getNormHoursOverride() {
  const raw = localStorage.getItem(KEYS.normHoursOverride);
  if (!raw) return null;
  const n = Number(raw);
  return Number.isFinite(n) && n > 0 ? n : null;
}

export function setNormHoursOverride(value) {
  if (value === null || value === "" || !Number.isFinite(Number(value)) || Number(value) <= 0) {
    localStorage.removeItem(KEYS.normHoursOverride);
  } else {
    localStorage.setItem(KEYS.normHoursOverride, String(Number(value)));
  }
}

export function getDeviceId() {
  let id = localStorage.getItem(KEYS.deviceId);
  if (!id) {
    id = "dev-" + Math.random().toString(36).slice(2) + Date.now().toString(36);
    localStorage.setItem(KEYS.deviceId, id);
  }
  return id;
}

export function isDemoMode() {
  return new URLSearchParams(location.search).get("demo") === "1";
}

export function isConfigured() {
  return isDemoMode() || (!!getScriptUrl() && !!getPin());
}
