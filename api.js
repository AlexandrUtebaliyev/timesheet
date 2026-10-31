// api.js — обращения к Apps Script веб-приложению. В режиме ?demo=1 подменяется локальным
// хранилищем на фикстуре demo-data.js, чтобы интерфейс работал без сети и без таблицы.

import { getScriptUrl, getPin, getDeviceId, isDemoMode } from "./config.js";
import { buildDemoSnapshot } from "./demo-data.js";

const DEMO_KEY = "timesheet.demoJournal";

function loadDemoJournal() {
  const raw = localStorage.getItem(DEMO_KEY);
  if (raw) return JSON.parse(raw);
  const snapshot = buildDemoSnapshot();
  localStorage.setItem(DEMO_KEY, JSON.stringify(snapshot.journal));
  return snapshot.journal;
}

function saveDemoJournal(journal) {
  localStorage.setItem(DEMO_KEY, JSON.stringify(journal));
}

export class ApiError extends Error {}

async function callScript(action, params) {
  const url = new URL(getScriptUrl());
  url.searchParams.set("action", action);
  url.searchParams.set("pin", getPin());
  for (const [k, v] of Object.entries(params || {})) url.searchParams.set(k, v);

  const res = await fetch(url.toString());
  const data = await res.json();
  if (!data.ok) throw new ApiError(data.error || "Ошибка сервера");
  return data;
}

export async function fetchInit() {
  if (isDemoMode()) {
    const snapshot = buildDemoSnapshot();
    // ?demo=1&role=viewer|timekeeper — проверить вид/поведение ограниченной роли без реального PIN.
    const roleParam = new URLSearchParams(location.search).get("role");
    const role = roleParam === "viewer" || roleParam === "timekeeper" ? roleParam : "master";
    return { employees: snapshot.employees, codes: snapshot.codes, normHours: snapshot.normHours, role };
  }
  const data = await callScript("init");
  return { employees: data.employees, codes: data.codes, normHours: data.normHours, role: data.role || "master" };
}

export async function fetchMonth(month) {
  if (isDemoMode()) {
    const journal = loadDemoJournal();
    return { entries: latestEntriesForMonth(journal, month) };
  }
  const data = await callScript("month", { month });
  return { entries: data.entries };
}

function latestEntriesForMonth(journal, month) {
  const latest = new Map();
  for (const entry of journal) {
    if (entry.date.slice(0, 7) !== month) continue;
    latest.set(`${entry.date}__${entry.tabNum}`, entry);
  }
  return [...latest.values()];
}

// records: [{uuid, date, id, value, clientTs, deviceId}]
export async function sendBatch(records) {
  if (isDemoMode()) {
    const journal = loadDemoJournal();
    const known = new Set(journal.map((e) => e.uuid));
    const results = [];
    for (const rec of records) {
      if (known.has(rec.uuid)) {
        results.push({ uuid: rec.uuid, ok: true, duplicate: true });
        continue;
      }
      journal.push({
        date: rec.date,
        tabNum: rec.id,
        value: rec.value,
        who: "мастер",
        ts: new Date().toISOString(),
        device: rec.deviceId,
        uuid: rec.uuid,
      });
      results.push({ uuid: rec.uuid, ok: true });
    }
    saveDemoJournal(journal);
    return results;
  }

  const url = new URL(getScriptUrl());
  const res = await fetch(url.toString(), {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ pin: getPin(), records }),
  });
  const data = await res.json();
  if (!data.ok) throw new ApiError(data.error || "Ошибка сервера");
  return data.results;
}

export function makeRecord(date, tabNum, value, knownTs) {
  return {
    uuid: (crypto.randomUUID && crypto.randomUUID()) || `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    date,
    id: tabNum,
    tabNum,
    value,
    clientTs: new Date().toISOString(),
    deviceId: getDeviceId(),
    knownTs,
    demo: isDemoMode(), // чтобы очередь никогда не путала демо-правки с боевыми, см. DECISIONS.md
  };
}
