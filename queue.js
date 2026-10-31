// queue.js — кеш справочников/отметок и очередь несинхронизированных изменений в IndexedDB.

const DB_NAME = "timesheet";
const DB_VERSION = 1;
const STORE_CACHE = "cache";
const STORE_OUTBOX = "outbox";

let dbPromise = null;

function openDb() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_CACHE)) {
        db.createObjectStore(STORE_CACHE, { keyPath: "key" });
      }
      if (!db.objectStoreNames.contains(STORE_OUTBOX)) {
        db.createObjectStore(STORE_OUTBOX, { keyPath: "uuid" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

function tx(storeName, mode) {
  return openDb().then((db) => db.transaction(storeName, mode).objectStore(storeName));
}

function promisify(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

// ---------- кеш справочников и отметок месяца ----------

export async function setCache(key, value) {
  const store = await tx(STORE_CACHE, "readwrite");
  await promisify(store.put({ key, value, savedAt: new Date().toISOString() }));
}

export async function getCache(key) {
  const store = await tx(STORE_CACHE, "readonly");
  const row = await promisify(store.get(key));
  return row ? { value: row.value, savedAt: row.savedAt } : null;
}

// ---------- очередь на отправку ----------

export async function enqueue(record) {
  const store = await tx(STORE_OUTBOX, "readwrite");
  await promisify(store.put(record));
}

export async function getOutbox() {
  const store = await tx(STORE_OUTBOX, "readonly");
  return promisify(store.getAll());
}

export async function removeFromOutbox(uuid) {
  const store = await tx(STORE_OUTBOX, "readwrite");
  await promisify(store.delete(uuid));
}

export async function countOutbox() {
  const store = await tx(STORE_OUTBOX, "readonly");
  return promisify(store.count());
}

// Отправляет накопленную очередь через переданную функцию sendBatch(records) -> results.
// Успешно принятые (ok, включая дубликаты) удаляются из очереди; неуспешные остаются
// и получают пометку lastError для показа пользователю.
export async function flush(sendBatch) {
  const records = await getOutbox();
  if (!records.length) return { sent: 0, failed: 0 };

  let results;
  try {
    results = await sendBatch(records);
  } catch (err) {
    return { sent: 0, failed: records.length, networkError: true };
  }

  let sent = 0;
  let failed = 0;
  const byUuid = new Map(results.map((r) => [r.uuid, r]));
  for (const record of records) {
    const result = byUuid.get(record.uuid);
    if (result && result.ok) {
      await removeFromOutbox(record.uuid);
      sent++;
    } else {
      failed++;
      const store = await tx(STORE_OUTBOX, "readwrite");
      await promisify(store.put({ ...record, lastError: result ? result.error : "Нет ответа" }));
    }
  }
  return { sent, failed };
}
