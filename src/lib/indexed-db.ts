import type { SavedCloudReport } from "./cloud-reports";

const DB_NAME = "tempos_movimentos_db";
const DB_VERSION = 1;
const STORE_NAME = "reports";

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof window === "undefined" || !window.indexedDB) {
      return reject(new Error("IndexedDB not supported in this environment"));
    }

    const request = window.indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };

    request.onsuccess = () => {
      resolve(request.result);
    };

    request.onerror = () => {
      reject(request.error || new Error("Failed to open IndexedDB"));
    };
  });
}

export async function saveReportsToIndexedDB(reports: SavedCloudReport[]): Promise<void> {
  if (typeof window === "undefined" || !window.indexedDB || !Array.isArray(reports)) return;
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_NAME, "readwrite");
    const store = tx.objectStore(STORE_NAME);

    for (const report of reports) {
      if (report && report.id) {
        store.put(report);
      }
    }

    return new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => {
        db.close();
        resolve();
      };
      tx.onerror = () => {
        db.close();
        reject(tx.error);
      };
    });
  } catch (err) {
    console.warn("[IndexedDB] Error saving reports:", err);
  }
}

export async function getReportsFromIndexedDB(): Promise<SavedCloudReport[]> {
  if (typeof window === "undefined" || !window.indexedDB) return [];
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_NAME, "readonly");
    const store = tx.objectStore(STORE_NAME);
    const request = store.getAll();

    return new Promise<SavedCloudReport[]>((resolve, reject) => {
      request.onsuccess = () => {
        db.close();
        const results = request.result;
        if (Array.isArray(results)) {
          resolve(results as SavedCloudReport[]);
        } else {
          resolve([]);
        }
      };
      request.onerror = () => {
        db.close();
        reject(request.error);
      };
    });
  } catch (err) {
    console.warn("[IndexedDB] Error fetching reports:", err);
    return [];
  }
}

export async function deleteReportFromIndexedDB(id: string): Promise<void> {
  if (typeof window === "undefined" || !window.indexedDB || !id) return;
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_NAME, "readwrite");
    const store = tx.objectStore(STORE_NAME);
    store.delete(id);

    return new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => {
        db.close();
        resolve();
      };
      tx.onerror = () => {
        db.close();
        reject(tx.error);
      };
    });
  } catch (err) {
    console.warn("[IndexedDB] Error deleting report:", err);
  }
}
