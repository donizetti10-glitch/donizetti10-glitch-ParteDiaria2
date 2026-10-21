import { createFileRoute } from "@tanstack/react-router";
import fs from "node:fs";
import path from "node:path";

export interface DefectRecord {
  value: string;
  timestamp: number;
}

const DATA_DIR = path.join(process.cwd(), "data");
const JSON_FILE_PATH = path.join(DATA_DIR, "last-defect.json");
const TXT_FILE_PATH = path.join(DATA_DIR, "last-defect.txt");

export function loadRecordFromDisk(): DefectRecord {
  try {
    if (fs.existsSync(JSON_FILE_PATH)) {
      const raw = fs.readFileSync(JSON_FILE_PATH, "utf-8").trim();
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed.value === "string" && typeof parsed.timestamp === "number") {
        return parsed;
      }
    }
  } catch {}

  // Fallback to old text file if json doesn't exist yet
  try {
    if (fs.existsSync(TXT_FILE_PATH)) {
      const txt = fs.readFileSync(TXT_FILE_PATH, "utf-8").trim();
      const stat = fs.statSync(TXT_FILE_PATH);
      return { value: txt, timestamp: stat.mtimeMs || 1 };
    }
  } catch {}

  return { value: "", timestamp: 0 };
}

export function saveRecordToDisk(rec: DefectRecord): void {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
    fs.writeFileSync(JSON_FILE_PATH, JSON.stringify(rec), "utf-8");
    fs.writeFileSync(TXT_FILE_PATH, rec.value, "utf-8");
  } catch {}
}

const listeners = new Set<(rec: DefectRecord) => void>();

export function addDefectListener(fn: (rec: DefectRecord) => void) {
  listeners.add(fn);
}

export function removeDefectListener(fn: (rec: DefectRecord) => void) {
  listeners.delete(fn);
}

export function notifyDefectListeners(rec: DefectRecord) {
  listeners.forEach((fn) => {
    try {
      fn(rec);
    } catch {}
  });
}

let serverRecord = loadRecordFromDisk();

export function getLatestRecord(): DefectRecord {
  const diskRec = loadRecordFromDisk();
  if (diskRec.timestamp > serverRecord.timestamp) {
    serverRecord = diskRec;
  }
  return serverRecord;
}

export const Route = createFileRoute("/api/last-defect")({
  server: {
    handlers: {
      GET: async () => {
        const rec = getLatestRecord();
        return new Response(JSON.stringify(rec), {
          headers: {
            "Content-Type": "application/json",
            "Cache-Control": "no-store, no-cache, must-revalidate",
          },
        });
      },
      POST: async ({ request }) => {
        try {
          const body = await request.json();
          if (body && typeof body.value === "string") {
            const val = body.value.trim();
            const ts = typeof body.timestamp === "number" ? body.timestamp : Date.now();
            if (ts >= serverRecord.timestamp) {
              serverRecord = { value: val, timestamp: ts };
              saveRecordToDisk(serverRecord);
              notifyDefectListeners(serverRecord);
            }
          }
        } catch {}
        return new Response(JSON.stringify(serverRecord), {
          headers: {
            "Content-Type": "application/json",
            "Cache-Control": "no-store, no-cache, must-revalidate",
          },
        });
      },
    },
  },
});
