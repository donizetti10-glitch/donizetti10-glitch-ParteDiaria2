import { createFileRoute } from "@tanstack/react-router";
import fs from "node:fs";
import path from "node:path";
import type { SavedCloudReport } from "@/lib/cloud-reports";

const DATA_DIR = path.join(process.cwd(), "data");
const JSON_FILE_PATH = path.join(DATA_DIR, "cloud-reports.json");

export function deduplicateServerReports(reports: SavedCloudReport[]): SavedCloudReport[] {
  if (!Array.isArray(reports)) return [];
  const byDate = new Map<string, SavedCloudReport[]>();

  for (const r of reports) {
    if (!r || !r.id) continue;
    const dateKey = r.data ? r.data.trim() : "SEM_DATA";
    if (!byDate.has(dateKey)) {
      byDate.set(dateKey, []);
    }
    byDate.get(dateKey)!.push(r);
  }

  const unique: SavedCloudReport[] = [];
  for (const list of byDate.values()) {
    list.sort((a, b) => {
      const timeA = a.updatedAt || a.createdAt || 0;
      const timeB = b.updatedAt || b.createdAt || 0;
      if (timeB !== timeA) return timeB - timeA;
      const eventsA = Array.isArray(a.events) ? a.events.length : 0;
      const eventsB = Array.isArray(b.events) ? b.events.length : 0;
      return eventsB - eventsA;
    });
    unique.push(list[0]);
  }

  unique.sort((a, b) => {
    const da = a.data || "";
    const db = b.data || "";
    if (db !== da) return db.localeCompare(da);
    return (b.updatedAt || b.createdAt || 0) - (a.updatedAt || a.createdAt || 0);
  });

  return unique;
}

export function loadReportsFromDisk(): SavedCloudReport[] {
  try {
    if (fs.existsSync(JSON_FILE_PATH)) {
      const raw = fs.readFileSync(JSON_FILE_PATH, "utf-8").trim();
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          return deduplicateServerReports(parsed);
        }
      }
    }
  } catch (err) {
    console.warn("[Server Cloud Reports] Error reading reports from disk:", err);
  }
  return [];
}

export function saveReportsToDisk(reports: SavedCloudReport[]): void {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
    const deduplicated = deduplicateServerReports(reports);
    fs.writeFileSync(JSON_FILE_PATH, JSON.stringify(deduplicated, null, 2), "utf-8");
  } catch (err) {
    console.warn("[Server Cloud Reports] Error saving reports to disk:", err);
  }
}

const reportListeners = new Set<(reports: SavedCloudReport[]) => void>();

export function addReportsListener(fn: (reports: SavedCloudReport[]) => void) {
  reportListeners.add(fn);
}

export function removeReportsListener(fn: (reports: SavedCloudReport[]) => void) {
  reportListeners.delete(fn);
}

export function notifyReportsListeners(reports: SavedCloudReport[]) {
  reportListeners.forEach((fn) => {
    try {
      fn(reports);
    } catch (e) {
      console.warn("Listener error:", e);
    }
  });
}

let memoryReports: SavedCloudReport[] = loadReportsFromDisk();

export function getLatestServerReports(): SavedCloudReport[] {
  const disk = loadReportsFromDisk();
  if (disk.length >= memoryReports.length) {
    memoryReports = disk;
  }
  return memoryReports;
}

export function upsertServerReports(newReports: SavedCloudReport[]): SavedCloudReport[] {
  const existingMap = new Map<string, SavedCloudReport>();
  for (const r of memoryReports) {
    if (r && r.id) existingMap.set(r.id, r);
  }

  for (const r of newReports) {
    if (!r || !r.id) continue;
    const existing = existingMap.get(r.id);
    if (!existing) {
      existingMap.set(r.id, r);
    } else {
      const tOld = existing.updatedAt || existing.createdAt || 0;
      const tNew = r.updatedAt || r.createdAt || 0;
      if (tNew >= tOld) {
        existingMap.set(r.id, r);
      }
    }
  }

  const merged = Array.from(existingMap.values());
  const deduplicated = deduplicateServerReports(merged);
  memoryReports = deduplicated;
  saveReportsToDisk(deduplicated);
  notifyReportsListeners(deduplicated);
  return deduplicated;
}

export function removeServerReport(id: string): SavedCloudReport[] {
  memoryReports = memoryReports.filter((r) => r.id !== id);
  saveReportsToDisk(memoryReports);
  notifyReportsListeners(memoryReports);
  return memoryReports;
}

export const Route = createFileRoute("/api/cloud-reports")({
  server: {
    handlers: {
      GET: async () => {
        const reports = getLatestServerReports();
        return new Response(JSON.stringify({ ok: true, reports }), {
          headers: {
            "Content-Type": "application/json",
            "Cache-Control": "no-store, no-cache, must-revalidate",
          },
        });
      },
      POST: async ({ request }) => {
        try {
          const body = await request.json();
          let incoming: SavedCloudReport[] = [];

          if (Array.isArray(body)) {
            incoming = body;
          } else if (body && Array.isArray(body.reports)) {
            incoming = body.reports;
          } else if (body && body.report && body.report.id) {
            incoming = [body.report];
          } else if (body && body.id) {
            incoming = [body as SavedCloudReport];
          }

          if (incoming.length > 0) {
            const updated = upsertServerReports(incoming);
            return new Response(
              JSON.stringify({ ok: true, count: updated.length, reports: updated }),
              {
                headers: {
                  "Content-Type": "application/json",
                  "Cache-Control": "no-store, no-cache, must-revalidate",
                },
              },
            );
          }
        } catch (err) {
          console.warn("[Server Cloud Reports POST error]:", err);
        }

        return new Response(
          JSON.stringify({ ok: true, count: memoryReports.length, reports: memoryReports }),
          {
            headers: {
              "Content-Type": "application/json",
              "Cache-Control": "no-store, no-cache, must-revalidate",
            },
          },
        );
      },
      DELETE: async ({ request }) => {
        try {
          const url = new URL(request.url);
          let targetId = url.searchParams.get("id");
          if (!targetId) {
            const body = await request.json().catch(() => null);
            targetId = body?.id;
          }

          if (targetId) {
            removeServerReport(targetId);
          }
        } catch (err) {
          console.warn("[Server Cloud Reports DELETE error]:", err);
        }

        return new Response(JSON.stringify({ ok: true }), {
          headers: {
            "Content-Type": "application/json",
            "Cache-Control": "no-store, no-cache, must-revalidate",
          },
        });
      },
    },
  },
});
