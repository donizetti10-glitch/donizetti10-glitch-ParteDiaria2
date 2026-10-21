import { createFileRoute } from "@tanstack/react-router";
import fs from "node:fs";
import path from "node:path";
import type { EquipmentChecklist } from "@/lib/checklist";

const DATA_DIR = path.join(process.cwd(), "data");
const JSON_FILE_PATH = path.join(DATA_DIR, "checklists.json");

export function loadChecklistsFromDisk(): EquipmentChecklist[] {
  try {
    if (fs.existsSync(JSON_FILE_PATH)) {
      const raw = fs.readFileSync(JSON_FILE_PATH, "utf-8").trim();
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          return parsed;
        }
      }
    }
  } catch (err) {
    console.warn("[Server Checklists] Error reading checklists from disk:", err);
  }
  return [];
}

export function saveChecklistsToDisk(checklists: EquipmentChecklist[]): void {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
    fs.writeFileSync(JSON_FILE_PATH, JSON.stringify(checklists, null, 2), "utf-8");
  } catch (err) {
    console.warn("[Server Checklists] Error saving checklists to disk:", err);
  }
}

const checklistListeners = new Set<(checklists: EquipmentChecklist[]) => void>();

export function addChecklistsListener(fn: (checklists: EquipmentChecklist[]) => void) {
  checklistListeners.add(fn);
}

export function removeChecklistsListener(fn: (checklists: EquipmentChecklist[]) => void) {
  checklistListeners.delete(fn);
}

export function notifyChecklistsListeners(checklists: EquipmentChecklist[]) {
  checklistListeners.forEach((fn) => {
    try {
      fn(checklists);
    } catch (e) {
      console.warn("Checklist listener error:", e);
    }
  });
}

let memoryChecklists: EquipmentChecklist[] = loadChecklistsFromDisk();

export function getLatestServerChecklists(): EquipmentChecklist[] {
  const disk = loadChecklistsFromDisk();
  if (disk.length >= memoryChecklists.length) {
    memoryChecklists = disk;
  }
  return memoryChecklists;
}

export const Route = createFileRoute("/api/checklists")({
  server: {
    handlers: {
      GET: async () => {
        const current = getLatestServerChecklists();
        return new Response(JSON.stringify({ checklists: current }), {
          status: 200,
          headers: {
            "Content-Type": "application/json",
            "Cache-Control": "no-store, no-cache, must-revalidate",
          },
        });
      },
      POST: async ({ request }: { request: Request }) => {
        try {
          const body = await request.json();
          let incoming: EquipmentChecklist[] = [];
          if (Array.isArray(body?.checklists)) {
            incoming = body.checklists;
          } else if (body?.checklist && body.checklist.id) {
            incoming = [body.checklist];
          }

          if (incoming.length === 0) {
            return new Response(JSON.stringify({ error: "No checklists provided" }), {
              status: 400,
              headers: { "Content-Type": "application/json" },
            });
          }

          const current = getLatestServerChecklists();
          const map = new Map<string, EquipmentChecklist>();
          for (const c of current) {
            if (c?.id) map.set(c.id, c);
          }
          for (const inc of incoming) {
            if (inc?.id) {
              const existing = map.get(inc.id);
              if (!existing) {
                map.set(inc.id, inc);
              } else {
                const timeInc = inc.updatedAt || inc.createdAt || 0;
                const timeExist = existing.updatedAt || existing.createdAt || 0;
                map.set(inc.id, timeInc >= timeExist ? inc : existing);
              }
            }
          }

          const updated = Array.from(map.values()).sort(
            (a, b) => (b.createdAt || 0) - (a.createdAt || 0),
          );

          memoryChecklists = updated;
          saveChecklistsToDisk(updated);
          notifyChecklistsListeners(updated);

          return new Response(JSON.stringify({ success: true, count: updated.length }), {
            status: 200,
            headers: { "Content-Type": "application/json" },
          });
        } catch (err) {
          return new Response(JSON.stringify({ error: String(err) }), {
            status: 500,
            headers: { "Content-Type": "application/json" },
          });
        }
      },
      DELETE: async ({ request }: { request: Request }) => {
        try {
          const url = new URL(request.url);
          const id = url.searchParams.get("id");
          if (!id) {
            return new Response(JSON.stringify({ error: "Missing id" }), {
              status: 400,
              headers: { "Content-Type": "application/json" },
            });
          }

          const current = getLatestServerChecklists();
          const filtered = current.filter((c) => c.id !== id);
          memoryChecklists = filtered;
          saveChecklistsToDisk(filtered);
          notifyChecklistsListeners(filtered);

          return new Response(JSON.stringify({ success: true, id }), {
            status: 200,
            headers: { "Content-Type": "application/json" },
          });
        } catch (err) {
          return new Response(JSON.stringify({ error: String(err) }), {
            status: 500,
            headers: { "Content-Type": "application/json" },
          });
        }
      },
    },
  },
});
