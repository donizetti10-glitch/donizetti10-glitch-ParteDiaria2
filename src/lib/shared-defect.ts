import { supabase } from "@/integrations/supabase/client";
import { db, doc, getDoc, setDoc, onSnapshot } from "./firebase";

export interface DefectRecord {
  value: string;
  timestamp: number;
}

const KEY = "last_defect_rec";
const LEGACY_KEY = "last_defect";
const CHANNEL_NAME = "last_defect_bc_v3";

let currentRecord: DefectRecord = { value: "", timestamp: 0 };

// Load initial value from localStorage synchronously if available
if (typeof localStorage !== "undefined") {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed.value === "string" && typeof parsed.timestamp === "number") {
        currentRecord = parsed;
      }
    } else {
      const legacy = localStorage.getItem(LEGACY_KEY);
      if (legacy) {
        currentRecord = { value: legacy, timestamp: 1 };
      }
    }
  } catch {}
}

const localSubscribers = new Set<(value: string, record: DefectRecord) => void>();

// Helper for BroadcastChannel cross-tab communication
function getBroadcastChannel(): BroadcastChannel | null {
  if (typeof window !== "undefined" && "BroadcastChannel" in window) {
    try {
      return new BroadcastChannel(CHANNEL_NAME);
    } catch {}
  }
  return null;
}

export function updateIfNewer(incoming: DefectRecord): boolean {
  if (!incoming || typeof incoming.value !== "string" || typeof incoming.timestamp !== "number") {
    return false;
  }
  if (incoming.timestamp > currentRecord.timestamp) {
    currentRecord = { value: incoming.value, timestamp: incoming.timestamp };
    if (typeof localStorage !== "undefined") {
      try {
        localStorage.setItem(KEY, JSON.stringify(currentRecord));
        localStorage.setItem(LEGACY_KEY, currentRecord.value);
      } catch {}
    }
    localSubscribers.forEach((fn) => {
      try {
        fn(currentRecord.value, currentRecord);
      } catch {}
    });
    return true;
  }
  return false;
}

export async function fetchLastDefect(): Promise<string> {
  // 1. Fetch from Firestore
  try {
    const docRef = doc(db, "shared_state", "last_defect");
    const snap = await getDoc(docRef);
    if (snap.exists()) {
      const d = snap.data();
      if (d && typeof d.value === "string" && typeof d.timestamp === "number") {
        updateIfNewer({ value: d.value, timestamp: d.timestamp });
      }
    }
  } catch {}

  // 2. Fetch from server API endpoint
  try {
    const res = await fetch("/api/last-defect", { cache: "no-store" });
    if (res.ok) {
      const data = await res.json();
      if (data && typeof data.value === "string" && typeof data.timestamp === "number") {
        updateIfNewer(data);
      }
    }
  } catch {}

  // 3. Fetch from Supabase
  try {
    const { data } = await supabase
      .from("shared_state")
      .select("value, updated_at")
      .eq("key", "last_defect")
      .maybeSingle();
    if (data && typeof data.value === "string") {
      const ts = data.updated_at ? new Date(data.updated_at).getTime() : 1;
      updateIfNewer({ value: data.value, timestamp: ts || 1 });
    }
  } catch {}

  return currentRecord.value;
}

export async function saveLastDefect(value: string): Promise<void> {
  const newRec: DefectRecord = {
    value: value.trim(),
    timestamp: Date.now(),
  };

  // Update locally instantly with fresh timestamp
  updateIfNewer(newRec);

  // Notify other tabs via BroadcastChannel
  const bc = getBroadcastChannel();
  if (bc) {
    try {
      bc.postMessage({ type: "LAST_DEFECT_UPDATED", record: newRec });
      bc.close();
    } catch {}
  }

  // 1. Save to Firestore (primary cloud)
  try {
    const docRef = doc(db, "shared_state", "last_defect");
    await setDoc(docRef, { ...newRec, key: "last_defect" }, { merge: true });
  } catch (err) {
    console.warn("Firestore last_defect save error:", err);
  }

  // 2. Send to server API endpoint
  try {
    const res = await fetch("/api/last-defect", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(newRec),
    });
    if (res.ok) {
      const data = await res.json();
      if (data && typeof data.value === "string" && typeof data.timestamp === "number") {
        updateIfNewer(data);
      }
    }
  } catch (err) {
    console.warn("API save warning:", err);
  }

  // 3. Save to Supabase Cloud Database for backup
  try {
    await supabase.from("shared_state").upsert(
      {
        key: "last_defect",
        value: newRec.value,
        updated_at: new Date(newRec.timestamp).toISOString(),
      },
      { onConflict: "key" },
    );
  } catch (err) {
    console.warn("Supabase save warning:", err);
  }
}

export function subscribeLastDefect(onChange: (value: string) => void) {
  // Immediately emit current record value to subscriber
  onChange(currentRecord.value);

  const subscriber = (val: string) => {
    onChange(val);
  };
  localSubscribers.add(subscriber);

  // 1. Listen to BroadcastChannel for instant local cross-tab updates
  const bc = getBroadcastChannel();
  if (bc) {
    bc.onmessage = (event) => {
      if (event.data?.type === "LAST_DEFECT_UPDATED" && event.data?.record) {
        updateIfNewer(event.data.record);
      }
    };
  }

  // 2. Listen to localStorage storage events
  const handleStorage = (e: StorageEvent) => {
    if (e.key === KEY && typeof e.newValue === "string") {
      try {
        const parsed = JSON.parse(e.newValue);
        if (parsed && typeof parsed.value === "string" && typeof parsed.timestamp === "number") {
          updateIfNewer(parsed);
        }
      } catch {}
    }
  };
  if (typeof window !== "undefined") {
    window.addEventListener("storage", handleStorage);
  }

  // 3. SSE (Server-Sent Events) connection for instant push updates (<50ms) across all devices
  let eventSource: EventSource | null = null;
  if (typeof window !== "undefined" && "EventSource" in window) {
    try {
      eventSource = new EventSource("/api/last-defect-stream");
      eventSource.onmessage = (event) => {
        try {
          const data = JSON.parse(event.data);
          if (data && typeof data.value === "string" && typeof data.timestamp === "number") {
            updateIfNewer(data);
          }
        } catch {}
      };
    } catch {}
  }

  // 4. Firestore Realtime listener (multi-device instantaneous sync)
  let unsubscribeFirestore: (() => void) | null = null;
  try {
    const docRef = doc(db, "shared_state", "last_defect");
    unsubscribeFirestore = onSnapshot(docRef, (snap) => {
      if (snap.exists()) {
        const d = snap.data();
        if (d && typeof d.value === "string" && typeof d.timestamp === "number") {
          updateIfNewer({ value: d.value, timestamp: d.timestamp });
        }
      }
    });
  } catch {}

  // 5. Listen to Supabase Realtime changes
  let supabaseChannel: ReturnType<typeof supabase.channel> | null = null;
  try {
    supabaseChannel = supabase
      .channel("shared_state_last_defect_v3")
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "shared_state", filter: "key=eq.last_defect" },
        (payload) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const newRow: any = (payload as any).new;
          if (newRow && typeof newRow.value === "string") {
            let ts = typeof newRow.timestamp === "number" ? newRow.timestamp : 0;
            if (!ts && newRow.updated_at) {
              ts = new Date(newRow.updated_at).getTime();
            }
            updateIfNewer({ value: newRow.value, timestamp: ts || Date.now() });
          }
        },
      )
      .subscribe();
  } catch {}

  // 6. Periodic polling interval (every 2s) to guarantee real-time sync across lossy mobile networks
  const intervalId = setInterval(() => {
    fetchLastDefect().catch(() => {});
  }, 2000);

  return () => {
    localSubscribers.delete(subscriber);
    if (unsubscribeFirestore) {
      try {
        unsubscribeFirestore();
      } catch {}
    }
    if (bc) {
      try {
        bc.close();
      } catch {}
    }
    if (typeof window !== "undefined") {
      window.removeEventListener("storage", handleStorage);
    }
    if (eventSource) {
      try {
        eventSource.close();
      } catch {}
    }
    if (supabaseChannel) {
      try {
        supabase.removeChannel(supabaseChannel);
      } catch {}
    }
    clearInterval(intervalId);
  };
}
