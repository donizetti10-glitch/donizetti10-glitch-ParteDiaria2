import { supabase } from "@/integrations/supabase/client";
import { db, collection, doc, getDocs, setDoc, deleteDoc, onSnapshot } from "./firebase";
import {
  saveReportsToIndexedDB,
  getReportsFromIndexedDB,
  deleteReportFromIndexedDB,
} from "./indexed-db";
import {
  type DailyHeader,
  type WorkEvent,
  type ReportPhoto,
  totals,
  emptyHeader,
  sortAndNormalizeEvents,
} from "./parte-diaria";

export interface SavedCloudReport {
  id: string;
  createdAt: number;
  updatedAt: number;
  data: string;
  nome: string;
  local: string;
  sinais: string;
  linha: string;
  via: string;
  gda: string;
  om: string;
  empresa: string;
  atividade: string;
  colaboradores: string;
  descritivo: string;
  kmProgramado: string;
  executado: string;
  observacaoGeral: string;
  telefone: string;
  events: WorkEvent[];
  photos?: ReportPhoto[];
  totals: {
    jornada: number;
    atividade: number;
    preparacao: number;
    parada: number;
    producaoTotal: number;
    defeitosTotal: number;
  };
}

export function mapFirestoreDocToReport(
  id: string,
  row: Record<string, unknown>,
): SavedCloudReport {
  return {
    id: id || String(row.id || ""),
    createdAt: Number(
      row.createdAt || (row.created_at ? new Date(String(row.created_at)).getTime() : Date.now()),
    ),
    updatedAt: Number(
      row.updatedAt || (row.updated_at ? new Date(String(row.updated_at)).getTime() : Date.now()),
    ),
    data: String(row.data || ""),
    nome: String(row.nome || ""),
    local: String(row.local || ""),
    sinais: String(row.sinais || ""),
    linha: String(row.linha || ""),
    via: String(row.via || ""),
    gda: String(row.gda || ""),
    om: String(row.om || ""),
    empresa: String(row.empresa || ""),
    atividade: String(row.atividade || ""),
    colaboradores: String(row.colaboradores || ""),
    descritivo: String(row.descritivo || ""),
    kmProgramado: String(row.kmProgramado || row.km_programado || ""),
    executado: String(row.executado || ""),
    observacaoGeral: String(row.observacaoGeral || row.observacao_geral || ""),
    telefone: String(row.telefone || ""),
    events: Array.isArray(row.events) ? sortAndNormalizeEvents(row.events as WorkEvent[]) : [],
    photos: Array.isArray(row.photos) ? (row.photos as ReportPhoto[]) : [],
    totals:
      row.totals && typeof row.totals === "object"
        ? (row.totals as SavedCloudReport["totals"])
        : {
            jornada: 0,
            atividade: 0,
            preparacao: 0,
            parada: 0,
            producaoTotal: 0,
            defeitosTotal: 0,
          },
  };
}

export function mapSupabaseRowToReport(row: Record<string, unknown>): SavedCloudReport {
  return mapFirestoreDocToReport(String(row.id || ""), row);
}

const COLLECTION_NAME = "relatorios";

export async function fetchFirestoreReports(): Promise<SavedCloudReport[]> {
  try {
    const colRef = collection(db, COLLECTION_NAME);
    const snap = await getDocs(colRef);
    if (!snap.empty) {
      const reports: SavedCloudReport[] = [];
      snap.forEach((docSnap) => {
        reports.push(mapFirestoreDocToReport(docSnap.id, docSnap.data()));
      });
      unblacklistActiveReports(reports);
      const { unique } = deduplicateReportsByDay(reports);
      return unique;
    }
  } catch (err) {
    console.warn("[Firestore] Error fetching reports:", err);
  }
  return [];
}

export async function fetchSupabaseReports(): Promise<SavedCloudReport[]> {
  // Dá prioridade primária ao Firestore oficial do projeto
  const fsReports = await fetchFirestoreReports();
  if (fsReports && fsReports.length > 0) {
    return fsReports;
  }
  try {
    if (!supabase || typeof supabase.from !== "function") return [];
    const { data, error } = await supabase
      .from("relatorios")
      .select("*")
      .order("createdAt", { ascending: false });

    if (!error && Array.isArray(data) && data.length > 0) {
      return data.map(mapSupabaseRowToReport);
    }
  } catch (err) {
    console.warn("Supabase fetch notice:", err);
  }
  return [];
}
const CACHE_KEY = "cloud_reports_cache_v3";
const DELETED_IDS_KEY = "cloud_reports_deleted_ids_v2";

export function getDeletedIds(): Set<string> {
  if (typeof localStorage === "undefined") return new Set();
  try {
    const raw = localStorage.getItem(DELETED_IDS_KEY);
    return raw ? new Set(JSON.parse(raw)) : new Set();
  } catch {
    return new Set();
  }
}

export function addDeletedId(id: string) {
  if (typeof localStorage === "undefined" || !id) return;
  try {
    const ids = getDeletedIds();
    ids.add(id);
    localStorage.setItem(DELETED_IDS_KEY, JSON.stringify(Array.from(ids)));
  } catch (e) {
    console.warn("Error saving deleted ID:", e);
  }
}

export function removeDeletedId(id: string) {
  if (typeof localStorage === "undefined" || !id) return;
  try {
    const ids = getDeletedIds();
    if (ids.has(id)) {
      ids.delete(id);
      localStorage.setItem(DELETED_IDS_KEY, JSON.stringify(Array.from(ids)));
    }
  } catch (e) {
    console.warn("Error removing deleted ID:", e);
  }
}

/**
 * Remove do conjunto de IDs excluídos do dispositivo qualquer relatório ativo recebido do servidor.
 * Isso garante que celulares com caches antigos ou blacklists corrompidas recuperem imediatamente os 34 relatórios.
 */
export function unblacklistActiveReports(reports: SavedCloudReport[]) {
  if (typeof localStorage === "undefined" || !Array.isArray(reports) || reports.length === 0) {
    return;
  }
  try {
    const ids = getDeletedIds();
    let changed = false;
    for (const r of reports) {
      if (r && r.id && ids.has(r.id)) {
        ids.delete(r.id);
        changed = true;
      }
    }
    if (changed) {
      localStorage.setItem(DELETED_IDS_KEY, JSON.stringify(Array.from(ids)));
    }
  } catch (e) {
    console.warn("Notice unblacklisting active reports:", e);
  }
}

export function sanitizeForFirestore<T>(obj: T): T {
  if (obj === null || obj === undefined) {
    return null as unknown as T;
  }
  if (typeof obj !== "object") {
    return obj;
  }
  if (Array.isArray(obj)) {
    return obj.map((item) => sanitizeForFirestore(item)) as unknown as T;
  }
  const cleaned: Record<string, unknown> = {};
  for (const [key, val] of Object.entries(obj as Record<string, unknown>)) {
    if (val !== undefined) {
      cleaned[key] = sanitizeForFirestore(val);
    }
  }
  return cleaned as T;
}

export function deduplicateReportsByDay(reports: SavedCloudReport[]): {
  unique: SavedCloudReport[];
  duplicateIds: string[];
} {
  const deletedSet = getDeletedIds();
  const byDate = new Map<string, SavedCloudReport[]>();

  for (const r of reports) {
    if (!r || !r.id || deletedSet.has(r.id)) continue;
    const dateKey = r.data?.trim() || "SEM_DATA";
    if (!byDate.has(dateKey)) {
      byDate.set(dateKey, []);
    }
    byDate.get(dateKey)!.push(r);
  }

  const unique: SavedCloudReport[] = [];
  const duplicateIds: string[] = [];

  for (const list of byDate.values()) {
    // Ordena do mais recente para o mais antigo:
    // 1) Maior timestamp (updatedAt ou createdAt)
    // 2) Maior número de eventos (para desempatar)
    list.sort((a, b) => {
      const timeA = a.updatedAt || a.createdAt || 0;
      const timeB = b.updatedAt || b.createdAt || 0;
      if (timeB !== timeA) return timeB - timeA;
      const eventsA = a.events?.length || 0;
      const eventsB = b.events?.length || 0;
      return eventsB - eventsA;
    });

    const winner = list[0];
    unique.push(winner);

    for (let i = 1; i < list.length; i++) {
      duplicateIds.push(list[i].id);
    }
  }

  // Ordena lista única por data decrescente (ou createdAt)
  unique.sort((a, b) => {
    if (a.data && b.data && a.data !== b.data) {
      return b.data.localeCompare(a.data);
    }
    return (b.createdAt || 0) - (a.createdAt || 0);
  });

  return { unique, duplicateIds };
}

export function mergeReports(
  remote: SavedCloudReport[],
  local: SavedCloudReport[],
): SavedCloudReport[] {
  const deletedSet = getDeletedIds();
  const map = new Map<string, SavedCloudReport>();
  for (const r of remote) {
    if (r && r.id && !deletedSet.has(r.id)) map.set(r.id, r);
  }
  for (const l of local) {
    if (l && l.id && !deletedSet.has(l.id)) {
      const existing = map.get(l.id);
      if (!existing) {
        map.set(l.id, l);
      } else {
        const timeLocal = l.updatedAt || l.createdAt || 0;
        const timeExisting = existing.updatedAt || existing.createdAt || 0;
        const base = timeLocal >= timeExisting ? l : existing;
        // Preserva fotos de qualquer uma das cópias se disponível
        const photos =
          l.photos && l.photos.length > 0 && l.photos.some((p) => Boolean(p.dataUrl))
            ? l.photos
            : existing.photos && existing.photos.length > 0
              ? existing.photos
              : l.photos || [];
        map.set(l.id, { ...base, photos });
      }
    }
  }

  // Aplica a regra de 1 relatório por dia mantendo o mais completo
  const { unique } = deduplicateReportsByDay(Array.from(map.values()));
  return unique;
}

// Cache em memória para renderização e cálculo instantâneo sem perda de fotos
let memoryReportsCache: SavedCloudReport[] | null = null;

export function getLocalCache(): SavedCloudReport[] {
  if (memoryReportsCache && memoryReportsCache.length > 0) {
    return memoryReportsCache;
  }

  if (typeof localStorage === "undefined") return [];
  try {
    const deletedSet = getDeletedIds();
    let raw = localStorage.getItem(CACHE_KEY);

    // Fallback em cascata para versões anteriores do cache para nunca perder dados
    if (!raw || raw === "[]") {
      const candidates = [
        "cloud_reports_cache_v2",
        "cloud_reports_cache_v1",
        "cloud_reports_cache",
      ];
      for (const cand of candidates) {
        const item = localStorage.getItem(cand);
        if (item && item !== "[]") {
          raw = item;
          try {
            localStorage.setItem(CACHE_KEY, item);
          } catch (e) {
            console.warn("Notice updating cache key:", e);
          }
          break;
        }
      }
    }

    const reports: SavedCloudReport[] = raw ? JSON.parse(raw) : [];
    const validReports = reports
      .filter((r) => r && r.id && !deletedSet.has(r.id))
      .map((r) => ({
        ...r,
        events: sortAndNormalizeEvents(r.events ?? []),
      }));

    const { unique } = deduplicateReportsByDay(validReports);
    memoryReportsCache = unique;

    // Hidrata fotos do IndexedDB em segundo plano se disponível
    if (typeof window !== "undefined") {
      getReportsFromIndexedDB()
        .then((idbReports) => {
          if (idbReports && idbReports.length > 0) {
            const merged = mergeReports(idbReports, memoryReportsCache || unique);
            memoryReportsCache = merged;
          }
        })
        .catch(() => {});
    }

    return unique;
  } catch {
    return [];
  }
}

export function setLocalCache(reports: SavedCloudReport[]) {
  const { unique } = deduplicateReportsByDay(reports);
  memoryReportsCache = unique;

  // 1. Salva a versão completa (com todas as fotos) no IndexedDB (sem limite de 5MB)
  saveReportsToIndexedDB(unique).catch((err) =>
    console.warn("[Cache] Error saving full reports to IndexedDB:", err),
  );

  // 2. Salva no localStorage para acesso síncrono instantâneo
  if (typeof localStorage === "undefined") return;
  try {
    const deletedSet = getDeletedIds();
    const filtered = unique.filter((r) => r && r.id && !deletedSet.has(r.id));
    localStorage.setItem(CACHE_KEY, JSON.stringify(filtered));
  } catch (e) {
    console.warn(
      "Notice: localStorage full with photos. Saving lightweight cache with ALL reports preserved:",
      e,
    );
    try {
      // IMPORTANTE: NUNCA corta ou descarta relatórios (sem slice).
      // Apenas retira os dados binários pesados de fotos do localStorage, pois ficam salvos no IndexedDB e no servidor.
      const deletedSet = getDeletedIds();
      const lightweight = unique
        .filter((r) => r && r.id && !deletedSet.has(r.id))
        .map((r) => ({
          ...r,
          photos: Array.isArray(r.photos)
            ? r.photos.map((p) => ({ id: p.id, caption: p.caption, dataUrl: "" }))
            : [],
        }));
      localStorage.setItem(CACHE_KEY, JSON.stringify(lightweight));
    } catch (finalErr) {
      console.warn("Could not save to localStorage fallback:", finalErr);
    }
  }
}

export async function fetchServerReports(): Promise<SavedCloudReport[]> {
  if (typeof window === "undefined") return [];
  try {
    const res = await fetch("/api/cloud-reports", {
      headers: { Accept: "application/json" },
      cache: "no-store",
    });
    if (!res.ok) return [];
    const data = await res.json();
    if (data && Array.isArray(data.reports)) {
      unblacklistActiveReports(data.reports);
      return data.reports;
    }
  } catch (e) {
    console.warn("[Cloud Reports] Fetch from server notice:", e);
  }
  return [];
}

export async function pushReportsToServer(reports: SavedCloudReport[]): Promise<void> {
  if (typeof window === "undefined" || !Array.isArray(reports) || reports.length === 0) return;
  try {
    await fetch("/api/cloud-reports", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reports }),
    });
  } catch (e) {
    console.warn("[Cloud Reports] Push to server notice:", e);
  }
}

export async function deleteServerReport(id: string): Promise<void> {
  if (typeof window === "undefined" || !id) return;
  try {
    await fetch(`/api/cloud-reports?id=${encodeURIComponent(id)}`, {
      method: "DELETE",
    });
  } catch (e) {
    console.warn("[Cloud Reports] Delete from server notice:", e);
  }
}

export async function syncCloudReportsAcrossDevices(): Promise<{
  reports: SavedCloudReport[];
  source: "server" | "local" | "supabase" | "firestore";
}> {
  // 1. Busca do Firestore oficial em nuvem
  let firestoreReports: SavedCloudReport[] = [];
  try {
    firestoreReports = await fetchFirestoreReports();
    if (firestoreReports && firestoreReports.length > 0) {
      unblacklistActiveReports(firestoreReports);
    }
  } catch (e) {
    console.warn("Notice fetching Firestore reports in sync:", e);
  }

  // 2. Busca do servidor de apoio
  let serverReports: SavedCloudReport[] = [];
  try {
    serverReports = await fetchServerReports();
    if (serverReports && serverReports.length > 0) {
      unblacklistActiveReports(serverReports);
    }
  } catch (e) {
    console.warn("Notice fetching initial server reports:", e);
  }

  const local = getLocalCache();
  const baseCloud = firestoreReports.length > 0 ? firestoreReports : serverReports;

  // Se a nuvem tem relatórios, mescla com o cache do dispositivo
  if (baseCloud.length > 0) {
    const merged = mergeReports(baseCloud, local);
    setLocalCache(merged);

    // Propaga relatórios locais novos para o Firestore e servidor
    for (const r of merged) {
      if (!firestoreReports.some((f) => f.id === r.id)) {
        try {
          const docRef = doc(db, COLLECTION_NAME, r.id);
          setDoc(docRef, sanitizeForFirestore(r), { merge: true }).catch(() => {});
        } catch {}
      }
    }
    pushReportsToServer(merged).catch(() => {});
    return {
      reports: merged,
      source: firestoreReports.length > 0 ? "firestore" : "server",
    };
  }

  // Se o servidor e Firestore estiverem vazios e este cliente tiver relatórios:
  if (local.length > 0) {
    for (const r of local) {
      try {
        const docRef = doc(db, COLLECTION_NAME, r.id);
        setDoc(docRef, sanitizeForFirestore(r), { merge: true }).catch(() => {});
      } catch {}
    }
    pushReportsToServer(local).catch(() => {});
    return { reports: local, source: "local" };
  }

  return { reports: local, source: "local" };
}

export async function saveReportToCloud(
  header: DailyHeader,
  events: WorkEvent[],
  existingId?: string,
  photos?: ReportPhoto[],
): Promise<SavedCloudReport> {
  const now = Date.now();
  const reportDate = header.data || new Date().toISOString().slice(0, 10);

  // Procura se já existe algum relatório com essa mesma data no cache local ou memória
  const cache = getLocalCache();
  const existingDayReport =
    cache.find(
      (r) =>
        r.data === reportDate && (!header.nome || !r.nome || r.nome.trim() === header.nome.trim()),
    ) || cache.find((r) => r.data === reportDate);

  // Reutiliza o ID do relatório do dia para evitar criar múltiplos arquivos para o mesmo dia
  const id =
    existingId ||
    existingDayReport?.id ||
    `REP-${reportDate.replace(/-/g, "")}-${Math.random()
      .toString(36)
      .substring(2, 7)
      .toUpperCase()}`;

  // Se restaurando ou salvando um ID, remove da blacklist de deletados
  removeDeletedId(id);

  // Normalização estrita dos eventos (sem lacunas/sobreposições)
  const normalizedEvents = sortAndNormalizeEvents(events ?? []);
  const t = totals(normalizedEvents);

  const reportData: SavedCloudReport = {
    id,
    createdAt: existingDayReport?.createdAt || now,
    updatedAt: now,
    data: reportDate,
    nome: header.nome || "",
    local: header.local || "",
    sinais: header.sinais || "",
    linha: header.linha || "",
    via: header.via || "",
    gda: header.gda || "",
    om: header.om || "",
    empresa: header.empresa || "",
    atividade: header.atividade || "",
    colaboradores: header.colaboradores || "",
    descritivo: header.descritivo || "",
    kmProgramado: header.kmProgramado || "",
    executado: header.executado || "",
    observacaoGeral: header.observacaoGeral || "",
    telefone: header.telefone || "",
    events: normalizedEvents,
    photos: photos || [],
    totals: {
      jornada: t.jornada,
      atividade: t.atividade,
      preparacao: t.preparacao,
      parada: t.parada,
      producaoTotal: t.producaoTotal,
      defeitosTotal: t.defeitosTotal,
    },
  };

  // 1. Remove da nuvem e do cache qualquer outro relatório duplicado com a mesma data
  const olderDuplicates = cache.filter((r) => r.data === reportDate && r.id !== id);
  for (const dup of olderDuplicates) {
    addDeletedId(dup.id);
    try {
      const docRef = doc(db, COLLECTION_NAME, dup.id);
      deleteDoc(docRef).catch(() => {});
    } catch {}
    if (supabase && typeof supabase.from === "function") {
      void supabase
        .from("relatorios")
        .delete()
        .eq("id", dup.id)
        .then(null, () => {});
    }
  }

  // 2. Salva no Firestore como banco oficial em nuvem
  try {
    const docRef = doc(db, COLLECTION_NAME, id);
    const sanitized = sanitizeForFirestore(reportData);
    await setDoc(docRef, sanitized, { merge: true });
  } catch (fsErr) {
    console.warn("[Firestore] Error saving report:", fsErr);
  }

  // 3. Salva no Supabase se configurado
  try {
    if (supabase && typeof supabase.from === "function") {
      await supabase.from("relatorios").upsert(
        {
          id,
          createdAt: Number(reportData.createdAt),
          updatedAt: Number(now),
          created_at: new Date(reportData.createdAt).toISOString(),
          updated_at: new Date(now).toISOString(),
          data: reportData.data,
          nome: reportData.nome,
          local: reportData.local,
          sinais: reportData.sinais,
          linha: reportData.linha,
          via: reportData.via,
          gda: reportData.gda,
          om: reportData.om,
          empresa: reportData.empresa,
          atividade: reportData.atividade,
          colaboradores: reportData.colaboradores,
          descritivo: reportData.descritivo,
          kmProgramado: reportData.kmProgramado,
          executado: reportData.executado,
          observacaoGeral: reportData.observacaoGeral,
          telefone: reportData.telefone,
          events: normalizedEvents,
          photos: photos || reportData.photos || [],
          totals: reportData.totals,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
        } as any,
        { onConflict: "id" },
      );
    }
  } catch (err) {
    console.warn("Supabase save notice:", err);
  }

  // 4. Atualiza o cache local (eliminando qualquer outro do mesmo dia)
  const remainingCache = cache.filter((r) => r.id !== id && r.data !== reportDate);
  const updatedCache = [reportData, ...remainingCache];
  setLocalCache(updatedCache);

  // 5. Sincroniza imediatamente com o servidor da aplicação
  pushReportsToServer([reportData]).catch((err) =>
    console.warn("Notice syncing report to server:", err),
  );
  for (const dup of olderDuplicates) {
    deleteServerReport(dup.id).catch(() => {});
  }

  return reportData;
}

/**
 * Varre os relatórios locais, do servidor e da nuvem, eliminando relatórios repetidos para o mesmo dia,
 * preservando estritamente apenas o último sincronizado de cada data.
 */
export async function deduplicateCloudReports(): Promise<{
  totalChecked: number;
  duplicatesRemoved: number;
}> {
  try {
    const current = getLocalCache();
    const { unique, duplicateIds } = deduplicateReportsByDay(current);

    if (duplicateIds.length > 0) {
      for (const dupId of duplicateIds) {
        addDeletedId(dupId);
        try {
          const docRef = doc(db, COLLECTION_NAME, dupId);
          deleteDoc(docRef).catch(() => {});
          deleteServerReport(dupId).catch(() => {});
          if (supabase && typeof supabase.from === "function") {
            await supabase.from("relatorios").delete().eq("id", dupId);
          }
        } catch (err) {
          console.warn("Deduplicate cleanup error:", err);
        }
      }
    }

    setLocalCache(unique);
    pushReportsToServer(unique).catch(() => {});
    return {
      totalChecked: current.length,
      duplicatesRemoved: duplicateIds.length,
    };
  } catch (err) {
    console.warn("Erro em deduplicateCloudReports:", err);
    return { totalChecked: 0, duplicatesRemoved: 0 };
  }
}

export async function fetchCloudReports(): Promise<SavedCloudReport[]> {
  const deletedSet = getDeletedIds();

  // 1. Tenta buscar diretamente do Firestore (banco oficial em nuvem do projeto)
  try {
    const fsReports = await fetchFirestoreReports();
    if (fsReports && fsReports.length > 0) {
      const valid = fsReports.filter((r) => r && r.id && !deletedSet.has(r.id));
      const { unique } = deduplicateReportsByDay(valid);
      const local = getLocalCache();
      const merged = mergeReports(unique, local);
      setLocalCache(merged);
      pushReportsToServer(merged).catch(() => {});
      return merged;
    }
  } catch (err) {
    console.warn("Notice checking Firestore reports:", err);
  }

  // 2. Tenta buscar do Supabase se configurado
  try {
    const sbReports = await fetchSupabaseReports();
    if (sbReports && sbReports.length > 0) {
      const valid = sbReports.filter((r) => r && r.id && !deletedSet.has(r.id));
      const { unique } = deduplicateReportsByDay(valid);
      const local = getLocalCache();
      const merged = mergeReports(unique, local);
      setLocalCache(merged);
      pushReportsToServer(merged).catch(() => {});
      return merged;
    }
  } catch (err) {
    console.warn("Notice checking Supabase reports:", err);
  }

  // 3. Tenta sincronizar via servidor TanStack/Node API (resolve celular zerado instantaneamente)
  try {
    const { reports: srvMerged } = await syncCloudReportsAcrossDevices();
    if (srvMerged && srvMerged.length > 0) {
      setLocalCache(srvMerged);
    }
  } catch (e) {
    console.warn("Server sync error in fetchCloudReports:", e);
  }

  return getLocalCache();
}

export function subscribeCloudReports(onChange: (reports: SavedCloudReport[]) => void): () => void {
  let isSubscribed = true;

  // 1. Emite cache do dispositivo imediatamente (sem tela em branco)
  const initialCache = getLocalCache();
  onChange(initialCache);

  // 2. Conexão em tempo real nativa com o Firestore (onSnapshot)
  // Permite que qualquer celular ou computador receba relatórios novos ou modificados instantaneamente
  let unsubscribeFirestore: (() => void) | null = null;
  try {
    const colRef = collection(db, COLLECTION_NAME);
    unsubscribeFirestore = onSnapshot(
      colRef,
      (snapshot) => {
        if (!isSubscribed) return;
        const reports: SavedCloudReport[] = [];
        snapshot.forEach((docSnap) => {
          reports.push(mapFirestoreDocToReport(docSnap.id, docSnap.data()));
        });
        if (reports.length > 0) {
          unblacklistActiveReports(reports);
          const cur = getLocalCache();
          const merged = mergeReports(reports, cur);
          setLocalCache(merged);
          onChange(merged);
        }
      },
      (err) => {
        console.warn("[Firestore Realtime notice]:", err);
      },
    );
  } catch (err) {
    console.warn("Firestore subscription error:", err);
  }

  // 3. Sincroniza imediatamente com Firestore e servidor
  fetchFirestoreReports()
    .then((fsReports) => {
      if (isSubscribed && fsReports && fsReports.length > 0) {
        const cur = getLocalCache();
        const merged = mergeReports(fsReports, cur);
        setLocalCache(merged);
        onChange(merged);
      }
    })
    .catch(() => {});

  syncCloudReportsAcrossDevices()
    .then(({ reports }) => {
      if (isSubscribed && reports && reports.length > 0) {
        onChange(reports);
      }
    })
    .catch((err) => console.warn("Notice initial sync with server:", err));

  // 4. Conexão em tempo real do Supabase Realtime se disponível
  let sbChannel: ReturnType<typeof supabase.channel> | null = null;
  try {
    if (supabase && typeof supabase.channel === "function") {
      sbChannel = supabase
        .channel("relatorios_realtime_stream")
        .on("postgres_changes", { event: "*", schema: "public", table: "relatorios" }, async () => {
          try {
            const fresh = await fetchSupabaseReports();
            if (fresh && fresh.length > 0 && isSubscribed) {
              const cur = getLocalCache();
              const merged = mergeReports(fresh, cur);
              setLocalCache(merged);
              onChange(merged);
            }
          } catch (err) {
            console.warn("Error receiving Supabase Realtime change:", err);
          }
        })
        .subscribe();
    }
  } catch (err) {
    console.warn("Supabase Realtime subscription notice:", err);
  }

  // 5. Conexão em tempo real via Server-Sent Events (SSE) para sincronização instantânea PC <-> Celular
  let eventSource: EventSource | null = null;
  if (typeof window !== "undefined" && typeof EventSource !== "undefined") {
    try {
      eventSource = new EventSource("/api/cloud-reports-stream");
      eventSource.onmessage = (event) => {
        try {
          const payload = JSON.parse(event.data);
          if (payload && Array.isArray(payload.reports)) {
            const currentLocal = getLocalCache();
            const merged = mergeReports(payload.reports, currentLocal);
            setLocalCache(merged);
            if (isSubscribed) {
              onChange(merged);
            }
          }
        } catch (e) {
          console.warn("Notice parsing SSE event:", e);
        }
      };
      eventSource.onerror = () => {
        // Reconexão automática pelo navegador
      };
    } catch (err) {
      console.warn("SSE connection notice:", err);
    }
  }

  // 6. Sincronização periódica e ao focar/desbloquear a tela do celular
  const pollInterval =
    typeof window !== "undefined"
      ? setInterval(() => {
          fetchFirestoreReports()
            .then((fs) => {
              if (fs && fs.length > 0 && isSubscribed) {
                const cur = getLocalCache();
                const merged = mergeReports(fs, cur);
                if (merged.length !== cur.length) {
                  setLocalCache(merged);
                  onChange(merged);
                }
              }
            })
            .catch(() => {});
        }, 15000)
      : null;

  const handleVisibilityOrFocus = () => {
    if (typeof document !== "undefined" && !document.hidden && isSubscribed) {
      fetchFirestoreReports()
        .then((fs) => {
          if (isSubscribed && fs && fs.length > 0) {
            const cur = getLocalCache();
            const merged = mergeReports(fs, cur);
            setLocalCache(merged);
            onChange(merged);
          }
        })
        .catch(() => {});

      syncCloudReportsAcrossDevices()
        .then(({ reports }) => {
          if (isSubscribed && reports && reports.length > 0) onChange(reports);
        })
        .catch(() => {});
    }
  };

  if (typeof window !== "undefined") {
    window.addEventListener("visibilitychange", handleVisibilityOrFocus);
    window.addEventListener("focus", handleVisibilityOrFocus);
  }

  return () => {
    isSubscribed = false;
    if (unsubscribeFirestore) {
      try {
        unsubscribeFirestore();
      } catch {}
    }
    if (sbChannel && supabase && typeof supabase.removeChannel === "function") {
      try {
        supabase.removeChannel(sbChannel);
      } catch {}
    }
    if (eventSource) {
      eventSource.close();
    }
    if (pollInterval) {
      clearInterval(pollInterval);
    }
    if (typeof window !== "undefined") {
      window.removeEventListener("visibilitychange", handleVisibilityOrFocus);
      window.removeEventListener("focus", handleVisibilityOrFocus);
    }
  };
}

export async function deleteCloudReport(id: string): Promise<void> {
  if (!id) return;

  // 1. Marca como excluído no cache local imediatamente
  addDeletedId(id);

  // 2. Remove do IndexedDB, do cache local e emite atualização na UI
  deleteReportFromIndexedDB(id).catch(() => {});
  const updatedCache = getLocalCache().filter((r) => r.id !== id);
  setLocalCache(updatedCache);

  // 3. Exclui do Firestore
  try {
    const docRef = doc(db, COLLECTION_NAME, id);
    await deleteDoc(docRef);
  } catch (fsErr) {
    console.warn("[Firestore] Error deleting report:", fsErr);
  }

  // 4. Exclui do servidor da aplicação
  deleteServerReport(id).catch((e) => console.warn("Error deleting from server:", e));

  // 5. Exclui do Supabase se configurado
  try {
    if (supabase && typeof supabase.from === "function") {
      await supabase.from("relatorios").delete().eq("id", id);
    }
  } catch (err) {
    console.warn("Supabase delete notice:", err);
  }
}

export function cloudReportToHeaderAndEvents(rawReport: SavedCloudReport): {
  header: DailyHeader;
  events: WorkEvent[];
  photos: ReportPhoto[];
} {
  // Procura se em memória temos a versão com fotos preservadas
  let report = rawReport;
  if (
    (!report.photos ||
      report.photos.length === 0 ||
      !report.photos.some((p) => Boolean(p.dataUrl))) &&
    memoryReportsCache
  ) {
    const fromMem = memoryReportsCache.find((r) => r.id === rawReport.id);
    if (
      fromMem &&
      fromMem.photos &&
      fromMem.photos.length > 0 &&
      fromMem.photos.some((p) => Boolean(p.dataUrl))
    ) {
      report = fromMem;
    }
  }

  const header: DailyHeader = {
    ...emptyHeader(),
    data: report.data || "",
    nome: report.nome || "",
    local: report.local || "",
    sinais: report.sinais || "",
    linha: report.linha || "7",
    via: report.via || "",
    gda: report.gda || "",
    om: report.om || "",
    empresa: report.empresa || "",
    atividade: report.atividade || "",
    colaboradores: report.colaboradores || "",
    descritivo: report.descritivo || "",
    kmProgramado: report.kmProgramado || "",
    executado: report.executado || "",
    observacaoGeral: report.observacaoGeral || "",
    telefone: report.telefone || "",
  };

  return {
    header,
    events: sortAndNormalizeEvents(report.events ?? []),
    photos: report.photos || [],
  };
}
