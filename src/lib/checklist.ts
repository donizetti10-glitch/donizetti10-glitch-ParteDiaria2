import { supabase } from "@/integrations/supabase/client";
import { db, collection, doc, getDocs, setDoc, deleteDoc, onSnapshot } from "./firebase";

export type ChecklistItemStatus = "conforme" | "nao_conforme" | "nao_se_aplica" | "pendente";
export type EquipmentOperatingStatus = "apto" | "atencao" | "inapto";

export interface ChecklistItem {
  id: string;
  category: string;
  name: string;
  description?: string;
  required?: boolean;
  status: ChecklistItemStatus;
  notes?: string;
  acaoSolucao?: string;
}

export interface ChecklistPhoto {
  id: string;
  dataUrl: string;
  caption?: string;
  itemId?: string;
  itemName?: string;
  timestamp?: number;
}

export interface EquipmentChecklist {
  id: string;
  createdAt: number;
  updatedAt: number;
  data: string; // YYYY-MM-DD
  hora: string; // HH:MM
  equipamento: string;
  prefixo: string;
  horimetroMotor: string;
  horimetroImplementoFerroviario: string;
  horimetroGerador: string;
  quilometragemCaminhao: string;
  horimetroOuKm?: string;
  operador: string;
  turno: string;
  local: string;
  statusGeral: EquipmentOperatingStatus;
  observacoesGerais: string;
  items: ChecklistItem[];
  fotosCount?: number;
  fotos?: ChecklistPhoto[];
}

export const CHECKLIST_STORAGE_KEY = "equipment-checklists-v1";
export const CHECKLIST_DRAFT_KEY = "equipment-checklist-draft-v1";
export const CHECKLIST_TEMPLATE_KEY = "equipment-checklist-template-vrf-ph1016";
export const FIRESTORE_CHECKLISTS_COLLECTION = "checklists_equipamento";

export const VRF_DOCUMENT_CODE = "VRF-SM PH-1016";
export const VRF_DOCUMENT_TITLE = "Inspeção Pré-Operação";
export const VRF_DOCUMENT_SUBTITLE =
  "Antes do início de cada operação é necessário que o operador faça um checklist diário dos seguintes componentes:";

export const DEFAULT_CHECKLIST_TEMPLATE: Omit<ChecklistItem, "status" | "notes">[] = [
  {
    id: "vrf-01",
    category: "Motor e Fluidos",
    name: "Verificar o nível de óleo do motor e freio;",
    description: "Conferir nível da vareta e reservatório com motor desligado",
    required: true,
  },
  {
    id: "vrf-02",
    category: "Motor e Fluidos",
    name: "Verificar o nível de arrefecimento;",
    description: "Líquido do reservatório de expansão / radiador",
    required: true,
  },
  {
    id: "vrf-03",
    category: "Motor e Fluidos",
    name: "Verificar o nível de combustível;",
    description: "Volume suficiente para o turno planejado de operação",
    required: true,
  },
  {
    id: "vrf-arla",
    category: "Motor e Fluidos",
    name: "Verificar nível do ARLA;",
    description: "Conferir nível do reservatório de ARLA 32 (Agente Redutor Líquido Automotivo)",
    required: true,
  },
  {
    id: "vrf-04",
    category: "Motor e Fluidos",
    name: "Verificar o nível de óleo da caixa de reversão;",
    description: "Visor ou vareta no nível indicado pelo fabricante",
    required: true,
  },
  {
    id: "vrf-05",
    category: "Eixos e Rodagem",
    name: "Verificar se existe vazamento no eixo traseiro;",
    description: "Inspeção visual de retentores, carcaça e cubo de roda",
    required: true,
  },
  {
    id: "vrf-06",
    category: "Eixos e Rodagem",
    name: "Verificar porcas e parafusos de fixação da roda e eixo;",
    description: "Torque e ausência de parafusos soltos ou faltantes",
    required: true,
  },
  {
    id: "vrf-07",
    category: "Eixos e Rodagem",
    name: "Verificar calibragem e desgaste/avarias dos pneus;",
    description: "Pressão correta, sem cortes, bolhas ou arames expostos",
    required: true,
  },
  {
    id: "vrf-08",
    category: "Kits Ferroviários",
    name: "Verificar cilindros hidráulicos dos kits ferroviários e mangueiras;",
    description: "Sem vazamentos de óleo hidráulico ou trincas em mangueiras",
    required: true,
  },
  {
    id: "vrf-09",
    category: "Kits Ferroviários",
    name: "Verificar avarias e fixação do limpa trilhos e borrachas;",
    description: "Borrachas íntegras e limpa trilhos com altura e fixação corretas",
    required: true,
  },
  {
    id: "vrf-10",
    category: "Cabine e Acessórios",
    name: "Verificar componentes e acessórios (Retrovisores, espelhos, luzes auxiliares, limpador de para-brisas...);",
    description: "Espelhos regulados, iluminação e limpadores funcionando",
    required: true,
  },
  {
    id: "vrf-11",
    category: "Segurança e Operação",
    name: "Verificar integridade do cinto de segurança e sua fixação;",
    description: "Travamento eficaz e ancoragem sem folgas ou desgaste",
    required: true,
  },
  {
    id: "vrf-12",
    category: "Segurança e Operação",
    name: "Verificar câmera e alarme de ré;",
    description: "Visão no monitor e sinal sonoro nítido ao engatar a marcha a ré",
    required: true,
  },
  {
    id: "vrf-13",
    category: "Comandos e Acionamento",
    name: "Verificar funcionamento da tomada de força;",
    description: "Engate suave e operação estável sem ruídos anormais",
    required: true,
  },
  {
    id: "vrf-14",
    category: "Comandos e Acionamento",
    name: "Verificar acionamentos e comandos via IHM;",
    description: "Tela sensível ou botões da interface respondendo perfeitamente",
    required: true,
  },
  {
    id: "vrf-15",
    category: "Kits Ferroviários",
    name: "Verificar acionamento dos kits ferroviários;",
    description: "Descida e subida dos truques/rodas de guia rodoferroviárias",
    required: true,
  },
  {
    id: "vrf-16",
    category: "Segurança e Operação",
    name: "Verificar funcionamento do freio de serviço e estacionamento;",
    description: "Pedal firme com retenção e freio de estacionamento segurando",
    required: true,
  },
  {
    id: "vrf-17",
    category: "Engates e Estrutura",
    name: "Verificar integridade do cambão;",
    description: "Haste de reboque, pinos e travas de segurança sem deformações",
    required: true,
  },
  {
    id: "vrf-18",
    category: "Carro Ultrassom",
    name: "Verificar componentes do carro ultrassom;",
    description: "Sapatas, transdutores, cabos e estrutura do carrinho de ultrassom",
    required: true,
  },
  {
    id: "vrf-19",
    category: "Segurança e Operação",
    name: "Verificar botoeiras de parada de emergência;",
    description: "Corte imediato de força quando acionada e rearme correto",
    required: true,
  },
  {
    id: "vrf-20",
    category: "Carro Ultrassom",
    name: "Testar funções do carro ultrassom.",
    description: "Emissão de sinal acústico, comunicação e rastreamento da via",
    required: true,
  },
];

export function createEmptyChecklist(): EquipmentChecklist {
  const now = new Date();
  const dateStr = now.toISOString().slice(0, 10);
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");

  const template = getChecklistTemplate();

  return {
    id: `chk-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    data: dateStr,
    hora: `${hours}:${minutes}`,
    equipamento: "",
    prefixo: "",
    horimetroMotor: "",
    horimetroImplementoFerroviario: "",
    horimetroGerador: "",
    quilometragemCaminhao: "",
    horimetroOuKm: "",
    operador: "",
    turno: "Diurno",
    local: "",
    statusGeral: "apto",
    observacoesGerais: "",
    items: template.map((item) => ({
      ...item,
      status: "pendente",
      notes: "",
      acaoSolucao: "",
    })),
    fotos: [],
    fotosCount: 0,
  };
}

export function getChecklistTemplate(): Omit<ChecklistItem, "status" | "notes" | "acaoSolucao">[] {
  if (typeof window === "undefined") return DEFAULT_CHECKLIST_TEMPLATE;
  try {
    const stored = localStorage.getItem(CHECKLIST_TEMPLATE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      if (Array.isArray(parsed) && parsed.length > 0) {
        // Assegurar que o item ARLA esteja presente mesmo em templates já salvos no cache local
        const hasArla = parsed.some(
          (item: { name?: unknown }) =>
            typeof item?.name === "string" && item.name.toLowerCase().includes("arla"),
        );
        if (!hasArla) {
          const arlaItem: Omit<ChecklistItem, "status" | "notes" | "acaoSolucao"> = {
            id: "vrf-arla",
            category: "Motor e Fluidos",
            name: "Verificar nível do ARLA;",
            description:
              "Conferir nível do reservatório de ARLA 32 (Agente Redutor Líquido Automotivo)",
            required: true,
          };
          const combustivelIdx = parsed.findIndex(
            (it: { name?: unknown }) =>
              typeof it?.name === "string" && it.name.toLowerCase().includes("combustível"),
          );
          if (combustivelIdx >= 0) {
            parsed.splice(combustivelIdx + 1, 0, arlaItem);
          } else {
            parsed.push(arlaItem);
          }
          try {
            localStorage.setItem(CHECKLIST_TEMPLATE_KEY, JSON.stringify(parsed));
          } catch (e) {
            console.warn("Erro ao atualizar template com ARLA:", e);
          }
        }
        return parsed;
      }
    }
  } catch (err) {
    console.warn("Erro ao ler template de checklist:", err);
  }
  return DEFAULT_CHECKLIST_TEMPLATE;
}

export function saveChecklistTemplate(items: Omit<ChecklistItem, "status" | "notes">[]) {
  if (typeof window === "undefined") return;
  try {
    localStorage.setItem(CHECKLIST_TEMPLATE_KEY, JSON.stringify(items));
  } catch (err) {
    console.warn("Erro ao salvar template de checklist:", err);
  }
}

export function calculateOperatingStatus(items: ChecklistItem[]): EquipmentOperatingStatus {
  const nonConforming = items.filter((i) => i.status === "nao_conforme");
  if (nonConforming.length === 0) {
    return "apto";
  }
  const hasCriticalNonConforming = nonConforming.some(
    (i) =>
      i.required ||
      i.category === "Segurança e Operação" ||
      i.category === "Eixos e Rodagem" ||
      i.category === "Kits Ferroviários" ||
      i.category === "Motor e Fluidos",
  );
  return hasCriticalNonConforming ? "inapto" : "atencao";
}

// Local Storage helpers
export function getLocalChecklists(): EquipmentChecklist[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = localStorage.getItem(CHECKLIST_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function saveLocalChecklist(checklist: EquipmentChecklist): void {
  if (typeof window === "undefined") return;
  try {
    const current = getLocalChecklists();
    const existingIdx = current.findIndex((c) => c.id === checklist.id);
    let updated: EquipmentChecklist[];
    if (existingIdx >= 0) {
      updated = [...current];
      updated[existingIdx] = checklist;
    } else {
      updated = [checklist, ...current];
    }
    localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(updated));
  } catch (err) {
    console.warn("Error saving local checklist:", err);
  }
}

export function deleteLocalChecklist(id: string): void {
  if (typeof window === "undefined") return;
  try {
    const current = getLocalChecklists();
    const filtered = current.filter((c) => c.id !== id);
    localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(filtered));
  } catch (err) {
    console.warn("Erro ao remover checklist local:", err);
  }
}

/**
 * Verifica se um checklist possui preenchimento relevante que justifique manter como rascunho
 */
export function isChecklistMeaningful(checklist: EquipmentChecklist): boolean {
  if (!checklist) return false;
  if (
    Boolean(checklist.equipamento?.trim()) ||
    Boolean(checklist.prefixo?.trim()) ||
    Boolean(checklist.operador?.trim()) ||
    Boolean(checklist.local?.trim()) ||
    Boolean(checklist.horimetroMotor?.trim()) ||
    Boolean(checklist.horimetroImplementoFerroviario?.trim()) ||
    Boolean(checklist.horimetroGerador?.trim()) ||
    Boolean(checklist.quilometragemCaminhao?.trim()) ||
    Boolean(checklist.horimetroOuKm?.trim()) ||
    Boolean(checklist.observacoesGerais?.trim())
  ) {
    return true;
  }
  if (Array.isArray(checklist.fotos) && checklist.fotos.length > 0) {
    return true;
  }
  if (Array.isArray(checklist.items)) {
    const hasProgress = checklist.items.some(
      (it) =>
        it.status !== "pendente" || Boolean(it.notes?.trim()) || Boolean(it.acaoSolucao?.trim()),
    );
    if (hasProgress) return true;
  }
  return false;
}

/**
 * Salva o rascunho em andamento no armazenamento local persistente do celular/computador.
 * Permite que o operador feche o aplicativo, perca a internet ou reinicie o aparelho sem perder nada.
 */
export function saveChecklistDraft(checklist: EquipmentChecklist): void {
  if (typeof window === "undefined") return;
  try {
    if (!isChecklistMeaningful(checklist)) {
      localStorage.removeItem(CHECKLIST_DRAFT_KEY);
      return;
    }
    const payload = {
      checklist: {
        ...checklist,
        updatedAt: Date.now(),
      },
      savedAt: Date.now(),
    };
    localStorage.setItem(CHECKLIST_DRAFT_KEY, JSON.stringify(payload));
  } catch (err) {
    console.warn("Erro ao salvar rascunho de checklist local:", err);
  }
}

/**
 * Recupera o rascunho do checklist salvo localmente no aparelho.
 */
export function getChecklistDraft(): { checklist: EquipmentChecklist; savedAt: number } | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = localStorage.getItem(CHECKLIST_DRAFT_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed) return null;

    const checklist: EquipmentChecklist = parsed.checklist || parsed;
    const savedAt: number = Number(parsed.savedAt) || checklist.updatedAt || Date.now();

    if (checklist && Array.isArray(checklist.items) && checklist.items.length > 0) {
      if (isChecklistMeaningful(checklist)) {
        return { checklist, savedAt };
      }
    }
    return null;
  } catch (err) {
    console.warn("Erro ao ler rascunho do checklist local:", err);
    return null;
  }
}

/**
 * Descarta ou remove o rascunho ativo do dispositivo após finalização ou cancelamento.
 */
export function clearChecklistDraft(): void {
  if (typeof window === "undefined") return;
  try {
    localStorage.removeItem(CHECKLIST_DRAFT_KEY);
  } catch (err) {
    console.warn("Erro ao remover rascunho do checklist local:", err);
  }
}

// Cloud Firestore operations
export function sanitizeChecklistForFirestore<T>(obj: T): T {
  if (obj === null || obj === undefined) {
    return null as unknown as T;
  }
  if (typeof obj !== "object") {
    return obj;
  }
  if (Array.isArray(obj)) {
    return obj.map((item) => sanitizeChecklistForFirestore(item)) as unknown as T;
  }
  const cleaned: Record<string, unknown> = {};
  for (const [key, val] of Object.entries(obj as Record<string, unknown>)) {
    if (val !== undefined) {
      cleaned[key] = sanitizeChecklistForFirestore(val);
    }
  }
  return cleaned as T;
}

export function mapSupabaseRowToChecklist(row: Record<string, unknown>): EquipmentChecklist {
  return {
    id: String(row.id || ""),
    createdAt: Number(
      row.createdAt || (row.created_at ? new Date(String(row.created_at)).getTime() : Date.now()),
    ),
    updatedAt: Number(
      row.updatedAt || (row.updated_at ? new Date(String(row.updated_at)).getTime() : Date.now()),
    ),
    data: String(row.data || ""),
    hora: String(row.hora || ""),
    equipamento: String(row.equipamento || ""),
    prefixo: String(row.prefixo || ""),
    horimetroMotor: String(row.horimetroMotor || row.horimetro_motor || ""),
    horimetroImplementoFerroviario: String(
      row.horimetroImplementoFerroviario || row.horimetro_implemento_ferroviario || "",
    ),
    horimetroGerador: String(row.horimetroGerador || row.horimetro_gerador || ""),
    quilometragemCaminhao: String(row.quilometragemCaminhao || row.quilometragem_caminhao || ""),
    horimetroOuKm: String(row.horimetroOuKm || row.horimetro_ou_km || ""),
    operador: String(row.operador || ""),
    turno: String(row.turno || ""),
    local: String(row.local || ""),
    statusGeral: (row.statusGeral || row.status_geral || "apto") as EquipmentOperatingStatus,
    observacoesGerais: String(row.observacoesGerais || row.observacoes_gerais || ""),
    items: Array.isArray(row.items) ? (row.items as ChecklistItem[]) : [],
    fotos: Array.isArray(row.fotos) ? (row.fotos as ChecklistPhoto[]) : [],
    fotosCount: Number(
      row.fotosCount || row.fotos_count || (Array.isArray(row.fotos) ? row.fotos.length : 0),
    ),
  };
}

export async function fetchFirestoreChecklists(): Promise<EquipmentChecklist[]> {
  try {
    const colRef = collection(db, FIRESTORE_CHECKLISTS_COLLECTION);
    const snap = await getDocs(colRef);
    if (!snap.empty) {
      const list: EquipmentChecklist[] = [];
      snap.forEach((docSnap) => {
        list.push(mapSupabaseRowToChecklist({ id: docSnap.id, ...docSnap.data() }));
      });
      list.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
      return list;
    }
  } catch (err) {
    console.warn("[Firestore] Error fetching checklists:", err);
  }
  return [];
}

export async function fetchSupabaseChecklists(): Promise<EquipmentChecklist[]> {
  // Dá prioridade primária ao Firestore oficial em nuvem
  const fsList = await fetchFirestoreChecklists();
  if (fsList && fsList.length > 0) {
    return fsList;
  }
  try {
    if (!supabase || typeof supabase.from !== "function") return [];
    const { data, error } = await supabase
      .from("checklists_equipamento")
      .select("*")
      .order("createdAt", { ascending: false });

    if (!error && Array.isArray(data) && data.length > 0) {
      return data.map(mapSupabaseRowToChecklist);
    }
  } catch (err) {
    console.warn("Supabase fetch checklists notice:", err);
  }
  return [];
}

export async function saveChecklistToCloud(
  checklist: EquipmentChecklist,
): Promise<EquipmentChecklist> {
  const normalizedItems: ChecklistItem[] = (checklist.items || []).map((item, idx) => ({
    id: item.id || `item-${idx}`,
    category: item.category || "Geral",
    name: item.name || "",
    description: item.description || "",
    required: item.required ?? true,
    status: item.status || "conforme",
    notes: item.notes || "",
    acaoSolucao: item.acaoSolucao || "",
  }));

  const fotos: ChecklistPhoto[] = Array.isArray(checklist.fotos) ? checklist.fotos : [];
  const fotosCount = fotos.length;

  const now = Date.now();
  const itemToSave: EquipmentChecklist = {
    ...checklist,
    id: checklist.id || `chk-${now}`,
    data: checklist.data || new Date().toISOString().split("T")[0],
    hora: checklist.hora || "",
    equipamento: checklist.equipamento || "",
    prefixo: checklist.prefixo || "",
    horimetroMotor: checklist.horimetroMotor || "",
    horimetroImplementoFerroviario: checklist.horimetroImplementoFerroviario || "",
    horimetroGerador: checklist.horimetroGerador || "",
    quilometragemCaminhao: checklist.quilometragemCaminhao || "",
    horimetroOuKm: checklist.horimetroOuKm || "",
    operador: checklist.operador || "",
    turno: checklist.turno || "",
    local: checklist.local || "",
    statusGeral: checklist.statusGeral || "apto",
    observacoesGerais: checklist.observacoesGerais || "",
    items: normalizedItems,
    fotos,
    fotosCount,
    updatedAt: now,
    createdAt: checklist.createdAt || now,
  };

  // 1. Sempre salva localmente primeiro (resiliência offline)
  saveLocalChecklist(itemToSave);

  // 2. Salva no Firestore como banco oficial em nuvem (visível em todos os celulares)
  try {
    const docRef = doc(db, FIRESTORE_CHECKLISTS_COLLECTION, itemToSave.id);
    const sanitized = sanitizeChecklistForFirestore(itemToSave);
    await setDoc(docRef, sanitized, { merge: true });
  } catch (fsErr) {
    console.warn("[Firestore] Error saving checklist:", fsErr);
  }

  // 3. Salva no Supabase se configurado
  try {
    if (supabase && typeof supabase.from === "function") {
      await supabase.from("checklists_equipamento").upsert(
        {
          id: itemToSave.id,
          data: itemToSave.data,
          hora: itemToSave.hora,
          equipamento: itemToSave.equipamento,
          prefixo: itemToSave.prefixo,
          horimetroMotor: itemToSave.horimetroMotor,
          horimetroImplementoFerroviario: itemToSave.horimetroImplementoFerroviario,
          horimetroGerador: itemToSave.horimetroGerador,
          quilometragemCaminhao: itemToSave.quilometragemCaminhao,
          horimetroOuKm: itemToSave.horimetroOuKm,
          operador: itemToSave.operador,
          turno: itemToSave.turno,
          local: itemToSave.local,
          statusGeral: itemToSave.statusGeral,
          observacoesGerais: itemToSave.observacoesGerais,
          items: itemToSave.items,
          fotos: itemToSave.fotos,
          fotosCount: itemToSave.fotosCount || 0,
          createdAt: Number(itemToSave.createdAt),
          updatedAt: Number(itemToSave.updatedAt),
          created_at: new Date(itemToSave.createdAt).toISOString(),
          updated_at: new Date(now).toISOString(),
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
        } as any,
        { onConflict: "id" },
      );
    }
  } catch (err) {
    console.warn("Notice saving checklist to Supabase:", err);
  }

  // 4. Sincroniza com o servidor da aplicação para replicação instantânea
  pushChecklistsToServer([itemToSave]).catch(() => {});

  return itemToSave;
}

export async function fetchServerChecklists(): Promise<EquipmentChecklist[]> {
  if (typeof window === "undefined") return [];
  try {
    const res = await fetch("/api/checklists", { cache: "no-store" });
    if (res.ok) {
      const data = await res.json();
      if (data && Array.isArray(data.checklists)) {
        return data.checklists;
      }
    }
  } catch (err) {
    console.warn("Notice fetching server checklists:", err);
  }
  return [];
}

export async function pushChecklistsToServer(checklists: EquipmentChecklist[]): Promise<void> {
  if (typeof window === "undefined" || !Array.isArray(checklists) || checklists.length === 0)
    return;
  try {
    await fetch("/api/checklists", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ checklists }),
    });
  } catch (err) {
    console.warn("Notice pushing checklists to server:", err);
  }
}

export async function deleteServerChecklist(id: string): Promise<void> {
  if (typeof window === "undefined" || !id) return;
  try {
    await fetch(`/api/checklists?id=${encodeURIComponent(id)}`, { method: "DELETE" });
  } catch (err) {
    console.warn("Notice deleting checklist from server:", err);
  }
}

export async function deleteChecklistFromCloud(id: string): Promise<void> {
  deleteLocalChecklist(id);

  // 1. Exclui do Firestore
  try {
    const docRef = doc(db, FIRESTORE_CHECKLISTS_COLLECTION, id);
    await deleteDoc(docRef);
  } catch (fsErr) {
    console.warn("[Firestore] Error deleting checklist:", fsErr);
  }

  // 2. Exclui do servidor
  deleteServerChecklist(id).catch(() => {});

  // 3. Exclui do Supabase se configurado
  try {
    if (supabase && typeof supabase.from === "function") {
      await supabase.from("checklists_equipamento").delete().eq("id", id);
    }
  } catch (err) {
    console.warn("Supabase delete checklist notice:", err);
  }
}

export async function fetchAllCloudChecklists(): Promise<EquipmentChecklist[]> {
  // 1. Tenta buscar no Firestore (banco oficial em nuvem do projeto)
  try {
    const fsChecklists = await fetchFirestoreChecklists();
    if (fsChecklists && fsChecklists.length > 0) {
      if (typeof window !== "undefined") {
        localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(fsChecklists));
      }
      pushChecklistsToServer(fsChecklists).catch(() => {});
      return fsChecklists;
    }
  } catch (err) {
    console.warn("Notice fetching Firestore checklists:", err);
  }

  // 2. Tenta buscar no Supabase
  try {
    const sbChecklists = await fetchSupabaseChecklists();
    if (sbChecklists && sbChecklists.length > 0) {
      if (typeof window !== "undefined") {
        localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(sbChecklists));
      }
      pushChecklistsToServer(sbChecklists).catch(() => {});
      return sbChecklists;
    }
  } catch (err) {
    console.warn("Notice fetching Supabase checklists:", err);
  }

  // 3. Tenta buscar do servidor
  try {
    const srvChecklists = await fetchServerChecklists();
    if (srvChecklists && srvChecklists.length > 0) {
      if (typeof window !== "undefined") {
        localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(srvChecklists));
      }
      return srvChecklists;
    }
  } catch (err) {
    console.warn("Notice fetching server checklists:", err);
  }

  return getLocalChecklists();
}

export function subscribeChecklists(
  onUpdate: (checklists: EquipmentChecklist[]) => void,
): () => void {
  let isSubscribed = true;

  // 1. Emite dados locais imediatamente
  onUpdate(getLocalChecklists());

  // 2. Conexão em tempo real nativa com o Firestore (onSnapshot)
  // Permite que outros celulares e computadores recebam checklists instantaneamente
  let unsubscribeFirestore: (() => void) | null = null;
  try {
    const colRef = collection(db, FIRESTORE_CHECKLISTS_COLLECTION);
    unsubscribeFirestore = onSnapshot(
      colRef,
      (snapshot) => {
        if (!isSubscribed) return;
        const list: EquipmentChecklist[] = [];
        snapshot.forEach((docSnap) => {
          list.push(mapSupabaseRowToChecklist({ id: docSnap.id, ...docSnap.data() }));
        });
        if (list.length > 0) {
          list.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
          if (typeof window !== "undefined") {
            localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(list));
          }
          onUpdate(list);
        }
      },
      (err) => {
        console.warn("[Firestore Checklists Realtime notice]:", err);
      },
    );
  } catch (err) {
    console.warn("Firestore checklist subscription error:", err);
  }

  // 3. Sincroniza com o Firestore e Supabase
  fetchFirestoreChecklists()
    .then((fsList) => {
      if (isSubscribed && fsList && fsList.length > 0) {
        if (typeof window !== "undefined") {
          localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(fsList));
        }
        pushChecklistsToServer(fsList).catch(() => {});
        onUpdate(fsList);
      }
    })
    .catch(() => {});

  // 4. Sincroniza com o servidor
  fetchServerChecklists()
    .then((srvList) => {
      if (isSubscribed && srvList && srvList.length > 0) {
        if (typeof window !== "undefined") {
          localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(srvList));
        }
        onUpdate(srvList);
      }
    })
    .catch(() => {});

  // 5. Supabase Realtime channel se disponível
  let sbChannel: ReturnType<typeof supabase.channel> | null = null;
  try {
    if (supabase && typeof supabase.channel === "function") {
      sbChannel = supabase
        .channel("checklists_realtime_stream")
        .on(
          "postgres_changes",
          { event: "*", schema: "public", table: "checklists_equipamento" },
          async () => {
            try {
              const fresh = await fetchSupabaseChecklists();
              if (fresh && fresh.length > 0 && isSubscribed) {
                if (typeof window !== "undefined") {
                  localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(fresh));
                }
                onUpdate(fresh);
              }
            } catch (err) {
              console.warn("Error receiving Supabase Realtime checklist change:", err);
            }
          },
        )
        .subscribe();
    }
  } catch (err) {
    console.warn("Supabase Realtime checklist notice:", err);
  }

  // 6. Server-Sent Events (SSE) para sincronização instantânea de checklists
  let eventSource: EventSource | null = null;
  if (typeof window !== "undefined" && typeof EventSource !== "undefined") {
    try {
      eventSource = new EventSource("/api/checklists-stream");
      eventSource.onmessage = (event) => {
        try {
          const payload = JSON.parse(event.data);
          if (payload && Array.isArray(payload.checklists) && isSubscribed) {
            if (typeof window !== "undefined") {
              localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(payload.checklists));
            }
            onUpdate(payload.checklists);
          }
        } catch (e) {
          console.warn("Notice parsing checklist SSE event:", e);
        }
      };
    } catch (err) {
      console.warn("Checklist SSE error:", err);
    }
  }

  // 7. Atualização ao focar/desbloquear a tela
  const handleVisibilityOrFocus = () => {
    if (typeof document !== "undefined" && !document.hidden && isSubscribed) {
      fetchFirestoreChecklists()
        .then((fs) => {
          if (isSubscribed && fs && fs.length > 0) {
            if (typeof window !== "undefined") {
              localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(fs));
            }
            onUpdate(fs);
          }
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
    if (typeof window !== "undefined") {
      window.removeEventListener("visibilitychange", handleVisibilityOrFocus);
      window.removeEventListener("focus", handleVisibilityOrFocus);
    }
  };
}
