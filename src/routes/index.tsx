import { createFileRoute, Link } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {
  MoreVertical,
  Pencil,
  Trash2,
  Cloud,
  FileSpreadsheet,
  Check,
  Database,
  FileText,
  Copy,
  RefreshCw,
  Wifi,
  WifiOff,
  ClipboardCheck,
} from "lucide-react";
import { CloudReportsDatabase } from "@/components/CloudReportsDatabase";
import { PhotoUploader } from "@/components/PhotoUploader";
import { ThemeToggle } from "@/components/ThemeToggle";
import { buildPdfBlobForReport, downloadReportPdf } from "@/lib/pdf-export";
import {
  saveReportToCloud,
  cloudReportToHeaderAndEvents,
  subscribeCloudReports,
  getLocalCache,
  type SavedCloudReport,
} from "@/lib/cloud-reports";
import { exportSingleReportToExcel, buildSingleReportExcelBlob } from "@/lib/excel-export";
import { SapIw41Modal } from "@/components/SapIw41Modal";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  EVENT_LABEL,
  buildCSV,
  buildWhatsAppMessage,
  emptyHeader,
  formatDate,
  formatDuration,
  formatTime,
  generatePdfFilename,
  maskSinais,
  maskGda,
  maskProgramadoExecutado,
  totals,
  type DailyHeader,
  type EventKind,
  type WorkEvent,
  type ReportPhoto,
} from "@/lib/parte-diaria";
import { fetchLastDefect, saveLastDefect, subscribeLastDefect } from "@/lib/shared-defect";

export const Route = createFileRoute("/")({
  component: PartesDiariasPage,
});

const STORAGE_KEY = "parte-diaria-v1";
const ACTIVITIES_STORAGE_KEY = "activities-history-v1";
const COMPANIES_STORAGE_KEY = "companies-history-v1";
const TEAM_ROSTER_KEY = "team-roster-v1";

const LOCAL_OPTIONS = [
  "LUZ-BFU",
  "BFU-ABR",
  "ABR-LPA",
  "LPA-PQR",
  "PQR-PRT",
  "PRT-VCL",
  "VCL-JRG",
  "JRG-VAU",
  "VAU-PRU",
  "PRU-CAI",
  "CAI-FDR",
  "FDR-BFI",
  "BFI-FMO",
  "FMO-BTJ",
  "BTJ-CLP",
  "CLP-VPL",
  "VPL-JUN",
  "JUN-LOV",
];

function LocalSelectField({ value, onChange }: { value: string; onChange: (v: string) => void }) {
  const [isEditing, setIsEditing] = useState(false);

  const isPredefined = LOCAL_OPTIONS.includes(value);

  if (isEditing || (!isPredefined && value.trim() !== "")) {
    return (
      <div className="relative flex items-center gap-1.5 w-full">
        <input
          type="text"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder="Digite ou edite o local..."
          className="input pr-16"
          autoFocus
        />
        <button
          type="button"
          onClick={() => setIsEditing(false)}
          className="absolute right-1.5 px-2 py-1 text-xs font-medium text-muted-foreground hover:text-foreground bg-secondary hover:bg-secondary/80 rounded transition-colors"
          title="Ver lista de opções"
        >
          Opções
        </button>
      </div>
    );
  }

  return (
    <div className="flex items-center gap-1.5 w-full">
      <select
        value={value}
        onChange={(e) => {
          if (e.target.value === "__EDIT_MANUALLY__") {
            setIsEditing(true);
          } else {
            onChange(e.target.value);
          }
        }}
        className="input cursor-pointer flex-1"
      >
        <option value="">-- Selecione o Local --</option>
        {LOCAL_OPTIONS.map((opt) => (
          <option key={opt} value={opt}>
            {opt}
          </option>
        ))}
        <option value="__EDIT_MANUALLY__">✏️ Editar texto manualmente...</option>
      </select>
      {value && (
        <button
          type="button"
          onClick={() => setIsEditing(true)}
          className="p-2 border border-border rounded-md bg-background text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors shrink-0"
          title="Clique para editar o texto do item escolhido"
          aria-label="Editar texto do local"
        >
          <Pencil className="h-4 w-4" />
        </button>
      )}
    </div>
  );
}

function ActivitySelectField({
  value,
  onChange,
  activitiesHistory,
  onSaveHistory,
}: {
  value: string;
  onChange: (v: string) => void;
  activitiesHistory: string[];
  onSaveHistory: (actName: string) => void;
}) {
  const [isEditingManually, setIsEditingManually] = useState(false);

  const options = useMemo(() => {
    const DEFAULT_ACTIVITIES = [
      "Inspeção de Via Permanente",
      "Substituição de Dormentes",
      "Manutenção de AMV",
      "Alinhamento e Nivelamento de Via",
      "Roçada e Limpeza de Faixa de Domínio",
      "Inspeção e Esmerilhamento de Trilhos",
      "Manutenção de Sinalização e Subestação",
    ];
    const combined = [...DEFAULT_ACTIVITIES, ...activitiesHistory];
    if (value && value.trim() && !combined.includes(value.trim())) {
      combined.push(value.trim());
    }
    const set = new Set(combined.map((a) => a.trim()).filter(Boolean));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }, [activitiesHistory, value]);

  const isInList = options.includes(value.trim());

  if (isEditingManually || (!isInList && value.trim() !== "")) {
    return (
      <div className="relative flex items-center gap-1.5 w-full">
        <input
          type="text"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          onBlur={(e) => {
            if (e.target.value.trim()) {
              onSaveHistory(e.target.value.trim());
            }
          }}
          placeholder="Digite o nome da atividade..."
          className="input pr-20"
          autoFocus={isEditingManually}
        />
        <button
          type="button"
          onClick={() => setIsEditingManually(false)}
          className="absolute right-1.5 px-2 py-1 text-xs font-medium text-muted-foreground hover:text-foreground bg-secondary hover:bg-secondary/80 rounded transition-colors"
          title="Ver lista de atividades cadastradas"
        >
          Opções
        </button>
      </div>
    );
  }

  return (
    <div className="flex items-center gap-1.5 w-full">
      <select
        value={value}
        onChange={(e) => {
          if (e.target.value === "__NEW_ACTIVITY__") {
            setIsEditingManually(true);
            onChange("");
          } else {
            onChange(e.target.value);
            if (e.target.value.trim()) {
              onSaveHistory(e.target.value.trim());
            }
          }
        }}
        className="input cursor-pointer flex-1 font-medium"
      >
        <option value="">-- Selecione a atividade ({options.length}) --</option>
        {options.map((act) => (
          <option key={act} value={act}>
            {act}
          </option>
        ))}
        <option value="__NEW_ACTIVITY__">✏️ Digitar outra atividade...</option>
      </select>
      <button
        type="button"
        onClick={() => setIsEditingManually(true)}
        className="p-2 border border-border rounded-md bg-background text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors shrink-0"
        title="Editar ou digitar o nome da atividade manualmente"
        aria-label="Editar nome da atividade"
      >
        <Pencil className="h-4 w-4" />
      </button>
    </div>
  );
}

type State = { header: DailyHeader; events: WorkEvent[]; roster: string[] };

const DEFAULT_ROSTER: string[] = [
  "Alessandro Pires Costa",
  "Cássio Messias da Silva",
  "Elicarlos Oliveira Irmão",
  "Kayque Freitas Fagundes",
  "Leandro Donizete Fontana",
  "Moacir Borges Junior",
];

const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);

const pad2 = (n: number) => n.toString().padStart(2, "0");
const tsToLocalInput = (ts: number) => {
  const d = new Date(ts);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
};
const parseLocalInput = (v: string) => {
  const t = new Date(v).getTime();
  return Number.isFinite(t) ? t : null;
};

function PartesDiariasPage() {
  const [state, setState] = useState<State>({
    header: emptyHeader(),
    events: [],
    roster: DEFAULT_ROSTER,
  });
  const [photos, setPhotos] = useState<ReportPhoto[]>([]);
  const [currentReportId, setCurrentReportId] = useState<string | undefined>(undefined);
  const [lastCloudSyncTime, setLastCloudSyncTime] = useState<number | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isOnline, setIsOnline] = useState<boolean>(
    typeof window !== "undefined" ? navigator.onLine : true,
  );
  const [hasPendingSync, setHasPendingSync] = useState<boolean>(false);
  const currentReportIdRef = useRef<string | undefined>(currentReportId);

  useEffect(() => {
    currentReportIdRef.current = currentReportId;
  }, [currentReportId]);

  const [activeTab, setActiveTab] = useState<"registro" | "banco_nuvem">("registro");
  const [isSapModalOpen, setIsSapModalOpen] = useState(false);
  const [saveNotification, setSaveNotification] = useState<string | null>(null);
  const [pendingKind, setPendingKind] = useState<EventKind | null>(null);
  const [pendingDraft, setPendingDraft] = useState<Partial<WorkEvent>>({});
  const [now, setNow] = useState(Date.now());
  const [newMember, setNewMember] = useState("");
  const [sharedLastDefect, setSharedLastDefect] = useState<string>("");
  const [activitiesHistory, setActivitiesHistory] = useState<string[]>([]);
  const [companiesHistory, setCompaniesHistory] = useState<string[]>([]);
  const [editingDescEventId, setEditingDescEventId] = useState<string | null>(null);
  const [editingDescValue, setEditingDescValue] = useState<string>("");

  const syncToCloud = async (
    overrideEvents?: WorkEvent[],
    overrideHeader?: DailyHeader,
    overridePhotos?: ReportPhoto[],
    customNotification?: string,
  ) => {
    const targetEvents = overrideEvents ?? state.events;
    const targetHeader = overrideHeader ?? state.header;
    const targetPhotos = overridePhotos ?? photos;

    if (targetEvents.length === 0) return;

    // Sempre garante o salvamento local no celular em primeiro lugar
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          ...state,
          events: targetEvents,
          header: targetHeader,
          photos: targetPhotos,
          currentReportId: currentReportIdRef.current || currentReportId,
          hasPendingSync: true,
        }),
      );
    } catch {}

    // Se estiver offline, marca como pendente e avisa o usuário
    if (typeof window !== "undefined" && !navigator.onLine) {
      setIsOnline(false);
      setHasPendingSync(true);
      const notifMsg =
        customNotification ||
        "📱 Modo Offline: Seus dados estão salvos em segurança no celular. Serão enviados à nuvem quando houver conexão.";
      setSaveNotification(notifMsg);
      setTimeout(() => setSaveNotification(null), 4000);
      return;
    }

    setIsSyncing(true);
    try {
      const saved = await saveReportToCloud(
        targetHeader,
        targetEvents,
        currentReportIdRef.current || currentReportId,
        targetPhotos,
      );
      if (saved && saved.id) {
        setCurrentReportId(saved.id);
        currentReportIdRef.current = saved.id;
        setLastCloudSyncTime(Date.now());
        setHasPendingSync(false);
        setIsOnline(true);
        const nowStr = new Date().toLocaleTimeString("pt-BR", {
          hour: "2-digit",
          minute: "2-digit",
          second: "2-digit",
        });
        const notifMsg = customNotification || `☁️ Salvo e sincronizado na nuvem às ${nowStr}`;
        setSaveNotification(notifMsg);
        setTimeout(() => setSaveNotification(null), 4000);
      }
    } catch (err) {
      console.warn("Auto-save cloud warning:", err);
      setHasPendingSync(true);
      setSaveNotification(
        "📱 Sem sinal com a nuvem. Todos os dados continuam salvos com segurança no seu celular!",
      );
      setTimeout(() => setSaveNotification(null), 4000);
    } finally {
      setIsSyncing(false);
    }
  };

  const handleHeaderRefreshSync = async () => {
    // Força salvamento local imediato no celular
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          ...state,
          photos,
          currentReportId,
          hasPendingSync,
        }),
      );
    } catch {}

    if (state.events.length === 0) {
      setSaveNotification(
        "📱 Dados salvos no celular. Registre ao menos um evento para sincronizar na nuvem.",
      );
      setTimeout(() => setSaveNotification(null), 4000);
      return;
    }

    if (typeof window !== "undefined" && !navigator.onLine) {
      setIsOnline(false);
      setHasPendingSync(true);
      setSaveNotification(
        "📱 Sem internet no momento: Dados atualizados e salvos no celular com segurança!",
      );
      setTimeout(() => setSaveNotification(null), 4000);
      return;
    }

    await syncToCloud(
      undefined,
      undefined,
      undefined,
      `☁️ Dados atualizados com sucesso na nuvem às ${new Date().toLocaleTimeString("pt-BR")}!`,
    );
  };

  const handleSaveToCloud = async () => {
    if (state.events.length === 0) {
      alert(
        "Nenhum evento registrado ainda. Registre ao menos um evento antes de salvar o relatório na nuvem.",
      );
      return;
    }
    await syncToCloud(
      undefined,
      undefined,
      undefined,
      `Relatório de ${formatDate(state.header.data)} salvo no banco de dados na nuvem!`,
    );
  };

  const handleDownloadExcel = () => {
    if (state.events.length === 0) {
      alert("Nenhum evento registrado no relatório.");
      return;
    }
    const tempReport: SavedCloudReport = {
      id: `REL-${state.header.data || "hoje"}`,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      data: state.header.data || "",
      nome: state.header.nome || "",
      local: state.header.local || "",
      sinais: state.header.sinais || "",
      linha: state.header.linha || "",
      via: state.header.via || "",
      gda: state.header.gda || "",
      om: state.header.om || "",
      empresa: state.header.empresa || "",
      atividade: state.header.atividade || "",
      colaboradores: state.header.colaboradores || "",
      descritivo: state.header.descritivo || "",
      kmProgramado: state.header.kmProgramado || "",
      executado: state.header.executado || "",
      observacaoGeral: state.header.observacaoGeral || "",
      telefone: state.header.telefone || "",
      events: state.events,
      photos: photos,
      totals: totals(state.events),
    };
    exportSingleReportToExcel(tempReport);
  };

  useEffect(() => {
    try {
      let persistentRoster: string[] = [];
      let rosterWasInStorage = false;
      const rawRoster = localStorage.getItem(TEAM_ROSTER_KEY);
      if (rawRoster !== null) {
        try {
          const parsedRoster = JSON.parse(rawRoster);
          if (Array.isArray(parsedRoster)) {
            persistentRoster = parsedRoster;
            rosterWasInStorage = true;
          }
        } catch {}
      }

      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        const stateRoster = Array.isArray(parsed.roster) ? parsed.roster : [];
        const colabNames =
          typeof parsed.header?.colaboradores === "string"
            ? parsed.header.colaboradores
                .split(";")
                .map((s: string) => s.trim())
                .filter(Boolean)
            : [];

        const finalRoster = rosterWasInStorage
          ? persistentRoster
          : Array.from(new Set([...stateRoster, ...colabNames])).filter(Boolean);

        const loadedHeader = { ...emptyHeader(), ...(parsed.header ?? {}) };
        if (!loadedHeader.linha) {
          loadedHeader.linha = "7";
        }
        setState({
          header: loadedHeader,
          events: parsed.events ?? [],
          roster: finalRoster,
        });
        if (Array.isArray(parsed.photos)) {
          setPhotos(parsed.photos);
        }
        if (parsed.currentReportId) {
          setCurrentReportId(parsed.currentReportId);
          currentReportIdRef.current = parsed.currentReportId;
        }
        if (typeof parsed.hasPendingSync === "boolean") {
          setHasPendingSync(parsed.hasPendingSync);
        }
      } else if (rosterWasInStorage) {
        setState((s) => ({ ...s, roster: persistentRoster }));
      }
    } catch {}
  }, []);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(ACTIVITIES_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) setActivitiesHistory(parsed);
      }
    } catch {}

    const unsub = subscribeCloudReports((reports) => {
      const cloudActivities = reports
        .map((r) => r.atividade?.trim())
        .filter((a): a is string => Boolean(a));

      if (cloudActivities.length > 0) {
        setActivitiesHistory((prev) => {
          const combined = Array.from(new Set([...prev, ...cloudActivities]));
          try {
            localStorage.setItem(ACTIVITIES_STORAGE_KEY, JSON.stringify(combined));
          } catch {}
          return combined;
        });
      }
    });

    return unsub;
  }, []);

  const saveActivityToHistory = (actName: string) => {
    const trimmed = actName.trim();
    if (!trimmed) return;
    setActivitiesHistory((prev) => {
      if (prev.includes(trimmed)) return prev;
      const updated = [...prev, trimmed];
      try {
        localStorage.setItem(ACTIVITIES_STORAGE_KEY, JSON.stringify(updated));
      } catch {}
      return updated;
    });
  };

  const clearActivitiesHistory = () => {
    if (window.confirm("Tem certeza que deseja apagar todo o histórico de atividades?")) {
      setActivitiesHistory([]);
      try {
        localStorage.removeItem(ACTIVITIES_STORAGE_KEY);
      } catch {}
    }
  };

  useEffect(() => {
    try {
      const raw = localStorage.getItem(COMPANIES_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) setCompaniesHistory(parsed);
      }
    } catch {}
  }, []);

  const saveCompanyToHistory = (compName: string) => {
    const trimmed = compName.trim();
    if (!trimmed) return;
    setCompaniesHistory((prev) => {
      if (prev.includes(trimmed)) return prev;
      const updated = [...prev, trimmed];
      try {
        localStorage.setItem(COMPANIES_STORAGE_KEY, JSON.stringify(updated));
      } catch {}
      return updated;
    });
  };

  const clearCompaniesHistory = () => {
    if (window.confirm("Tem certeza que deseja apagar todo o histórico de empresas?")) {
      setCompaniesHistory([]);
      try {
        localStorage.removeItem(COMPANIES_STORAGE_KEY);
      } catch {}
    }
  };

  const clearRosterHistory = () => {
    if (window.confirm("Tem certeza que deseja apagar todo o histórico de colaboradores?")) {
      setState((s) => ({
        ...s,
        roster: [],
        header: { ...s.header, colaboradores: "" },
      }));
      try {
        localStorage.setItem(TEAM_ROSTER_KEY, JSON.stringify([]));
      } catch {}
    }
  };

  useEffect(() => {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({ ...state, photos, currentReportId, hasPendingSync }),
      );
    } catch {
      try {
        localStorage.setItem(
          STORAGE_KEY,
          JSON.stringify({ ...state, photos: [], currentReportId, hasPendingSync }),
        );
      } catch {}
    }
    try {
      localStorage.setItem(TEAM_ROSTER_KEY, JSON.stringify(state.roster));
    } catch {}
  }, [state, photos, currentReportId, hasPendingSync]);

  // Monitora conexão e realiza sincronização automática com a nuvem quando a internet voltar
  useEffect(() => {
    const handleOnline = () => {
      setIsOnline(true);
      if (state.events.length > 0) {
        syncToCloud(
          undefined,
          undefined,
          undefined,
          "📶 Conexão reestabelecida! Todos os dados foram sincronizados com a nuvem.",
        );
      } else {
        setSaveNotification("📶 Conexão de internet reestabelecida.");
        setTimeout(() => setSaveNotification(null), 3000);
      }
    };

    const handleOffline = () => {
      setIsOnline(false);
      setSaveNotification(
        "📱 Conexão com a internet perdida. Seus dados continuam sendo salvos em segurança no celular.",
      );
      setTimeout(() => setSaveNotification(null), 4000);
    };

    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);

    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, [state.events, state.header, photos]);

  // Sincroniza alterações no cabeçalho e fotos com a nuvem automaticamente quando houver eventos registrados
  useEffect(() => {
    if (state.events.length === 0) return;
    const timer = setTimeout(() => {
      syncToCloud();
    }, 1500);
    return () => clearTimeout(timer);
  }, [state.header, photos, state.events.length]);

  useEffect(() => {
    const id = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(id);
  }, []);

  // Sincroniza o "último defeito" com o backend (compartilhado entre usuários)
  useEffect(() => {
    fetchLastDefect()
      .then(setSharedLastDefect)
      .catch(() => {});
    const unsub = subscribeLastDefect(setSharedLastDefect);
    return unsub;
  }, []);

  const t = useMemo(() => totals(state.events), [state.events]);
  const sortedEvents = useMemo(
    () => [...state.events].sort((a, b) => a.timestamp - b.timestamp),
    [state.events],
  );
  const jornadaStarted = state.events.some((e) => e.kind === "inicio_jornada");
  const jornadaEnded = state.events.some((e) => e.kind === "fim_jornada");
  const openActivity =
    sortedEvents.filter((e) => e.kind === "inicio_atividade").length >
    sortedEvents.filter((e) => e.kind === "fim_atividade").length;

  const liveJornada = jornadaStarted && !jornadaEnded ? now - sortedEvents[0].timestamp : t.jornada;

  const registerEvent = (kind: EventKind) => {
    // Todos os eventos abrem o modal para permitir descrição/observações opcionais
    setPendingKind(kind);
    setPendingDraft({});
  };

  const addEvent = (kind: EventKind, draft: Partial<WorkEvent>) => {
    const ev: WorkEvent = { id: uid(), kind, timestamp: Date.now(), ...draft };
    const nextEvents = [...state.events, ev];
    setState((s) => ({
      ...s,
      events: nextEvents,
      header: draft.local && !s.header.local ? { ...s.header, local: draft.local } : s.header,
    }));

    // Salva o nome da atividade e empresa no histórico somente ao finalizar a jornada
    if (kind === "fim_jornada") {
      if (state.header.atividade && state.header.atividade.trim()) {
        saveActivityToHistory(state.header.atividade);
      }
      if (state.header.empresa && state.header.empresa.trim()) {
        saveCompanyToHistory(state.header.empresa);
      }
    }

    // Atualiza o último defeito em tempo real para todos os usuários se informado
    if (draft.ultimoDefeito && draft.ultimoDefeito.trim()) {
      const v = draft.ultimoDefeito.trim();
      setSharedLastDefect(v);
      saveLastDefect(v).catch(() => {});
    }
    setPendingKind(null);
    setPendingDraft({});

    // Sincroniza automaticamente com a nuvem a cada evento adicionado (e no início da jornada)
    const customMsg =
      kind === "inicio_jornada"
        ? "☁️ Jornada iniciada e salva na nuvem com sucesso!"
        : kind === "fim_jornada"
          ? "☁️ Jornada encerrada e relatório salvo na nuvem!"
          : "☁️ Evento registrado e atualizado na nuvem!";
    syncToCloud(nextEvents, undefined, undefined, customMsg);
  };

  const removeEvent = (id: string) => {
    const nextEvents = state.events.filter((e) => e.id !== id);
    setState((s) => ({ ...s, events: nextEvents }));
    if (nextEvents.length > 0) {
      syncToCloud(nextEvents, undefined, undefined, "☁️ Evento removido e atualizado na nuvem!");
    }
  };

  const updateEvent = (id: string, patch: Partial<WorkEvent>) => {
    const nextEvents = state.events.map((e) => (e.id === id ? { ...e, ...patch } : e));
    setState((s) => ({
      ...s,
      events: nextEvents,
    }));
    if (patch.ultimoDefeito && patch.ultimoDefeito.trim()) {
      const v = patch.ultimoDefeito.trim();
      setSharedLastDefect(v);
      saveLastDefect(v).catch(() => {});
    }
    syncToCloud(nextEvents, undefined, undefined, "☁️ Evento atualizado na nuvem!");
  };

  const resetDay = () => {
    if (confirm("Zerar todos os dados deste registro de atividade diária?")) {
      setState((s) => {
        const used = s.header.colaboradores
          .split(";")
          .map((n) => n.trim())
          .filter(Boolean);
        const mergedRoster = Array.from(new Set([...s.roster, ...used]));
        return { header: emptyHeader(), events: [], roster: mergedRoster };
      });
      setPhotos([]);
      setCurrentReportId(undefined);
      currentReportIdRef.current = undefined;
    }
  };

  const updateHeader = (patch: Partial<DailyHeader>) => {
    setState((s) => ({ ...s, header: { ...s.header, ...patch } }));
    if (patch.data && patch.data !== state.header.data) {
      const existing = getLocalCache().find((r) => r.data === patch.data);
      if (existing) {
        setCurrentReportId(existing.id);
        currentReportIdRef.current = existing.id;
      } else {
        setCurrentReportId(undefined);
        currentReportIdRef.current = undefined;
      }
    }
  };

  const report = { header: state.header, events: state.events };

  const sendWhatsApp = () => {
    const msg = encodeURIComponent(buildWhatsAppMessage(report));
    const phone = state.header.telefone.replace(/\D/g, "");
    const url = phone ? `https://wa.me/${phone}?text=${msg}` : `https://wa.me/?text=${msg}`;
    window.open(url, "_blank");
  };

  const csvBlob = () => new Blob([buildCSV(report)], { type: "text/csv;charset=utf-8;" });

  const csvFilename = () => `parte-diaria-${state.header.data || "hoje"}.csv`;
  const pdfFilename = () => generatePdfFilename(state.header, state.events);

  const downloadBlob = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const downloadCSV = () => downloadBlob(csvBlob(), csvFilename());

  const buildPdfBlob = async (): Promise<Blob> => {
    return buildPdfBlobForReport(state.header, state.events, photos);
  };

  const downloadPDF = async () => {
    downloadReportPdf(state.header, state.events, photos, pdfFilename());
  };

  const shareFileWhatsApp = async (blob: Blob, filename: string, mime: string) => {
    const file = new File([blob], filename, { type: mime });
    // Preferir Web Share API com arquivo (funciona em mobile e permite anexo)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const nav = navigator as any;
    if (nav.canShare && nav.canShare({ files: [file] })) {
      try {
        await nav.share({
          files: [file],
          title: "Registro de Atividade Diária",
          text: buildWhatsAppMessage(report),
        });
        return;
      } catch {
        // usuário cancelou ou falhou — cai no fallback
      }
    }
    // Fallback: baixa o arquivo e abre o WhatsApp com a mensagem para anexar manualmente
    downloadBlob(blob, filename);
    const msg = encodeURIComponent(
      buildWhatsAppMessage(report) + `\n\n(Arquivo ${filename} baixado — anexe no WhatsApp)`,
    );
    const phone = state.header.telefone.replace(/\D/g, "");
    const url = phone ? `https://wa.me/${phone}?text=${msg}` : `https://wa.me/?text=${msg}`;
    window.open(url, "_blank");
  };

  const sharePdfWhatsApp = async () =>
    shareFileWhatsApp(await buildPdfBlob(), pdfFilename(), "application/pdf");

  const shareExcelWhatsApp = () => {
    const tempReport: SavedCloudReport = {
      id: currentReportId || `REL-${state.header.data || "hoje"}`,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      data: state.header.data || "",
      nome: state.header.nome || "",
      local: state.header.local || "",
      sinais: state.header.sinais || "",
      linha: state.header.linha || "",
      via: state.header.via || "",
      gda: state.header.gda || "",
      om: state.header.om || "",
      empresa: state.header.empresa || "",
      atividade: state.header.atividade || "",
      colaboradores: state.header.colaboradores || "",
      descritivo: state.header.descritivo || "",
      kmProgramado: state.header.kmProgramado || "",
      executado: state.header.executado || "",
      observacaoGeral: state.header.observacaoGeral || "",
      telefone: state.header.telefone || "",
      events: state.events,
      photos: photos,
      totals: totals(state.events),
    };
    const { blob, filename } = buildSingleReportExcelBlob(tempReport);
    return shareFileWhatsApp(
      blob,
      filename,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
  };

  return (
    <main className="min-h-screen bg-background text-foreground pb-12">
      {saveNotification && (
        <div className="fixed top-4 right-4 z-50 flex items-center gap-2 rounded-lg bg-emerald-600 px-4 py-3 text-sm font-medium text-white shadow-xl animate-in fade-in slide-in-from-top-2">
          <Check className="h-5 w-5 shrink-0" />
          <span>{saveNotification}</span>
        </div>
      )}

      <header className="border-b border-border bg-card">
        <div className="mx-auto max-w-5xl px-4 py-4 sm:py-5">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div>
              <h1 className="text-2xl font-semibold tracking-tight sm:text-3xl">
                Registro de Atividade Diária
              </h1>
              <p className="mt-0.5 text-xs sm:text-sm text-muted-foreground">
                Tempos e movimentos do dia — registre eventos, salve na nuvem, gere relatórios em
                Excel e PDF.
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2 self-start sm:self-auto">
              <div className="flex bg-secondary p-1 rounded-lg border border-border">
                <button
                  type="button"
                  onClick={() => setActiveTab("registro")}
                  className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1.5 ${
                    activeTab === "registro"
                      ? "bg-card text-foreground shadow-sm"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                >
                  <FileText className="h-3.5 w-3.5" />
                  <span>Registro de Hoje</span>
                </button>

                <button
                  type="button"
                  onClick={() => setActiveTab("banco_nuvem")}
                  className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1.5 ${
                    activeTab === "banco_nuvem"
                      ? "bg-primary text-white shadow-sm"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                >
                  <Cloud className="h-3.5 w-3.5" />
                  <span>Banco de Dados na Nuvem</span>
                </button>
              </div>

              <button
                type="button"
                onClick={handleHeaderRefreshSync}
                disabled={isSyncing}
                className={`flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg border transition-all shadow-sm ${
                  !isOnline
                    ? "bg-amber-500/10 text-amber-700 border-amber-500/40 hover:bg-amber-500/20 dark:text-amber-300"
                    : hasPendingSync
                      ? "bg-amber-500/10 text-amber-700 border-amber-500/40 hover:bg-amber-500/20 dark:text-amber-300"
                      : "bg-emerald-500/10 text-emerald-700 border-emerald-500/40 hover:bg-emerald-500/20 dark:text-emerald-300"
                }`}
                title={
                  !isOnline
                    ? "Sem internet - Dados salvos no celular. Clique para tentar atualizar."
                    : hasPendingSync
                      ? "Internet disponível - Há dados pendentes para sincronizar com a nuvem."
                      : "Sincronizado com a nuvem - Clique para atualizar."
                }
              >
                <RefreshCw
                  className={`h-3.5 w-3.5 ${isSyncing ? "animate-spin text-primary" : ""}`}
                />
                <span>
                  {isSyncing
                    ? "Sincronizando..."
                    : !isOnline
                      ? "Salvo no celular (Offline)"
                      : hasPendingSync
                        ? "Sincronizar Pendente"
                        : "Atualizar na Nuvem"}
                </span>
                {isOnline ? (
                  <Wifi className="h-3.5 w-3.5 text-emerald-600 dark:text-emerald-400" />
                ) : (
                  <WifiOff className="h-3.5 w-3.5 text-amber-600 dark:text-amber-400" />
                )}
              </button>

              <Link
                to="/checklist"
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-primary/10 text-primary hover:bg-primary/20 border border-primary/30 transition-all shadow-xs"
                title="Abrir checklist de inspeção de máquinas e equipamentos"
              >
                <ClipboardCheck className="h-3.5 w-3.5" />
                <span>Checklist do Equipamento</span>
              </Link>

              <Link to="/dashboard" className="btn-secondary text-xs py-2 px-3">
                Dashboard
              </Link>

              <ThemeToggle className="py-2 px-3" />

              {activeTab === "registro" && (
                <div className="hidden text-right sm:block ml-2 border-l border-border pl-3">
                  <div className="text-[10px] uppercase tracking-wider text-muted-foreground">
                    Jornada
                  </div>
                  <div className="font-mono text-xl tabular-nums">
                    {formatDuration(liveJornada)}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-5xl px-4 py-6">
        {activeTab === "banco_nuvem" ? (
          <CloudReportsDatabase
            onLoadReportToForm={(r) => {
              const { header, events, photos: cloudPhotos } = cloudReportToHeaderAndEvents(r);
              setState((s) => ({ ...s, header, events }));
              setPhotos(cloudPhotos || []);
              setCurrentReportId(r.id);
              currentReportIdRef.current = r.id;
              setActiveTab("registro");
              setSaveNotification(`☁️ Relatório de ${formatDate(header.data)} carregado da nuvem!`);
              setTimeout(() => setSaveNotification(null), 4000);
            }}
          />
        ) : (
          <div className="grid gap-6 lg:grid-cols-3">
            {/* Left: Header form + Actions */}
            <section className="space-y-6 lg:col-span-1">
              <Card title="Cabeçalho">
                <div className="grid grid-cols-2 gap-3">
                  <Field label="Data">
                    <input
                      type="date"
                      value={state.header.data}
                      onChange={(e) => updateHeader({ data: e.target.value })}
                      className="input"
                    />
                  </Field>
                  <Field label="Responsável">
                    <input
                      value={state.header.nome}
                      onChange={(e) => updateHeader({ nome: e.target.value })}
                      className="input"
                    />
                  </Field>
                  <Field label="Local" full>
                    <LocalSelectField
                      value={state.header.local}
                      onChange={(local) => updateHeader({ local })}
                    />
                  </Field>
                  <Field label="Sinais" full>
                    <input
                      value={state.header.sinais}
                      onChange={(e) => updateHeader({ sinais: maskSinais(e.target.value) })}
                      placeholder="Ex: 23BFU-24ABR"
                      className="input"
                    />
                  </Field>
                  <Field label="Linha">
                    <input
                      value={state.header.linha}
                      onChange={(e) => updateHeader({ linha: e.target.value })}
                      className="input"
                    />
                  </Field>
                  <Field label="Via">
                    <input
                      value={state.header.via}
                      onChange={(e) => updateHeader({ via: e.target.value })}
                      className="input"
                    />
                  </Field>
                  <Field label="GDA">
                    <input
                      value={state.header.gda}
                      onChange={(e) => updateHeader({ gda: maskGda(e.target.value) })}
                      placeholder="Ex: 2026-001235"
                      className="input"
                    />
                  </Field>
                  <Field label="OM">
                    <input
                      value={state.header.om}
                      onChange={(e) => updateHeader({ om: e.target.value })}
                      className="input"
                    />
                  </Field>
                  <Field
                    label="Empresa"
                    full
                    action={
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <button
                            type="button"
                            className="rounded p-0.5 text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors"
                            title="Opções do histórico"
                            aria-label="Opções do histórico"
                          >
                            <MoreVertical className="h-4 w-4" />
                          </button>
                        </DropdownMenuTrigger>
                        <DropdownMenuContent align="end">
                          <DropdownMenuItem
                            className="text-destructive focus:text-destructive flex cursor-pointer items-center gap-2 text-xs"
                            onClick={clearCompaniesHistory}
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                            Apagar histórico de empresas
                          </DropdownMenuItem>
                        </DropdownMenuContent>
                      </DropdownMenu>
                    }
                  >
                    <div>
                      <input
                        list="companies-history-list"
                        value={state.header.empresa}
                        onChange={(e) => updateHeader({ empresa: e.target.value })}
                        onBlur={(e) => {
                          if (e.target.value.trim()) {
                            saveCompanyToHistory(e.target.value);
                          }
                        }}
                        className="input"
                      />
                      <datalist id="companies-history-list">
                        {companiesHistory.map((comp) => (
                          <option key={comp} value={comp} />
                        ))}
                      </datalist>
                    </div>
                  </Field>
                  <Field
                    label="Nome da Atividade"
                    full
                    action={
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <button
                            type="button"
                            className="rounded p-0.5 text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors"
                            title="Opções do histórico"
                            aria-label="Opções do histórico"
                          >
                            <MoreVertical className="h-4 w-4" />
                          </button>
                        </DropdownMenuTrigger>
                        <DropdownMenuContent align="end">
                          <DropdownMenuItem
                            className="text-destructive focus:text-destructive flex cursor-pointer items-center gap-2 text-xs"
                            onClick={clearActivitiesHistory}
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                            Apagar histórico de atividades
                          </DropdownMenuItem>
                        </DropdownMenuContent>
                      </DropdownMenu>
                    }
                  >
                    <ActivitySelectField
                      value={state.header.atividade}
                      onChange={(atividade) => updateHeader({ atividade })}
                      activitiesHistory={activitiesHistory}
                      onSaveHistory={saveActivityToHistory}
                    />
                  </Field>
                  <Field
                    label="Equipe / Colaboradores"
                    full
                    action={
                      state.roster.length > 0 ? (
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <button
                              type="button"
                              className="rounded p-0.5 text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors"
                              title="Opções da equipe"
                              aria-label="Opções da equipe"
                            >
                              <MoreVertical className="h-4 w-4" />
                            </button>
                          </DropdownMenuTrigger>
                          <DropdownMenuContent align="end">
                            <DropdownMenuItem
                              className="text-destructive focus:text-destructive flex cursor-pointer items-center gap-2 text-xs"
                              onClick={clearRosterHistory}
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                              Apagar histórico de colaboradores
                            </DropdownMenuItem>
                          </DropdownMenuContent>
                        </DropdownMenu>
                      ) : null
                    }
                  >
                    <TeamPicker
                      roster={state.roster}
                      selected={state.header.colaboradores}
                      onSelectedChange={(v) => updateHeader({ colaboradores: v })}
                      onRosterChange={(roster) => setState((s) => ({ ...s, roster }))}
                      newMember={newMember}
                      setNewMember={setNewMember}
                    />
                  </Field>
                  <Field label="Descritivo da Atividade" full>
                    <textarea
                      value={state.header.descritivo}
                      onChange={(e) => updateHeader({ descritivo: e.target.value })}
                      className="input min-h-16"
                    />
                  </Field>
                  <Field label="Programado" full>
                    <input
                      value={state.header.kmProgramado}
                      onChange={(e) =>
                        updateHeader({ kmProgramado: maskProgramadoExecutado(e.target.value) })
                      }
                      placeholder="Ex: 25/21-26/21"
                      className="input"
                    />
                  </Field>
                  <Field
                    label="Executado"
                    full
                    action={
                      state.header.kmProgramado ? (
                        <button
                          type="button"
                          onClick={() => updateHeader({ executado: state.header.kmProgramado })}
                          className="text-xs font-medium text-primary hover:underline flex items-center gap-1 transition-colors"
                          title="Copiar mesmo valor do Programado"
                        >
                          <Copy className="h-3 w-3" />
                          <span>Mesmo do programado</span>
                        </button>
                      ) : null
                    }
                  >
                    <div className="space-y-1.5">
                      <input
                        value={state.header.executado}
                        onChange={(e) =>
                          updateHeader({ executado: maskProgramadoExecutado(e.target.value) })
                        }
                        placeholder="Ex: 25/21-26/21"
                        className="input"
                      />
                      {state.header.kmProgramado &&
                        state.header.executado !== state.header.kmProgramado && (
                          <button
                            type="button"
                            onClick={() => updateHeader({ executado: state.header.kmProgramado })}
                            className="inline-flex items-center gap-1 text-xs text-muted-foreground hover:text-primary transition-colors"
                          >
                            <Copy className="h-3 w-3" />
                            <span>Usar valor do programado ({state.header.kmProgramado})</span>
                          </button>
                        )}
                    </div>
                  </Field>
                  <Field label="Observação" full>
                    <textarea
                      value={state.header.observacaoGeral}
                      onChange={(e) => updateHeader({ observacaoGeral: e.target.value })}
                      className="input min-h-16"
                    />
                  </Field>
                </div>
              </Card>

              <Card title="Registrar evento">
                <div
                  className={`mb-3 flex items-center justify-between rounded-lg border px-3 py-2 text-xs font-medium ${
                    !isOnline
                      ? "border-amber-500/30 bg-amber-500/10 text-amber-800 dark:text-amber-200"
                      : "border-emerald-500/30 bg-emerald-500/10 text-emerald-800 dark:text-emerald-200"
                  }`}
                >
                  <div className="flex items-center gap-2">
                    {isOnline ? (
                      <Cloud
                        className={`h-4 w-4 shrink-0 text-emerald-600 dark:text-emerald-400 ${
                          isSyncing ? "animate-bounce" : ""
                        }`}
                      />
                    ) : (
                      <WifiOff className="h-4 w-4 shrink-0 text-amber-600 dark:text-amber-400" />
                    )}
                    <div>
                      <span className="font-semibold block">
                        {isOnline
                          ? "Salvo na nuvem e no celular"
                          : "Salvo no celular (Modo Offline)"}
                      </span>
                      <span className="text-[10px] text-muted-foreground block">
                        {isSyncing
                          ? "Sincronizando agora..."
                          : !isOnline
                            ? "Sem internet — enviará para a nuvem quando reconectar"
                            : lastCloudSyncTime
                              ? `Última sincronização às ${new Date(
                                  lastCloudSyncTime,
                                ).toLocaleTimeString("pt-BR")}`
                              : "Ativo — dados seguros no celular"}
                      </span>
                    </div>
                  </div>
                  <span
                    className={`inline-flex items-center rounded-full px-2 py-0.5 text-[10px] font-semibold ${
                      !isOnline
                        ? "bg-amber-500/20 text-amber-700 dark:text-amber-300"
                        : "bg-emerald-500/20 text-emerald-700 dark:text-emerald-300"
                    }`}
                  >
                    {!isOnline ? "Offline 📱" : "Nuvem Ativa ⚡"}
                  </span>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <EventBtn
                    onClick={() => registerEvent("inicio_jornada")}
                    disabled={jornadaStarted}
                    variant="primary"
                  >
                    Início da jornada
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("fim_jornada")}
                    disabled={!jornadaStarted || jornadaEnded || openActivity}
                    variant="danger"
                  >
                    Fim da jornada
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("preparacao")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Preparação/Manutenção
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("parada")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Parada
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("inicio_atividade")}
                    disabled={!jornadaStarted || jornadaEnded || openActivity}
                    variant="success"
                  >
                    Início da atividade
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("fim_atividade")}
                    disabled={!openActivity}
                    variant="warning"
                  >
                    Fim da atividade
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("deslocamento_ferroviario")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Deslocamento Ferroviário
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("deslocamento_rodoviario")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Deslocamento Rodoviário
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("pronto_cco")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Pronto pra sair ao CCO
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("pronto_patio")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Pronto pra sair no pátio
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("pedido_cco")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Pedido ao CCO
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("liberacao_acesso")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Liberação de acesso a via
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("chegada_local")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Chegada ao Local
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("retorno_base")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Retorno a base
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("dimensionamento_defeito")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Dimensionamento Defeito
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("refeicao")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Janta/Almoço
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("cancelado")}
                    disabled={!jornadaStarted || jornadaEnded}
                    variant="danger"
                  >
                    Cancelado
                  </EventBtn>
                  <EventBtn
                    onClick={() => registerEvent("outros")}
                    disabled={!jornadaStarted || jornadaEnded}
                  >
                    Outros...
                  </EventBtn>
                </div>
                {jornadaEnded && (
                  <p className="mt-3 rounded-md bg-accent/60 px-3 py-2 text-xs text-accent-foreground">
                    Jornada encerrada. Gere o relatório abaixo.
                  </p>
                )}
              </Card>

              <PhotoUploader photos={photos} onChange={setPhotos} />

              <Card title="Relatório">
                <div className="flex flex-col gap-2">
                  <button
                    type="button"
                    onClick={handleSaveToCloud}
                    disabled={state.events.length === 0}
                    className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-medium py-2.5 px-3 rounded-md transition-colors flex items-center justify-center gap-2 shadow-sm disabled:opacity-50 text-xs sm:text-sm"
                    title="Salvar este relatório no banco de dados na nuvem para acesso de toda a equipe"
                  >
                    <Cloud className="h-4 w-4" />
                    <span>Salvar Relatório na Nuvem</span>
                  </button>

                  <button
                    onClick={sendWhatsApp}
                    className="btn-primary text-xs sm:text-sm"
                    disabled={state.events.length === 0}
                  >
                    Enviar resumo por WhatsApp
                  </button>

                  <div className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground mt-2">
                    Apropriação SAP & Excel
                  </div>

                  {/* Botão Oficial SAP IW41 */}
                  <button
                    type="button"
                    onClick={() => setIsSapModalOpen(true)}
                    className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2.5 px-3 rounded-md transition-colors flex items-center justify-between gap-2 shadow-xs text-xs sm:text-sm border border-emerald-500"
                    title="Gerar arquivo em Excel (.xlsm) com Macro integrada para apropriação no SAP via transação IW41"
                  >
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet className="h-4 w-4 shrink-0" />
                      <span>Apropriação SAP (IW41)</span>
                    </div>
                    <span className="bg-black/25 text-[10px] font-mono px-1.5 py-0.5 rounded text-white/90 shrink-0">
                      .xlsm + Macro
                    </span>
                  </button>

                  <div className="grid grid-cols-2 gap-2">
                    <button
                      type="button"
                      onClick={handleDownloadExcel}
                      className="btn-secondary text-xs text-emerald-700 dark:text-emerald-400 font-semibold border-emerald-500/30 flex items-center justify-center gap-1"
                      disabled={state.events.length === 0}
                      title="Baixar relatório diário em formato de planilha Excel (.xlsx)"
                    >
                      <FileSpreadsheet className="h-3.5 w-3.5" />
                      <span>Baixar Excel</span>
                    </button>
                    <button
                      type="button"
                      onClick={shareExcelWhatsApp}
                      className="btn-secondary text-xs text-emerald-700 dark:text-emerald-400 font-semibold border-emerald-500/30 flex items-center justify-center gap-1.5"
                      disabled={state.events.length === 0}
                      title="Compartilhar ou anexar relatório em planilha Excel via WhatsApp"
                    >
                      <FileSpreadsheet className="h-3.5 w-3.5" />
                      <span>Excel WhatsApp</span>
                    </button>
                  </div>

                  <div className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground mt-2">
                    PDF
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      onClick={downloadPDF}
                      className="btn-secondary text-xs"
                      disabled={state.events.length === 0}
                    >
                      Baixar PDF
                    </button>
                    <button
                      onClick={sharePdfWhatsApp}
                      className="btn-secondary text-xs"
                      disabled={state.events.length === 0}
                    >
                      PDF via WhatsApp
                    </button>
                  </div>

                  <button onClick={resetDay} className="btn-ghost mt-2 text-destructive text-xs">
                    Zerar dia
                  </button>
                </div>
              </Card>
            </section>

            {/* Right: Totals + timeline */}
            <section className="space-y-6 lg:col-span-2">
              <LastDefectCard
                value={sharedLastDefect}
                onSave={(val) => {
                  const v = val.trim();
                  setSharedLastDefect(v);
                  saveLastDefect(v).catch(() => {});
                }}
              />

              <Card title="Resumo">
                <div className="grid grid-cols-2 gap-3 sm:grid-cols-3">
                  <Stat label="Jornada" value={formatDuration(liveJornada)} tone="primary" />
                  <Stat label="Atividade" value={formatDuration(t.atividade)} tone="success" />
                  <Stat label="Preparação" value={formatDuration(t.preparacao)} />
                  <Stat label="Paradas" value={formatDuration(t.parada)} tone="warning" />
                  <Stat label="Produção (m)" value={String(t.producaoTotal || 0)} />
                  <Stat label="Defeitos" value={String(t.defeitosTotal || 0)} />
                </div>
              </Card>

              <Card title="Linha do tempo">
                {sortedEvents.length === 0 ? (
                  <p className="text-sm text-muted-foreground">
                    Nenhum evento registrado ainda. Comece pelo "Início da jornada".
                  </p>
                ) : (
                  <ol className="space-y-3">
                    {sortedEvents.map((e, i) => {
                      const startTs = e.timestamp;
                      const next = sortedEvents[i + 1];
                      const hasSubsequent = Boolean(next);

                      let endStr = "";
                      let durStr = "";
                      let isOngoing = false;

                      if (hasSubsequent) {
                        const endTs = next.timestamp;
                        const durMs = Math.max(0, endTs - startTs);
                        endStr = formatTime(endTs);
                        durStr = formatDuration(durMs);
                      } else if (e.kind === "fim_jornada") {
                        endStr = formatTime(startTs);
                        durStr = "00:00:00";
                      } else {
                        isOngoing = true;
                        const durMs = Math.max(0, now - startTs);
                        endStr = "Em andamento";
                        durStr = `${formatDuration(durMs)} (em andamento)`;
                      }

                      return (
                        <li
                          key={e.id}
                          className="flex flex-col gap-2 rounded-lg border border-border bg-card p-3 shadow-xs"
                        >
                          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-border/50 pb-2">
                            <div className="flex flex-wrap items-center gap-2">
                              <span className={`badge badge-${e.kind}`}>{EVENT_LABEL[e.kind]}</span>
                              {isOngoing ? (
                                <span className="inline-flex items-center gap-1.5 rounded-full bg-amber-500/10 px-2.5 py-0.5 text-xs font-semibold text-amber-600 dark:text-amber-400 border border-amber-500/20 animate-pulse">
                                  <span className="h-2 w-2 rounded-full bg-amber-500"></span>
                                  <span>Duração: {durStr}</span>
                                </span>
                              ) : (
                                <span className="text-xs font-mono font-medium text-muted-foreground bg-secondary/80 px-2.5 py-0.5 rounded-md border border-border/40">
                                  Duração:{" "}
                                  <strong className="text-foreground font-bold">{durStr}</strong>
                                </span>
                              )}
                            </div>

                            <div className="flex items-center gap-2 text-xs">
                              <button
                                type="button"
                                onClick={() => {
                                  setEditingDescEventId(e.id);
                                  setEditingDescValue(e.descricao || "");
                                }}
                                className="inline-flex items-center gap-1 text-muted-foreground hover:text-primary transition-colors"
                                title="Editar descrição do evento"
                              >
                                <Pencil className="h-3 w-3" />
                                <span>{e.descricao ? "Editar desc." : "+ Descrição"}</span>
                              </button>
                              <span className="text-muted-foreground">•</span>
                              <button
                                type="button"
                                onClick={() => removeEvent(e.id)}
                                className="text-muted-foreground hover:text-destructive transition-colors"
                              >
                                remover
                              </button>
                            </div>
                          </div>

                          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs pt-1">
                            {/* Horários: Início e Fim */}
                            <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5 bg-secondary/30 p-2 rounded-md border border-border/40">
                              <div className="flex items-center gap-1.5">
                                <span className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider">
                                  Início:
                                </span>
                                <input
                                  type="datetime-local"
                                  step={1}
                                  value={tsToLocalInput(e.timestamp)}
                                  onChange={(ev) => {
                                    const ts = parseLocalInput(ev.target.value);
                                    if (ts) updateEvent(e.id, { timestamp: ts });
                                  }}
                                  className="input px-1.5 py-0.5 text-xs font-mono w-auto"
                                />
                              </div>
                              <div className="flex items-center gap-1.5">
                                <span className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider">
                                  Fim:
                                </span>
                                <span className="font-mono font-semibold text-foreground bg-background px-2 py-0.5 rounded border border-border/60">
                                  {endStr}
                                </span>
                              </div>
                            </div>

                            {/* Detalhes do Evento */}
                            <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-xs text-muted-foreground self-center">
                              {e.local && (
                                <span>
                                  local: <b className="text-foreground">{e.local}</b>
                                </span>
                              )}
                              {e.kmInicial && (
                                <span>
                                  km inicial: <b className="text-foreground">{e.kmInicial}</b>
                                </span>
                              )}
                              {e.kmFinal && (
                                <span>
                                  km final: <b className="text-foreground">{e.kmFinal}</b>
                                </span>
                              )}
                              {e.producao && (
                                <span>
                                  produção: <b className="text-foreground">{e.producao} m</b>
                                </span>
                              )}
                              {e.defeitos && (
                                <span>
                                  defeitos: <b className="text-foreground">{e.defeitos}</b>
                                </span>
                              )}
                              {e.ultimoDefeito && (
                                <span>
                                  último def.: <b className="text-foreground">{e.ultimoDefeito}</b>
                                </span>
                              )}
                              {e.descricao && (
                                <span className="basis-full">
                                  descrição: <b className="text-foreground">{e.descricao}</b>
                                </span>
                              )}
                              {e.observacoes && (
                                <span className="basis-full">obs: {e.observacoes}</span>
                              )}
                            </div>
                          </div>

                          {editingDescEventId === e.id && (
                            <div className="mt-1 flex flex-col gap-2 rounded-md border border-primary/30 bg-primary/5 p-2.5">
                              <div className="flex items-center justify-between">
                                <span className="text-xs font-semibold text-foreground">
                                  Editar Descrição do Evento
                                </span>
                                <span className="text-[10px] text-muted-foreground">
                                  {EVENT_LABEL[e.kind]}
                                </span>
                              </div>
                              <textarea
                                value={editingDescValue}
                                onChange={(ev) => setEditingDescValue(ev.target.value)}
                                className="input min-h-16 text-xs bg-background"
                                placeholder="Digite a descrição do evento..."
                                autoFocus
                              />
                              <div className="flex justify-end gap-2">
                                <button
                                  type="button"
                                  onClick={() => setEditingDescEventId(null)}
                                  className="btn-ghost py-1 px-2.5 text-xs"
                                >
                                  Cancelar
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    updateEvent(e.id, { descricao: editingDescValue.trim() });
                                    setEditingDescEventId(null);
                                  }}
                                  className="btn-primary py-1 px-3 text-xs"
                                >
                                  Salvar Descrição
                                </button>
                              </div>
                            </div>
                          )}
                        </li>
                      );
                    })}
                  </ol>
                )}
              </Card>
            </section>
          </div>
        )}
      </div>

      {pendingKind && (
        <ActivityModal
          kind={pendingKind}
          draft={pendingDraft}
          onChange={setPendingDraft}
          onCancel={() => {
            setPendingKind(null);
            setPendingDraft({});
          }}
          onConfirm={() => addEvent(pendingKind, pendingDraft)}
        />
      )}

      {isSapModalOpen && (
        <SapIw41Modal
          isOpen={isSapModalOpen}
          onClose={() => setIsSapModalOpen(false)}
          header={state.header}
          events={state.events}
          onSuccess={(filename) =>
            setSaveNotification(`✓ Planilha SAP IW41 gerada com sucesso: ${filename}`)
          }
        />
      )}

      {saveNotification && (
        <div className="fixed bottom-5 right-5 z-50 flex items-center gap-3 rounded-xl bg-emerald-600 px-4 py-3 text-xs sm:text-sm font-semibold text-white shadow-2xl transition-all animate-in fade-in slide-in-from-bottom-4 border border-emerald-400/40">
          <Cloud className="h-4 w-4 shrink-0 animate-pulse text-white" />
          <span>{saveNotification}</span>
        </div>
      )}
    </main>
  );
}

function Card({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="rounded-xl border border-border bg-card p-4 shadow-sm">
      <h2 className="mb-3 text-sm font-semibold uppercase tracking-wider text-muted-foreground">
        {title}
      </h2>
      {children}
    </div>
  );
}

function Field({
  label,
  full,
  action,
  children,
}: {
  label: string;
  full?: boolean;
  action?: React.ReactNode;
  children: React.ReactNode;
}) {
  return (
    <div
      className={`flex flex-col gap-1 text-xs font-medium text-muted-foreground ${full ? "col-span-2" : ""}`}
    >
      <div className="flex items-center justify-between min-h-[20px]">
        <span>{label}</span>
        {action}
      </div>
      {children}
    </div>
  );
}

function Stat({
  label,
  value,
  tone,
}: {
  label: string;
  value: string;
  tone?: "primary" | "success" | "warning";
}) {
  const toneClass =
    tone === "primary"
      ? "text-primary"
      : tone === "success"
        ? "text-[color:var(--success)]"
        : tone === "warning"
          ? "text-[color:var(--warning)]"
          : "text-foreground";
  return (
    <div className="rounded-lg border border-border bg-background px-3 py-3">
      <div className="text-[10px] font-semibold uppercase tracking-widest text-muted-foreground">
        {label}
      </div>
      <div className={`mt-1 font-mono text-xl tabular-nums ${toneClass}`}>{value}</div>
    </div>
  );
}

function EventBtn({
  children,
  onClick,
  disabled,
  variant,
}: {
  children: React.ReactNode;
  onClick: () => void;
  disabled?: boolean;
  variant?: "primary" | "success" | "warning" | "danger";
}) {
  const base =
    "rounded-lg px-3 py-3 text-sm font-medium transition-colors disabled:cursor-not-allowed disabled:opacity-40";
  const styles =
    variant === "primary"
      ? "bg-primary text-primary-foreground hover:bg-primary/90"
      : variant === "success"
        ? "bg-[color:var(--success)] text-[color:var(--success-foreground)] hover:opacity-90"
        : variant === "warning"
          ? "bg-[color:var(--warning)] text-[color:var(--warning-foreground)] hover:opacity-90"
          : variant === "danger"
            ? "bg-destructive text-destructive-foreground hover:bg-destructive/90"
            : "bg-secondary text-secondary-foreground hover:bg-secondary/80";
  return (
    <button onClick={onClick} disabled={disabled} className={`${base} ${styles}`}>
      {children}
    </button>
  );
}

function ActivityModal({
  kind,
  draft,
  onChange,
  onCancel,
  onConfirm,
}: {
  kind: EventKind;
  draft: Partial<WorkEvent>;
  onChange: (d: Partial<WorkEvent>) => void;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  const [errorMsg, setErrorMsg] = useState("");
  const isStart = kind === "inicio_atividade";
  const isEnd = kind === "fim_atividade";
  const isDefect = kind === "dimensionamento_defeito";
  const isCancelado = kind === "cancelado";
  const isProntoPatio = kind === "pronto_patio";

  const descLabel =
    kind === "cancelado"
      ? "Motivo do cancelamento *"
      : kind === "parada"
        ? "Motivo da parada"
        : kind === "outros"
          ? "Descreva o evento"
          : "Descrição (opcional)";

  const handleConfirm = () => {
    if (isCancelado && (!draft.descricao || !draft.descricao.trim())) {
      setErrorMsg("É obrigatório informar o motivo do cancelamento.");
      return;
    }
    if (isProntoPatio && (!draft.local || !draft.local.trim())) {
      setErrorMsg("É obrigatório informar o local ao acionar Pronto pra sair no pátio.");
      return;
    }
    setErrorMsg("");
    onConfirm();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
      <div className="w-full max-w-md rounded-xl border border-border bg-card p-5 shadow-xl">
        <h3 className="text-lg font-semibold">{EVENT_LABEL[kind]}</h3>
        <p className="mt-1 text-xs text-muted-foreground">
          Hora será registrada como agora: {formatTime(Date.now())}
        </p>
        <div className="mt-4 grid grid-cols-2 gap-3">
          {isStart && (
            <Field label="Km/poste inicial" full>
              <input
                autoFocus
                value={draft.kmInicial ?? ""}
                onChange={(e) => onChange({ ...draft, kmInicial: e.target.value })}
                className="input"
              />
            </Field>
          )}
          {isEnd && (
            <>
              <Field label="Km/poste final" full>
                <input
                  autoFocus
                  value={draft.kmFinal ?? ""}
                  onChange={(e) => onChange({ ...draft, kmFinal: e.target.value })}
                  className="input"
                />
              </Field>
              <Field label="Produção (m)">
                <input
                  inputMode="numeric"
                  value={draft.producao ?? ""}
                  onChange={(e) => onChange({ ...draft, producao: e.target.value })}
                  className="input"
                />
              </Field>
              <Field label="Nº defeitos">
                <input
                  inputMode="numeric"
                  value={draft.defeitos ?? ""}
                  onChange={(e) => onChange({ ...draft, defeitos: e.target.value })}
                  className="input"
                />
              </Field>
              <Field label="Nº último defeito" full>
                <input
                  value={draft.ultimoDefeito ?? ""}
                  onChange={(e) => onChange({ ...draft, ultimoDefeito: e.target.value })}
                  className="input"
                />
              </Field>
            </>
          )}
          {isDefect && (
            <Field label="Nº do novo defeito" full>
              <input
                autoFocus
                value={draft.ultimoDefeito ?? ""}
                onChange={(e) => onChange({ ...draft, ultimoDefeito: e.target.value })}
                className="input"
              />
            </Field>
          )}
          {isProntoPatio && (
            <Field label="Local *" full>
              <input
                autoFocus
                value={draft.local ?? ""}
                onChange={(e) => {
                  if (errorMsg) setErrorMsg("");
                  onChange({ ...draft, local: e.target.value });
                }}
                className={`input ${errorMsg && !draft.local?.trim() ? "border-destructive focus-visible:ring-destructive" : ""}`}
                placeholder="Informe o local do pátio..."
              />
            </Field>
          )}
          <Field label={descLabel} full>
            <textarea
              value={draft.descricao ?? ""}
              onChange={(e) => {
                if (errorMsg) setErrorMsg("");
                onChange({ ...draft, descricao: e.target.value });
              }}
              className={`input min-h-20 ${errorMsg ? "border-destructive focus-visible:ring-destructive" : ""}`}
              placeholder={
                isCancelado ? "Digite o motivo do cancelamento obrigatoriamente..." : undefined
              }
              autoFocus={isCancelado}
            />
            {errorMsg && (
              <span className="text-xs font-medium text-destructive mt-1 block">{errorMsg}</span>
            )}
          </Field>
          <Field label="Observações" full>
            <textarea
              value={draft.observacoes ?? ""}
              onChange={(e) => onChange({ ...draft, observacoes: e.target.value })}
              className="input min-h-20"
            />
          </Field>
        </div>
        <div className="mt-5 flex justify-end gap-2">
          <button onClick={onCancel} className="btn-ghost">
            Cancelar
          </button>
          <button onClick={handleConfirm} className="btn-primary">
            Registrar
          </button>
        </div>
      </div>
    </div>
  );
}

function TeamPicker({
  roster,
  selected,
  onSelectedChange,
  onRosterChange,
  newMember,
  setNewMember,
}: {
  roster: string[];
  selected: string;
  onSelectedChange: (v: string) => void;
  onRosterChange: (v: string[]) => void;
  newMember: string;
  setNewMember: (v: string) => void;
}) {
  const selectedList = selected
    .split(";")
    .map((n) => n.trim())
    .filter(Boolean);

  const isSelected = (n: string) => selectedList.includes(n);

  const toggle = (n: string) => {
    const next = isSelected(n) ? selectedList.filter((x) => x !== n) : [...selectedList, n];
    onSelectedChange(next.join("; "));
  };

  const addMember = () => {
    const name = newMember.trim();
    if (!name) return;
    const updatedRoster = roster.includes(name) ? roster : [...roster, name];
    if (!roster.includes(name)) {
      onRosterChange(updatedRoster);
    }
    if (!selectedList.includes(name)) {
      onSelectedChange([...selectedList, name].join("; "));
    }
    setNewMember("");
  };

  const removeFromRoster = (n: string, e: React.MouseEvent) => {
    e.stopPropagation();
    e.preventDefault();
    onRosterChange(roster.filter((x) => x !== n));
    if (isSelected(n)) {
      onSelectedChange(selectedList.filter((x) => x !== n).join("; "));
    }
  };

  return (
    <div className="space-y-3">
      <div className="flex gap-2">
        <input
          list="roster-datalist"
          value={newMember}
          onChange={(e) => setNewMember(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              addMember();
            }
          }}
          className="input flex-1"
        />
        <datalist id="roster-datalist">
          {roster.map((name) => (
            <option key={name} value={name} />
          ))}
        </datalist>
        <button type="button" onClick={addMember} className="btn-secondary px-3">
          Adicionar
        </button>
      </div>

      {roster.length > 0 ? (
        <div className="space-y-1.5">
          <div className="text-[11px] font-medium text-muted-foreground">
            Clique no nome para escolher / desmarcar:
          </div>
          <div className="flex flex-wrap gap-1.5">
            {roster.map((n) => {
              const on = isSelected(n);
              return (
                <div
                  key={n}
                  className={`inline-flex items-center rounded-full border text-xs transition-colors overflow-hidden ${
                    on
                      ? "bg-primary text-primary-foreground border-primary"
                      : "bg-background text-foreground border-border hover:bg-accent"
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => toggle(n)}
                    className="flex items-center gap-1.5 px-3 py-1 text-xs font-medium focus:outline-none cursor-pointer"
                  >
                    {on && <span className="font-bold">✓</span>}
                    <span>{n}</span>
                  </button>
                  <button
                    type="button"
                    onClick={(e) => removeFromRoster(n, e)}
                    className={`px-2.5 py-1 text-xs font-bold transition-all border-l cursor-pointer ${
                      on
                        ? "text-primary-foreground/80 hover:text-white hover:bg-destructive border-primary-foreground/20"
                        : "text-muted-foreground hover:text-destructive hover:bg-destructive/10 border-border"
                    }`}
                    title={`Remover ${n} do histórico`}
                    aria-label={`Remover ${n} do histórico`}
                  >
                    ×
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      ) : (
        <span className="text-xs text-muted-foreground">
          Nenhum nome cadastrado. Digite acima para adicionar.
        </span>
      )}

      {selectedList.length > 0 && (
        <div className="rounded-md bg-secondary/50 p-2 text-[11px] text-foreground">
          <span className="font-semibold">
            Selecionados para a atividade ({selectedList.length}):
          </span>{" "}
          {selectedList.join("; ")}
        </div>
      )}
    </div>
  );
}

function LastDefectCard({ value: last, onSave }: { value: string; onSave: (val: string) => void }) {
  const [editing, setEditing] = useState(false);
  const [value, setValue] = useState(last);
  useEffect(() => {
    if (!editing) setValue(last);
  }, [last, editing]);
  return (
    <div className="rounded-xl border-2 border-primary/40 bg-primary/5 p-4 shadow-sm">
      <div className="flex items-center justify-between gap-3">
        <div className="min-w-0">
          <div className="text-[11px] font-semibold uppercase tracking-wider text-primary">
            Último defeito encontrado
          </div>
          {editing ? (
            <input
              autoFocus
              value={value}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  onSave(value.trim());
                  setEditing(false);
                }
                if (e.key === "Escape") {
                  setEditing(false);
                }
              }}
              className="input mt-1 text-2xl font-bold"
            />
          ) : (
            <div className="mt-1 truncate text-3xl font-bold text-foreground">
              {last || (
                <span className="text-muted-foreground text-xl font-normal">
                  — nenhum registrado —
                </span>
              )}
            </div>
          )}
        </div>
        {editing ? (
          <div className="flex gap-2 shrink-0">
            <button
              className="btn-primary"
              onClick={() => {
                onSave(value.trim());
                setEditing(false);
              }}
            >
              Salvar
            </button>
            <button className="btn-ghost" onClick={() => setEditing(false)}>
              Cancelar
            </button>
          </div>
        ) : (
          <button className="btn-secondary shrink-0" onClick={() => setEditing(true)}>
            Editar
          </button>
        )}
      </div>
    </div>
  );
}
