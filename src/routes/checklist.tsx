import { createFileRoute, Link } from "@tanstack/react-router";
import { useState, useEffect, useMemo, useRef } from "react";
import {
  type EquipmentChecklist,
  type ChecklistItem,
  type ChecklistItemStatus,
  type EquipmentOperatingStatus,
  type ChecklistPhoto,
  createEmptyChecklist,
  saveChecklistToCloud,
  deleteChecklistFromCloud,
  subscribeChecklists,
  getLocalChecklists,
  getChecklistTemplate,
  saveChecklistTemplate,
  calculateOperatingStatus,
  VRF_DOCUMENT_CODE,
  VRF_DOCUMENT_TITLE,
  VRF_DOCUMENT_SUBTITLE,
  DEFAULT_CHECKLIST_TEMPLATE,
  getChecklistDraft,
  saveChecklistDraft,
  clearChecklistDraft,
  isChecklistMeaningful,
} from "@/lib/checklist";
import { generateChecklistPdf } from "@/lib/checklist-pdf";
import { ChecklistDashboard } from "@/components/ChecklistDashboard";
import { ChecklistPhotoUploader } from "@/components/ChecklistPhotoUploader";
import { ThemeToggle } from "@/components/ThemeToggle";
import {
  ClipboardCheck,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  FileText,
  Save,
  Download,
  Plus,
  Trash2,
  Eye,
  Search,
  Cloud,
  RefreshCw,
  ArrowLeft,
  Settings,
  LayoutDashboard,
  Clock,
  Sparkles,
  HelpCircle,
  X,
  Share2,
  RotateCcw,
  MessageSquare,
  ListOrdered,
  Layers,
  Gauge,
  Check,
  Camera,
  ImagePlus,
  ShieldCheck,
  WifiOff,
} from "lucide-react";

export const Route = createFileRoute("/checklist")({
  component: EquipmentChecklistPage,
  head: () => ({
    meta: [
      { title: "Checklist de Equipamentos • TIC Trens" },
      {
        name: "description",
        content:
          "Inspeção diária de máquinas, veículos e equipamentos operacionais com sincronização em tempo real e dashboard.",
      },
    ],
  }),
});

export function EquipmentChecklistPage() {
  const [activeTab, setActiveTab] = useState<"form" | "dashboard" | "historico" | "personalizar">(
    "form",
  );

  // Recuperação de rascunho persistido no dispositivo (computador ou celular)
  const [initialDraft] = useState(() => getChecklistDraft());
  const [currentChecklist, setCurrentChecklist] = useState<EquipmentChecklist>(() => {
    const draft = getChecklistDraft();
    if (draft && draft.checklist) {
      return draft.checklist;
    }
    return createEmptyChecklist();
  });
  const [draftSavedAt, setDraftSavedAt] = useState<number | null>(() => {
    const draft = getChecklistDraft();
    return draft?.savedAt || null;
  });
  const [draftNoticeDismissed, setDraftNoticeDismissed] = useState<boolean>(false);
  const [lastAutoSaveText, setLastAutoSaveText] = useState<string | null>(() => {
    const draft = getChecklistDraft();
    if (draft?.savedAt) {
      const d = new Date(draft.savedAt);
      return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
    }
    return null;
  });
  const currentChecklistRef = useRef(currentChecklist);
  currentChecklistRef.current = currentChecklist;

  const [allChecklists, setAllChecklists] = useState<EquipmentChecklist[]>(() =>
    getLocalChecklists(),
  );
  const [isSaving, setIsSaving] = useState(false);
  const [saveSuccessMessage, setSaveSuccessMessage] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isOnline, setIsOnline] = useState(
    typeof navigator !== "undefined" ? navigator.onLine : true,
  );

  // History search and view modal
  const [historySearch, setHistorySearch] = useState("");
  const [viewingChecklist, setViewingChecklist] = useState<EquipmentChecklist | null>(null);

  // Template customizer state
  const [customItems, setCustomItems] = useState(() => getChecklistTemplate());
  const [newItemName, setNewItemName] = useState("");
  const [newItemCategory, setNewItemCategory] = useState("Geral");
  const [newItemRequired, setNewItemRequired] = useState(true);

  // Form view preferences & comment expansion
  const [itemsViewMode, setItemsViewMode] = useState<"sequencial" | "categoria">("sequencial");
  const [expandedComments, setExpandedComments] = useState<Record<string, boolean>>({});

  // Native camera & photo states
  const itemCameraInputRef = useRef<HTMLInputElement>(null);
  const [activePhotoItemId, setActivePhotoItemId] = useState<string | null>(null);
  const [expandedPhotoModalUrl, setExpandedPhotoModalUrl] = useState<string | null>(null);
  const [isExportingPdf, setIsExportingPdf] = useState(false);

  const toggleCommentInput = (itemId: string) => {
    setExpandedComments((prev) => ({
      ...prev,
      [itemId]: !prev[itemId],
    }));
  };

  const handleTriggerItemCamera = (itemId: string) => {
    setActivePhotoItemId(itemId);
    if (itemCameraInputRef.current) {
      itemCameraInputRef.current.click();
    }
  };

  const handleItemCameraCapture = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const reader = new FileReader();
      reader.onload = (ev) => {
        const raw = ev.target?.result as string;
        if (!raw) return;

        const matchedItem = currentChecklist.items.find((it) => it.id === activePhotoItemId);
        const newPhoto: ChecklistPhoto = {
          id: `photo-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
          dataUrl: raw,
          caption: matchedItem ? `Item: ${matchedItem.name}` : "",
          itemId: activePhotoItemId || undefined,
          itemName: matchedItem ? matchedItem.name : undefined,
          timestamp: Date.now(),
        };

        setCurrentChecklist((prev) => {
          const updatedFotos = [...(prev.fotos || []), newPhoto];
          return {
            ...prev,
            fotos: updatedFotos,
            fotosCount: updatedFotos.length,
          };
        });

        setSaveSuccessMessage(
          `Foto registrada com sucesso para o item "${matchedItem?.name || "do checklist"}"!`,
        );
        setTimeout(() => setSaveSuccessMessage(null), 4000);
      };
      reader.readAsDataURL(file);
    } catch (err) {
      console.warn("Erro ao registrar foto do item:", err);
    } finally {
      if (itemCameraInputRef.current) itemCameraInputRef.current.value = "";
    }
  };

  // Online / Offline listener
  useEffect(() => {
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, []);

  // Salvamento Imediato de Emergência: previne perda caso o usuário feche a aba, o app ou reinicie o celular
  useEffect(() => {
    const saveImmediately = () => {
      if (isChecklistMeaningful(currentChecklistRef.current)) {
        saveChecklistDraft(currentChecklistRef.current);
      }
    };

    const handleVisibility = () => {
      if (document.visibilityState === "hidden") {
        saveImmediately();
      }
    };

    window.addEventListener("beforeunload", saveImmediately);
    window.addEventListener("pagehide", saveImmediately);
    document.addEventListener("visibilitychange", handleVisibility);

    return () => {
      window.removeEventListener("beforeunload", saveImmediately);
      window.removeEventListener("pagehide", saveImmediately);
      document.removeEventListener("visibilitychange", handleVisibility);
    };
  }, []);

  // Autosave contínuo com debounce a cada alteração em qualquer campo ou foto
  useEffect(() => {
    const timer = setTimeout(() => {
      if (isChecklistMeaningful(currentChecklist)) {
        saveChecklistDraft(currentChecklist);
        const now = new Date();
        const formatted = `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}:${String(now.getSeconds()).padStart(2, "0")}`;
        setLastAutoSaveText(formatted);
        setDraftSavedAt(now.getTime());
      }
    }, 500);

    return () => clearTimeout(timer);
  }, [currentChecklist]);

  // Real-time Supabase / server subscription
  useEffect(() => {
    setIsSyncing(true);
    const unsub = subscribeChecklists((data) => {
      setAllChecklists(data);
      setIsSyncing(false);
    });
    return unsub;
  }, []);

  // Recalculate suggested operating status whenever items change
  const handleItemStatusChange = (itemId: string, status: ChecklistItemStatus) => {
    setCurrentChecklist((prev) => {
      const updatedItems = prev.items.map((item) =>
        item.id === itemId ? { ...item, status } : item,
      );
      const suggestedStatus = calculateOperatingStatus(updatedItems);
      return {
        ...prev,
        items: updatedItems,
        statusGeral: suggestedStatus,
      };
    });
  };

  const handleItemNotesChange = (itemId: string, notes: string) => {
    setCurrentChecklist((prev) => ({
      ...prev,
      items: prev.items.map((item) => (item.id === itemId ? { ...item, notes } : item)),
    }));
  };

  const handleItemActionChange = (itemId: string, acaoSolucao: string) => {
    setCurrentChecklist((prev) => ({
      ...prev,
      items: prev.items.map((item) => (item.id === itemId ? { ...item, acaoSolucao } : item)),
    }));
  };

  // Fast filling: Mark all pending items as Conforme
  const handleApproveAllConforme = () => {
    setCurrentChecklist((prev) => {
      const updatedItems = prev.items.map((item) => ({
        ...item,
        status: (item.status === "pendente" ? "conforme" : item.status) as ChecklistItemStatus,
      }));
      return {
        ...prev,
        items: updatedItems,
        statusGeral: calculateOperatingStatus(updatedItems),
      };
    });
    setSaveSuccessMessage("Todos os itens foram marcados como Conforme!");
    setTimeout(() => setSaveSuccessMessage(null), 3000);
  };

  // Reset checklist to a fresh form (descartar rascunho)
  const handleNewChecklist = () => {
    if (isChecklistMeaningful(currentChecklist)) {
      if (
        !confirm(
          "Deseja iniciar um novo checklist em branco? O rascunho salvo no celular/computador será descartado.",
        )
      ) {
        return;
      }
    }
    clearChecklistDraft();
    setDraftSavedAt(null);
    setLastAutoSaveText(null);
    setDraftNoticeDismissed(true);
    setCurrentChecklist(createEmptyChecklist());
    setSaveSuccessMessage("Novo checklist em branco pronto para preenchimento.");
    setTimeout(() => setSaveSuccessMessage(null), 3000);
    setErrorMessage(null);
  };

  // Save checklist (Cloud + Local)
  const handleSaveChecklist = async () => {
    if (!currentChecklist.equipamento.trim()) {
      setErrorMessage("Por favor, preencha o nome ou modelo do Equipamento.");
      return;
    }
    if (!currentChecklist.operador.trim()) {
      setErrorMessage("Por favor, preencha o nome do Operador ou Inspetor.");
      return;
    }

    setIsSaving(true);
    setErrorMessage(null);

    try {
      await saveChecklistToCloud(currentChecklist);
      // Rascunho finalizado com sucesso: limpa o rascunho temporário do celular/computador
      clearChecklistDraft();
      setDraftSavedAt(null);
      setLastAutoSaveText(null);
      setDraftNoticeDismissed(true);

      setSaveSuccessMessage("Checklist finalizado e salvo com sucesso na nuvem e no celular!");
      setTimeout(() => setSaveSuccessMessage(null), 5000);
    } catch (err) {
      console.error("Erro ao salvar checklist:", err);
      // Mesmo com falha de rede/cloud, garante que o rascunho e local checklist estejam seguros
      setSaveSuccessMessage(
        "Salvo com segurança no celular! Quando houver sinal de internet, os dados serão enviados à nuvem.",
      );
      setTimeout(() => setSaveSuccessMessage(null), 6000);
    } finally {
      setIsSaving(false);
    }
  };

  // PDF Export
  const handleExportPdf = async (checklistToExport?: EquipmentChecklist) => {
    const target = checklistToExport || currentChecklist;
    setIsExportingPdf(true);
    try {
      await generateChecklistPdf(target);
    } catch (err) {
      console.error("Erro ao gerar PDF:", err);
      alert("Erro ao exportar PDF. Tente novamente.");
    } finally {
      setIsExportingPdf(false);
    }
  };

  // Delete checklist from history
  const handleDeleteChecklist = async (id: string) => {
    if (!confirm("Tem certeza que deseja excluir permanentemente este checklist?")) return;
    try {
      await deleteChecklistFromCloud(id);
      setAllChecklists((prev) => prev.filter((c) => c.id !== id));
      if (viewingChecklist?.id === id) {
        setViewingChecklist(null);
      }
    } catch (err) {
      console.error("Erro ao excluir checklist:", err);
      alert("Erro ao excluir da nuvem. Tente novamente.");
    }
  };

  // Group items by category for ergonomic display
  const itemsByCategory = useMemo(() => {
    const groups: Record<string, ChecklistItem[]> = {};
    currentChecklist.items.forEach((item) => {
      const cat = item.category || "Geral";
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(item);
    });
    return groups;
  }, [currentChecklist.items]);

  // Filter history
  const filteredHistory = useMemo(() => {
    if (!historySearch.trim()) return allChecklists;
    const term = historySearch.toLowerCase();
    return allChecklists.filter(
      (c) =>
        (c.equipamento || "").toLowerCase().includes(term) ||
        (c.prefixo || "").toLowerCase().includes(term) ||
        (c.operador || "").toLowerCase().includes(term) ||
        (c.local || "").toLowerCase().includes(term) ||
        (c.data || "").includes(term),
    );
  }, [allChecklists, historySearch]);

  // Add custom check item
  const handleAddCustomItem = () => {
    if (!newItemName.trim()) return;
    const newItem = {
      id: `custom-${Date.now()}`,
      category: newItemCategory.trim() || "Geral",
      name: newItemName.trim(),
      required: newItemRequired,
    };
    const updated = [...customItems, newItem];
    setCustomItems(updated);
    saveChecklistTemplate(updated);
    setNewItemName("");
    // Also add to current active checklist if not present
    setCurrentChecklist((prev) => ({
      ...prev,
      items: [
        ...prev.items,
        {
          ...newItem,
          status: "pendente" as ChecklistItemStatus,
          notes: "",
        },
      ],
    }));
  };

  const handleRemoveCustomItem = (id: string) => {
    const updated = customItems.filter((item) => item.id !== id);
    setCustomItems(updated);
    saveChecklistTemplate(updated);
  };

  const handleResetToVrfTemplate = () => {
    if (
      confirm(
        "Deseja restaurar o modelo oficial VRF-SM PH-1016 com os 20 itens de inspeção pré-operação?",
      )
    ) {
      saveChecklistTemplate(DEFAULT_CHECKLIST_TEMPLATE);
      setCustomItems(DEFAULT_CHECKLIST_TEMPLATE);
      setCurrentChecklist((prev) => ({
        ...prev,
        items: DEFAULT_CHECKLIST_TEMPLATE.map((item) => ({
          ...item,
          status: "pendente" as ChecklistItemStatus,
          notes: "",
        })),
      }));
    }
  };

  return (
    <div className="min-h-screen bg-background text-foreground flex flex-col selection:bg-primary/20">
      {/* Top Header */}
      <header className="sticky top-0 z-40 bg-background/95 backdrop-blur-md border-b border-border shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3">
          {/* Brand & Title */}
          <div className="flex items-center gap-3">
            <div className="p-2 rounded-xl bg-primary/10 text-primary border border-primary/20">
              <ClipboardCheck className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="font-bold text-base sm:text-lg text-foreground tracking-tight">
                  Checklist de Equipamento
                </h1>
                <span className="text-[10px] font-semibold uppercase px-2 py-0.5 rounded-full bg-primary/10 text-primary border border-primary/20">
                  TIC Trens
                </span>
              </div>
              <p className="text-xs text-muted-foreground hidden sm:block">
                Inspeção diária de máquinas, veículos e implementos operacionais
              </p>
            </div>
          </div>

          {/* Quick Nav Actions */}
          <div className="flex items-center gap-2">
            {/* Sync status badge */}
            <div
              className={`hidden md:flex items-center gap-1.5 px-2.5 py-1 text-xs rounded-lg border font-medium ${
                !isOnline
                  ? "bg-amber-500/10 text-amber-600 border-amber-500/30"
                  : "bg-emerald-500/10 text-emerald-600 border-emerald-500/30"
              }`}
            >
              <Cloud className="w-3.5 h-3.5" />
              <span>{isOnline ? "Nuvem Sincronizada" : "Modo Offline"}</span>
            </div>

            {/* Link to Activities */}
            <Link
              to="/"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border transition-colors shadow-xs"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Atividades Diárias</span>
              <span className="sm:hidden">Parte Diária</span>
            </Link>

            {/* Link to Main Dashboard */}
            <Link
              to="/dashboard"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border transition-colors shadow-xs"
            >
              <LayoutDashboard className="w-3.5 h-3.5" />
              <span>Dashboard Geral</span>
            </Link>

            <ThemeToggle className="py-1.5 px-2.5" />
          </div>
        </div>

        {/* Tab Switcher */}
        <div className="max-w-7xl mx-auto px-4 sm:px-6">
          <div className="flex border-b border-border/80 gap-1 overflow-x-auto no-scrollbar">
            <button
              type="button"
              onClick={() => setActiveTab("form")}
              className={`py-2.5 px-4 text-xs font-semibold border-b-2 flex items-center gap-2 whitespace-nowrap transition-colors ${
                activeTab === "form"
                  ? "border-primary text-primary"
                  : "border-transparent text-muted-foreground hover:text-foreground"
              }`}
            >
              <ClipboardCheck className="w-4 h-4" />
              <span>Preencher Checklist</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveTab("dashboard")}
              className={`py-2.5 px-4 text-xs font-semibold border-b-2 flex items-center gap-2 whitespace-nowrap transition-colors ${
                activeTab === "dashboard"
                  ? "border-primary text-primary"
                  : "border-transparent text-muted-foreground hover:text-foreground"
              }`}
            >
              <LayoutDashboard className="w-4 h-4" />
              <span>Dashboard do Checklist</span>
              <span className="text-[10px] bg-primary/10 text-primary px-1.5 py-0.2 rounded-full font-mono">
                {allChecklists.length}
              </span>
            </button>

            <button
              type="button"
              onClick={() => setActiveTab("historico")}
              className={`py-2.5 px-4 text-xs font-semibold border-b-2 flex items-center gap-2 whitespace-nowrap transition-colors ${
                activeTab === "historico"
                  ? "border-primary text-primary"
                  : "border-transparent text-muted-foreground hover:text-foreground"
              }`}
            >
              <FileText className="w-4 h-4" />
              <span>Histórico & Registros</span>
              <span className="text-[10px] bg-muted text-muted-foreground px-1.5 py-0.2 rounded-full font-mono">
                {allChecklists.length}
              </span>
            </button>

            <button
              type="button"
              onClick={() => setActiveTab("personalizar")}
              className={`py-2.5 px-4 text-xs font-semibold border-b-2 flex items-center gap-2 whitespace-nowrap transition-colors ${
                activeTab === "personalizar"
                  ? "border-primary text-primary"
                  : "border-transparent text-muted-foreground hover:text-foreground"
              }`}
            >
              <Settings className="w-4 h-4" />
              <span>Personalizar Itens</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 space-y-6">
        {/* TAB 1: FORMULÁRIO DE CHECKLIST */}
        {activeTab === "form" && (
          <div className="space-y-6">
            {/* Input nativo invisível para câmera em item individual */}
            <input
              type="file"
              ref={itemCameraInputRef}
              onChange={handleItemCameraCapture}
              accept="image/*"
              capture="environment"
              className="hidden"
              id="item-camera-input-native"
            />
            {/* Status alerts */}
            {saveSuccessMessage && (
              <div className="p-3.5 rounded-xl bg-emerald-500/10 text-emerald-700 dark:text-emerald-300 border border-emerald-500/20 text-xs flex items-center gap-2 animate-in fade-in">
                <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-600 dark:text-emerald-400" />
                <span className="font-medium">{saveSuccessMessage}</span>
              </div>
            )}

            {errorMessage && (
              <div className="p-3.5 rounded-xl bg-destructive/10 text-destructive border border-destructive/20 text-xs flex items-center gap-2 animate-in fade-in">
                <AlertTriangle className="w-4 h-4 shrink-0" />
                <span className="font-medium">{errorMessage}</span>
              </div>
            )}

            {/* Aviso de Recuperação de Rascunho após Fechamento/Reinício */}
            {initialDraft && !draftNoticeDismissed && (
              <div
                id="draft-restored-banner"
                className="p-4 rounded-xl bg-amber-500/10 dark:bg-amber-500/15 border border-amber-500/30 text-foreground text-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3 animate-in fade-in shadow-xs"
              >
                <div className="flex items-start sm:items-center gap-3">
                  <div className="p-2 rounded-lg bg-amber-500/20 text-amber-700 dark:text-amber-300 shrink-0">
                    <RotateCcw className="w-4 h-4" />
                  </div>
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="font-semibold text-sm text-foreground">
                        Rascunho recuperado com sucesso!
                      </span>
                      <span className="px-2 py-0.5 rounded-full text-[10px] font-medium bg-amber-500/20 text-amber-700 dark:text-amber-300">
                        Memória do aparelho
                      </span>
                    </div>
                    <p className="text-xs text-muted-foreground mt-0.5">
                      Você estava preenchendo o checklist de{" "}
                      <strong className="text-foreground">
                        {initialDraft.checklist.equipamento ||
                          initialDraft.checklist.prefixo ||
                          "Equipamento em andamento"}
                      </strong>
                      . Os dados foram preservados mesmo com o celular reiniciado, fechado ou sem
                      sinal de internet.
                      {initialDraft.savedAt && (
                        <span>
                          {" "}
                          Última gravação às{" "}
                          {new Date(initialDraft.savedAt).toLocaleTimeString([], {
                            hour: "2-digit",
                            minute: "2-digit",
                          })}
                          .
                        </span>
                      )}
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2 self-end sm:self-auto shrink-0">
                  <button
                    type="button"
                    onClick={() => setDraftNoticeDismissed(true)}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-xs shadow-xs transition-colors"
                  >
                    <Check className="w-3.5 h-3.5" />
                    <span>Continuar Preenchimento</span>
                  </button>
                  <button
                    type="button"
                    onClick={handleNewChecklist}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-background hover:bg-muted text-muted-foreground hover:text-destructive border border-border text-xs transition-colors"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Descartar</span>
                  </button>
                </div>
              </div>
            )}

            {/* Barra de Status de Persistência Offline & Proteção contra Reinício */}
            <div className="flex flex-wrap items-center justify-between gap-2.5 px-3.5 py-2.5 rounded-xl bg-card border border-border text-xs">
              <div className="flex items-center gap-2">
                <span className="relative flex h-2.5 w-2.5">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500"></span>
                </span>
                <span className="font-semibold text-foreground flex items-center gap-1.5">
                  <ShieldCheck className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                  Proteção Antiperda Ativa:
                </span>
                <span className="text-muted-foreground">
                  {lastAutoSaveText ? (
                    <>
                      Rascunho salvo no celular/computador às{" "}
                      <strong className="text-foreground font-mono">{lastAutoSaveText}</strong>
                    </>
                  ) : (
                    "Rascunho será salvo automaticamente no dispositivo ao digitar"
                  )}
                </span>
              </div>

              <div className="flex items-center gap-3">
                <span className="inline-flex items-center gap-1.5">
                  {isOnline ? (
                    <span className="text-emerald-600 dark:text-emerald-400 flex items-center gap-1 font-medium">
                      <Cloud className="w-3.5 h-3.5" />
                      Online (Sincronização nuvem ativa)
                    </span>
                  ) : (
                    <span className="text-amber-600 dark:text-amber-400 flex items-center gap-1 font-medium">
                      <WifiOff className="w-3.5 h-3.5" />
                      Offline (Preenchimento seguro no celular)
                    </span>
                  )}
                </span>
                {isChecklistMeaningful(currentChecklist) && (
                  <button
                    type="button"
                    onClick={handleNewChecklist}
                    className="text-xs text-muted-foreground hover:text-destructive underline decoration-dotted transition-colors"
                    title="Descartar rascunho atual e começar um checklist em branco"
                  >
                    Descartar rascunho
                  </button>
                )}
              </div>
            </div>

            {/* Header info card */}
            <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-border/60">
                <div>
                  <h3 className="font-semibold text-foreground text-sm sm:text-base">
                    Identificação do Equipamento & Operação
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Preencha os dados da máquina e do responsável antes de iniciar a checagem
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    id="header-tirar-foto-btn"
                    onClick={() => {
                      const el = document.getElementById("camera-input-native") as HTMLInputElement;
                      if (el) {
                        el.click();
                      } else {
                        const uploaderEl = document.getElementById("checklist-photo-uploader");
                        uploaderEl?.scrollIntoView({ behavior: "smooth" });
                      }
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white shadow-xs transition-colors"
                    title="Tirar foto com a câmera nativa do celular agora"
                  >
                    <Camera className="w-3.5 h-3.5" />
                    <span>
                      Tirar Foto{" "}
                      {(currentChecklist.fotos?.length ?? 0) > 0
                        ? `(${currentChecklist.fotos?.length})`
                        : ""}
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={handleApproveAllConforme}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-primary/10 hover:bg-primary/20 text-primary border border-primary/20 shadow-xs transition-colors"
                    title="Marca todos os itens pendentes como Conforme de uma só vez"
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    <span>Aprovar Conformes</span>
                  </button>

                  <button
                    type="button"
                    onClick={handleNewChecklist}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border shadow-xs transition-colors"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Novo em Branco</span>
                  </button>
                </div>
              </div>

              {/* Form Fields Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5">
                {/* Equipamento */}
                <div className="space-y-1">
                  <label className="text-[11px] font-semibold text-foreground">
                    Equipamento / Máquina <span className="text-destructive">*</span>
                  </label>
                  <input
                    type="text"
                    list="equipamentos-sugestoes"
                    value={currentChecklist.equipamento}
                    onChange={(e) =>
                      setCurrentChecklist((prev) => ({ ...prev, equipamento: e.target.value }))
                    }
                    placeholder="Ex: Retroescavadeira, Socadora..."
                    className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                  />
                  <datalist id="equipamentos-sugestoes">
                    <option value="VRF Carro Ultrassom (PH-1016)" />
                    <option value="VRF - Veículo Rodoferroviário" />
                    <option value="Caminhão Rodoferroviário" />
                    <option value="Auto de Linha" />
                    <option value="Carro Ultrassom de Inspeção" />
                    <option value="Socadora de Lastro" />
                    <option value="Reguladora de Lastro" />
                    <option value="Guindaste Ferroviário" />
                    <option value="Retroescavadeira CAT 416" />
                    <option value="Plataforma Elevatória" />
                  </datalist>
                </div>

                {/* Prefixo / Placa */}
                <div className="space-y-1">
                  <label className="text-[11px] font-semibold text-foreground">
                    Prefixo / Placa / ID
                  </label>
                  <input
                    type="text"
                    value={currentChecklist.prefixo}
                    onChange={(e) =>
                      setCurrentChecklist((prev) => ({ ...prev, prefixo: e.target.value }))
                    }
                    placeholder="Ex: EQ-104, R-02, ABC-1234"
                    className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                  />
                </div>

                {/* Operador / Inspetor */}
                <div className="space-y-1">
                  <label className="text-[11px] font-semibold text-foreground">
                    Operador / Inspetor <span className="text-destructive">*</span>
                  </label>
                  <input
                    type="text"
                    value={currentChecklist.operador}
                    onChange={(e) =>
                      setCurrentChecklist((prev) => ({ ...prev, operador: e.target.value }))
                    }
                    placeholder="Nome completo do responsável"
                    className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                  />
                </div>

                {/* Turno */}
                <div className="space-y-1">
                  <label className="text-[11px] font-semibold text-foreground">
                    Turno de Trabalho
                  </label>
                  <select
                    value={currentChecklist.turno}
                    onChange={(e) =>
                      setCurrentChecklist((prev) => ({ ...prev, turno: e.target.value }))
                    }
                    className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                  >
                    <option value="Diurno">Diurno</option>
                    <option value="Noturno">Noturno</option>
                    <option value="Revezamento">Revezamento</option>
                  </select>
                </div>

                {/* Data */}
                <div className="space-y-1">
                  <label className="text-[11px] font-semibold text-foreground">
                    Data da Inspeção
                  </label>
                  <input
                    type="date"
                    value={currentChecklist.data}
                    onChange={(e) =>
                      setCurrentChecklist((prev) => ({ ...prev, data: e.target.value }))
                    }
                    className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                  />
                </div>

                {/* Hora */}
                <div className="space-y-1">
                  <label className="text-[11px] font-semibold text-foreground">Horário</label>
                  <input
                    type="time"
                    value={currentChecklist.hora}
                    onChange={(e) =>
                      setCurrentChecklist((prev) => ({ ...prev, hora: e.target.value }))
                    }
                    className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                  />
                </div>

                {/* Local / Frente */}
                <div className="space-y-1 sm:col-span-2">
                  <label className="text-[11px] font-semibold text-foreground">
                    Local / Frente de Trabalho
                  </label>
                  <input
                    type="text"
                    value={currentChecklist.local}
                    onChange={(e) =>
                      setCurrentChecklist((prev) => ({ ...prev, local: e.target.value }))
                    }
                    placeholder="Ex: KM 18 - Linha 1, Pátio Lapa"
                    className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                  />
                </div>
              </div>

              {/* Horímetros e Quilometragem (Subseção Dedicada) */}
              <div className="pt-3 border-t border-border/60 space-y-2.5">
                <div className="flex items-center gap-1.5">
                  <Gauge className="w-4 h-4 text-primary" />
                  <span className="text-xs font-bold text-foreground uppercase tracking-wide">
                    Horímetros & Quilometragem
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5">
                  {/* Horímetro do Motor */}
                  <div className="space-y-1">
                    <label className="text-[11px] font-semibold text-foreground">
                      Horímetro do Motor
                    </label>
                    <input
                      type="text"
                      value={currentChecklist.horimetroMotor}
                      onChange={(e) =>
                        setCurrentChecklist((prev) => ({
                          ...prev,
                          horimetroMotor: e.target.value,
                        }))
                      }
                      placeholder="Ex: 1.250 h"
                      className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background font-mono"
                    />
                  </div>

                  {/* Horímetro do Implemento Ferroviário */}
                  <div className="space-y-1">
                    <label className="text-[11px] font-semibold text-foreground">
                      Horímetro do Implemento Ferroviário
                    </label>
                    <input
                      type="text"
                      value={currentChecklist.horimetroImplementoFerroviario}
                      onChange={(e) =>
                        setCurrentChecklist((prev) => ({
                          ...prev,
                          horimetroImplementoFerroviario: e.target.value,
                        }))
                      }
                      placeholder="Ex: 420 h"
                      className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background font-mono"
                    />
                  </div>

                  {/* Horímetro Gerador */}
                  <div className="space-y-1">
                    <label className="text-[11px] font-semibold text-foreground">
                      Horímetro Gerador
                    </label>
                    <input
                      type="text"
                      value={currentChecklist.horimetroGerador}
                      onChange={(e) =>
                        setCurrentChecklist((prev) => ({
                          ...prev,
                          horimetroGerador: e.target.value,
                        }))
                      }
                      placeholder="Ex: 850 h"
                      className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background font-mono"
                    />
                  </div>

                  {/* Quilometragem do Caminhão */}
                  <div className="space-y-1">
                    <label className="text-[11px] font-semibold text-foreground">
                      Quilometragem do Caminhão
                    </label>
                    <input
                      type="text"
                      value={currentChecklist.quilometragemCaminhao}
                      onChange={(e) =>
                        setCurrentChecklist((prev) => ({
                          ...prev,
                          quilometragemCaminhao: e.target.value,
                        }))
                      }
                      placeholder="Ex: 48.350 km"
                      className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background font-mono"
                    />
                  </div>
                </div>
              </div>
            </div>

            {/* Document Header Banner (VRF-SM PH-1016) */}
            <div className="bg-primary/5 border border-primary/20 rounded-xl p-4 sm:p-5 shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="space-y-1">
                <div className="flex flex-wrap items-center gap-2">
                  <span className="px-2.5 py-0.5 rounded-md bg-primary text-primary-foreground font-mono font-bold text-xs">
                    {VRF_DOCUMENT_CODE}
                  </span>
                  <h3 className="font-bold text-sm sm:text-base text-foreground tracking-tight">
                    {VRF_DOCUMENT_TITLE}
                  </h3>
                </div>
                <p className="text-xs text-muted-foreground max-w-3xl leading-relaxed">
                  {VRF_DOCUMENT_SUBTITLE}
                </p>
              </div>

              {/* View mode toggle */}
              <div className="flex items-center gap-1.5 self-start md:self-auto bg-muted/40 p-1 rounded-lg border border-border/70 shrink-0">
                <button
                  type="button"
                  onClick={() => setItemsViewMode("sequencial")}
                  className={`px-3 py-1.5 text-xs font-semibold rounded-md flex items-center gap-1.5 transition-all ${
                    itemsViewMode === "sequencial"
                      ? "bg-primary text-primary-foreground shadow-xs"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                  title="Ordem de 1 a 20 como na folha impressa"
                >
                  <ListOrdered className="w-3.5 h-3.5" />
                  <span>Sequencial (1-20)</span>
                </button>
                <button
                  type="button"
                  onClick={() => setItemsViewMode("categoria")}
                  className={`px-3 py-1.5 text-xs font-semibold rounded-md flex items-center gap-1.5 transition-all ${
                    itemsViewMode === "categoria"
                      ? "bg-primary text-primary-foreground shadow-xs"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                  title="Agrupar por categorias técnicas"
                >
                  <Layers className="w-3.5 h-3.5" />
                  <span>Por Categoria</span>
                </button>
              </div>
            </div>

            {/* Checklist items rendering */}
            <div className="space-y-5">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="font-semibold text-foreground text-sm sm:text-base flex items-center gap-2">
                    <span>Componentes Diários para Checagem</span>
                    <span className="text-xs font-normal text-muted-foreground">
                      (Conforme folha {VRF_DOCUMENT_CODE})
                    </span>
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Avalie como <strong>OK</strong> (Conforme), <strong>NOK</strong> (Não Conforme)
                    ou <strong>N/A</strong> e adicione comentários quando necessário
                  </p>
                </div>
                <div className="text-xs text-muted-foreground font-medium">
                  <span className="font-bold text-foreground">
                    {currentChecklist.items.filter((i) => i.status !== "pendente").length}
                  </span>{" "}
                  de {currentChecklist.items.length} checados
                </div>
              </div>

              {/* SEQUENTIAL VIEW (1-20) */}
              {itemsViewMode === "sequencial" && (
                <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs divide-y divide-border/60">
                  {currentChecklist.items.map((item, index) => {
                    const isCommentOpen =
                      item.status === "nao_conforme" ||
                      Boolean(item.notes) ||
                      Boolean(expandedComments[item.id]);

                    return (
                      <div
                        key={item.id}
                        className={`py-3.5 first:pt-0 last:pb-0 space-y-2.5 transition-colors ${
                          item.status === "nao_conforme"
                            ? "bg-destructive/5 -mx-3 px-3 rounded-lg"
                            : ""
                        }`}
                      >
                        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                          {/* Item Info */}
                          <div className="space-y-1 max-w-2xl">
                            <div className="flex items-start gap-2.5">
                              <span className="font-mono text-xs font-bold text-primary px-1.5 py-0.5 rounded bg-primary/10 border border-primary/20 shrink-0 mt-0.5">
                                {String(index + 1).padStart(2, "0")}
                              </span>
                              <div>
                                <div className="flex flex-wrap items-center gap-2">
                                  <span className="text-xs font-semibold text-foreground">
                                    {item.name}
                                  </span>
                                  <span className="text-[10px] text-muted-foreground px-1.5 py-0.5 rounded bg-muted/60">
                                    {item.category}
                                  </span>
                                  {item.required && (
                                    <span
                                      className="text-[10px] font-bold text-destructive"
                                      title="Item de checagem prioritária"
                                    >
                                      *
                                    </span>
                                  )}
                                </div>
                                {item.description && (
                                  <p className="text-[11px] text-muted-foreground mt-0.5">
                                    {item.description}
                                  </p>
                                )}
                              </div>
                            </div>
                          </div>

                          {/* Status buttons & comment trigger */}
                          <div className="flex flex-wrap items-center gap-2 shrink-0">
                            <div className="flex items-center gap-1.5 bg-muted/40 p-1 rounded-lg border border-border/70">
                              {/* OK (Conforme) */}
                              <button
                                type="button"
                                onClick={() => handleItemStatusChange(item.id, "conforme")}
                                className={`px-3 py-1.5 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all ${
                                  item.status === "conforme"
                                    ? "bg-emerald-600 text-white shadow-xs"
                                    : "text-muted-foreground hover:text-foreground hover:bg-muted/60"
                                }`}
                                title="Conforme / Em perfeitas condições"
                              >
                                <CheckCircle2 className="w-3.5 h-3.5" />
                                <span>OK</span>
                              </button>

                              {/* NOK (Não Conforme) */}
                              <button
                                type="button"
                                onClick={() => {
                                  handleItemStatusChange(item.id, "nao_conforme");
                                  if (!expandedComments[item.id]) {
                                    toggleCommentInput(item.id);
                                  }
                                }}
                                className={`px-3 py-1.5 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all ${
                                  item.status === "nao_conforme"
                                    ? "bg-destructive text-white shadow-xs"
                                    : "text-muted-foreground hover:text-foreground hover:bg-muted/60"
                                }`}
                                title="Não Conforme / Apresenta anomalia ou defeito"
                              >
                                <XCircle className="w-3.5 h-3.5" />
                                <span>NOK</span>
                              </button>

                              {/* N/A (Não se Aplica) */}
                              <button
                                type="button"
                                onClick={() => handleItemStatusChange(item.id, "nao_se_aplica")}
                                className={`px-2.5 py-1.5 text-xs font-medium rounded-md flex items-center gap-1 transition-all ${
                                  item.status === "nao_se_aplica"
                                    ? "bg-muted text-foreground font-bold shadow-xs"
                                    : "text-muted-foreground hover:text-foreground hover:bg-muted/60"
                                }`}
                                title="Não se aplica a esta configuração de equipamento"
                              >
                                <span>N/A</span>
                              </button>
                            </div>

                            {/* Comment Button Toggle */}
                            <button
                              type="button"
                              onClick={() => toggleCommentInput(item.id)}
                              className={`p-1.5 text-xs rounded-lg border transition-colors flex items-center gap-1 ${
                                item.notes
                                  ? "bg-amber-500/10 border-amber-500/30 text-amber-600 dark:text-amber-400 font-semibold"
                                  : isCommentOpen
                                    ? "bg-muted border-border text-foreground"
                                    : "bg-transparent border-transparent text-muted-foreground hover:bg-muted hover:text-foreground"
                              }`}
                              title={
                                item.notes ? `Comentário: ${item.notes}` : "Adicionar comentário"
                              }
                            >
                              <MessageSquare className="w-3.5 h-3.5" />
                              {item.notes && (
                                <span className="text-[10px] hidden sm:inline">Comentário</span>
                              )}
                            </button>

                            {/* Botão de Câmera Nativa no Item */}
                            {(() => {
                              const itemPhotos = (currentChecklist.fotos || []).filter(
                                (p) => p.itemId === item.id,
                              );
                              return (
                                <button
                                  type="button"
                                  onClick={() => handleTriggerItemCamera(item.id)}
                                  className={`p-1.5 text-xs rounded-lg border transition-colors flex items-center gap-1 ${
                                    itemPhotos.length > 0
                                      ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-600 dark:text-emerald-400 font-semibold"
                                      : "bg-transparent border-transparent text-muted-foreground hover:bg-muted hover:text-foreground"
                                  }`}
                                  title={
                                    itemPhotos.length > 0
                                      ? `${itemPhotos.length} foto(s) anexada(s) a este item. Clique para tirar outra foto.`
                                      : "Tirar foto deste componente com a câmera do celular"
                                  }
                                >
                                  <Camera className="w-3.5 h-3.5" />
                                  {itemPhotos.length > 0 && (
                                    <span className="text-[10px] font-bold">
                                      {itemPhotos.length}
                                    </span>
                                  )}
                                </button>
                              );
                            })()}
                          </div>
                        </div>

                        {/* Miniaturas de fotos anexadas ao item */}
                        {(() => {
                          const itemPhotos = (currentChecklist.fotos || []).filter(
                            (p) => p.itemId === item.id,
                          );
                          if (itemPhotos.length === 0) return null;
                          return (
                            <div className="pl-0 sm:pl-8 flex flex-wrap items-center gap-2 pt-1">
                              <span className="text-[10px] text-muted-foreground font-semibold flex items-center gap-1">
                                <Camera className="w-3 h-3 text-emerald-600" />
                                Evidências:
                              </span>
                              {itemPhotos.map((ph, pi) => (
                                <div
                                  key={ph.id || pi}
                                  className="relative group w-10 h-10 rounded-md overflow-hidden border border-border/80 shadow-2xs cursor-pointer"
                                  onClick={() => setExpandedPhotoModalUrl(ph.dataUrl)}
                                  title={ph.caption || "Clique para ampliar foto"}
                                >
                                  <img
                                    src={ph.dataUrl}
                                    alt={ph.caption || "Evidência"}
                                    className="w-full h-full object-cover group-hover:scale-110 transition-transform"
                                  />
                                </div>
                              ))}
                            </div>
                          );
                        })()}

                        {/* Comment input field & Action for nao_conforme */}
                        {isCommentOpen && (
                          <div className="pl-0 sm:pl-8 animate-in fade-in space-y-1.5 pt-1">
                            <div className="flex items-center gap-2">
                              <span className="text-[11px] font-semibold text-muted-foreground whitespace-nowrap">
                                Comentários:
                              </span>
                              <input
                                type="text"
                                value={item.notes || ""}
                                onChange={(e) => handleItemNotesChange(item.id, e.target.value)}
                                placeholder="Ex: Nível baixo, lâmpada queimada, vazamento..."
                                className={`w-full text-xs px-2.5 py-1.5 bg-background border rounded-lg focus:outline-none focus:ring-1 text-foreground ${
                                  item.status === "nao_conforme"
                                    ? "border-destructive/50 focus:ring-destructive focus:border-destructive"
                                    : "border-input focus:ring-primary focus:border-primary"
                                }`}
                                autoFocus={item.status === "nao_conforme" && !item.notes}
                              />
                            </div>

                            {item.status === "nao_conforme" && (
                              <div className="flex items-center gap-2">
                                <span className="text-[11px] font-semibold text-amber-600 dark:text-amber-400 whitespace-nowrap">
                                  Ação p/ Solucionar:
                                </span>
                                <input
                                  type="text"
                                  value={item.acaoSolucao || ""}
                                  onChange={(e) => handleItemActionChange(item.id, e.target.value)}
                                  placeholder="Ação adotada para solucionar o problema (ex: Completar nível com ARLA 32, Troca de peça...)"
                                  className="w-full text-xs px-2.5 py-1.5 bg-background border border-amber-500/40 rounded-lg focus:outline-none focus:ring-1 focus:ring-amber-500 text-foreground"
                                />
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}

              {/* CATEGORIZED VIEW */}
              {itemsViewMode === "categoria" && (
                <div className="space-y-4">
                  {Object.entries(itemsByCategory).map(([category, items]) => (
                    <div
                      key={category}
                      className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-3"
                    >
                      <div className="flex items-center justify-between pb-2 border-b border-border/60">
                        <h4 className="font-semibold text-xs sm:text-sm text-foreground uppercase tracking-wide flex items-center gap-2">
                          <span className="w-2 h-2 rounded-full bg-primary" />
                          {category}
                        </h4>
                        <span className="text-[11px] text-muted-foreground">
                          {items.filter((i) => i.status === "conforme").length}/{items.length}{" "}
                          conformes (OK)
                        </span>
                      </div>

                      <div className="divide-y divide-border/60">
                        {items.map((item) => {
                          const globalIndex = currentChecklist.items.findIndex(
                            (i) => i.id === item.id,
                          );
                          const isCommentOpen =
                            item.status === "nao_conforme" ||
                            Boolean(item.notes) ||
                            Boolean(expandedComments[item.id]);

                          return (
                            <div
                              key={item.id}
                              className={`py-3 first:pt-1 last:pb-1 space-y-2 transition-colors ${
                                item.status === "nao_conforme"
                                  ? "bg-destructive/5 -mx-2 px-2 rounded-lg"
                                  : ""
                              }`}
                            >
                              <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                                <div className="space-y-1 max-w-xl">
                                  <div className="flex items-start gap-2">
                                    <span className="font-mono text-[10px] font-bold text-primary px-1 py-0.5 rounded bg-primary/10 shrink-0 mt-0.5">
                                      #{String(globalIndex + 1).padStart(2, "0")}
                                    </span>
                                    <div>
                                      <span className="text-xs font-semibold text-foreground">
                                        {item.name}
                                      </span>
                                      {item.description && (
                                        <p className="text-[11px] text-muted-foreground mt-0.5">
                                          {item.description}
                                        </p>
                                      )}
                                    </div>
                                  </div>
                                </div>

                                <div className="flex flex-wrap items-center gap-2 shrink-0">
                                  <div className="flex items-center gap-1.5 bg-muted/40 p-1 rounded-lg border border-border/70">
                                    <button
                                      type="button"
                                      onClick={() => handleItemStatusChange(item.id, "conforme")}
                                      className={`px-3 py-1.5 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all ${
                                        item.status === "conforme"
                                          ? "bg-emerald-600 text-white shadow-xs"
                                          : "text-muted-foreground hover:text-foreground hover:bg-muted/60"
                                      }`}
                                    >
                                      <CheckCircle2 className="w-3.5 h-3.5" />
                                      <span>OK</span>
                                    </button>

                                    <button
                                      type="button"
                                      onClick={() => {
                                        handleItemStatusChange(item.id, "nao_conforme");
                                        if (!expandedComments[item.id]) {
                                          toggleCommentInput(item.id);
                                        }
                                      }}
                                      className={`px-3 py-1.5 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all ${
                                        item.status === "nao_conforme"
                                          ? "bg-destructive text-white shadow-xs"
                                          : "text-muted-foreground hover:text-foreground hover:bg-muted/60"
                                      }`}
                                    >
                                      <XCircle className="w-3.5 h-3.5" />
                                      <span>NOK</span>
                                    </button>

                                    <button
                                      type="button"
                                      onClick={() =>
                                        handleItemStatusChange(item.id, "nao_se_aplica")
                                      }
                                      className={`px-2.5 py-1.5 text-xs font-medium rounded-md flex items-center gap-1 transition-all ${
                                        item.status === "nao_se_aplica"
                                          ? "bg-muted text-foreground font-semibold shadow-xs"
                                          : "text-muted-foreground hover:text-foreground hover:bg-muted/60"
                                      }`}
                                    >
                                      <span>N/A</span>
                                    </button>
                                  </div>

                                  <button
                                    type="button"
                                    onClick={() => toggleCommentInput(item.id)}
                                    className={`p-1.5 text-xs rounded-lg border transition-colors flex items-center gap-1 ${
                                      item.notes
                                        ? "bg-amber-500/10 border-amber-500/30 text-amber-600 font-semibold"
                                        : isCommentOpen
                                          ? "bg-muted border-border text-foreground"
                                          : "bg-transparent border-transparent text-muted-foreground hover:bg-muted"
                                    }`}
                                    title="Comentários"
                                  >
                                    <MessageSquare className="w-3.5 h-3.5" />
                                  </button>

                                  {/* Botão de Câmera Nativa no Item (Categoria) */}
                                  {(() => {
                                    const itemPhotos = (currentChecklist.fotos || []).filter(
                                      (p) => p.itemId === item.id,
                                    );
                                    return (
                                      <button
                                        type="button"
                                        onClick={() => handleTriggerItemCamera(item.id)}
                                        className={`p-1.5 text-xs rounded-lg border transition-colors flex items-center gap-1 ${
                                          itemPhotos.length > 0
                                            ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-600 dark:text-emerald-400 font-semibold"
                                            : "bg-transparent border-transparent text-muted-foreground hover:bg-muted hover:text-foreground"
                                        }`}
                                        title={
                                          itemPhotos.length > 0
                                            ? `${itemPhotos.length} foto(s) anexada(s) a este item. Clique para tirar outra foto.`
                                            : "Tirar foto deste componente com a câmera do celular"
                                        }
                                      >
                                        <Camera className="w-3.5 h-3.5" />
                                        {itemPhotos.length > 0 && (
                                          <span className="text-[10px] font-bold">
                                            {itemPhotos.length}
                                          </span>
                                        )}
                                      </button>
                                    );
                                  })()}
                                </div>
                              </div>

                              {/* Miniaturas de fotos anexadas ao item */}
                              {(() => {
                                const itemPhotos = (currentChecklist.fotos || []).filter(
                                  (p) => p.itemId === item.id,
                                );
                                if (itemPhotos.length === 0) return null;
                                return (
                                  <div className="flex flex-wrap items-center gap-2 pt-1">
                                    <span className="text-[10px] text-muted-foreground font-semibold flex items-center gap-1">
                                      <Camera className="w-3 h-3 text-emerald-600" />
                                      Evidências:
                                    </span>
                                    {itemPhotos.map((ph, pi) => (
                                      <div
                                        key={ph.id || pi}
                                        className="relative group w-10 h-10 rounded-md overflow-hidden border border-border/80 shadow-2xs cursor-pointer"
                                        onClick={() => setExpandedPhotoModalUrl(ph.dataUrl)}
                                        title={ph.caption || "Clique para ampliar foto"}
                                      >
                                        <img
                                          src={ph.dataUrl}
                                          alt={ph.caption || "Evidência"}
                                          className="w-full h-full object-cover group-hover:scale-110 transition-transform"
                                        />
                                      </div>
                                    ))}
                                  </div>
                                );
                              })()}

                              {isCommentOpen && (
                                <div className="animate-in fade-in pt-1 space-y-1.5">
                                  <input
                                    type="text"
                                    value={item.notes || ""}
                                    onChange={(e) => handleItemNotesChange(item.id, e.target.value)}
                                    placeholder="Comentário sobre a anomalia ou defeito..."
                                    className="w-full text-xs px-2.5 py-1.5 bg-background border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary text-foreground"
                                  />
                                  {item.status === "nao_conforme" && (
                                    <div className="flex items-center gap-2">
                                      <span className="text-[10px] font-semibold text-amber-600 dark:text-amber-400 whitespace-nowrap">
                                        Ação p/ Solucionar:
                                      </span>
                                      <input
                                        type="text"
                                        value={item.acaoSolucao || ""}
                                        onChange={(e) =>
                                          handleItemActionChange(item.id, e.target.value)
                                        }
                                        placeholder="Ação adotada para solucionar o problema..."
                                        className="w-full text-xs px-2.5 py-1.5 bg-background border border-amber-500/40 rounded-lg focus:outline-none focus:ring-1 focus:ring-amber-500 text-foreground"
                                      />
                                    </div>
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Seção Completa de Fotos e Câmera Nativa */}
            <ChecklistPhotoUploader
              photos={currentChecklist.fotos || []}
              onChange={(newPhotos) =>
                setCurrentChecklist((prev) => ({
                  ...prev,
                  fotos: newPhotos,
                  fotosCount: newPhotos.length,
                }))
              }
              availableItems={currentChecklist.items}
              maxPhotos={15}
            />

            {/* Operating Status & General Observations Card */}
            <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-4">
              <div className="flex items-center justify-between pb-3 border-b border-border/60">
                <div>
                  <h3 className="font-semibold text-foreground text-sm sm:text-base">
                    Parecer de Liberação & Observações Finais
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Classificação final de prontidão para a operação do equipamento
                  </p>
                </div>
              </div>

              {/* Radio selector for final operating status */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                {/* Apto */}
                <div
                  onClick={() => setCurrentChecklist((prev) => ({ ...prev, statusGeral: "apto" }))}
                  className={`p-3.5 rounded-xl border cursor-pointer transition-all flex items-start gap-3 ${
                    currentChecklist.statusGeral === "apto"
                      ? "bg-emerald-500/10 border-emerald-500 text-foreground ring-1 ring-emerald-500"
                      : "bg-muted/20 border-border hover:bg-muted/40 text-muted-foreground"
                  }`}
                >
                  <div className="p-1.5 rounded-full bg-emerald-500/20 text-emerald-600 dark:text-emerald-400 shrink-0 mt-0.5">
                    <CheckCircle2 className="w-4 h-4" />
                  </div>
                  <div>
                    <div className="font-semibold text-xs text-foreground">Apto para Operação</div>
                    <div className="text-[11px] text-muted-foreground mt-0.5">
                      Equipamento seguro e 100% liberado para trabalho em campo.
                    </div>
                  </div>
                </div>

                {/* Atencao / Restricao */}
                <div
                  onClick={() =>
                    setCurrentChecklist((prev) => ({ ...prev, statusGeral: "atencao" }))
                  }
                  className={`p-3.5 rounded-xl border cursor-pointer transition-all flex items-start gap-3 ${
                    currentChecklist.statusGeral === "atencao"
                      ? "bg-amber-500/10 border-amber-500 text-foreground ring-1 ring-amber-500"
                      : "bg-muted/20 border-border hover:bg-muted/40 text-muted-foreground"
                  }`}
                >
                  <div className="p-1.5 rounded-full bg-amber-500/20 text-amber-600 dark:text-amber-400 shrink-0 mt-0.5">
                    <AlertTriangle className="w-4 h-4" />
                  </div>
                  <div>
                    <div className="font-semibold text-xs text-foreground">Apto com Restrição</div>
                    <div className="text-[11px] text-muted-foreground mt-0.5">
                      Pode operar com limitações ou sob monitoramento da manutenção.
                    </div>
                  </div>
                </div>

                {/* Inapto */}
                <div
                  onClick={() =>
                    setCurrentChecklist((prev) => ({ ...prev, statusGeral: "inapto" }))
                  }
                  className={`p-3.5 rounded-xl border cursor-pointer transition-all flex items-start gap-3 ${
                    currentChecklist.statusGeral === "inapto"
                      ? "bg-destructive/10 border-destructive text-foreground ring-1 ring-destructive"
                      : "bg-muted/20 border-border hover:bg-muted/40 text-muted-foreground"
                  }`}
                >
                  <div className="p-1.5 rounded-full bg-destructive/20 text-destructive shrink-0 mt-0.5">
                    <XCircle className="w-4 h-4" />
                  </div>
                  <div>
                    <div className="font-semibold text-xs text-foreground">Impedido / Inapto</div>
                    <div className="text-[11px] text-muted-foreground mt-0.5">
                      Risco iminente. Bloqueado até reparo técnico pela equipe mecânica.
                    </div>
                  </div>
                </div>
              </div>

              {/* General observations text area */}
              <div className="space-y-1.5">
                <label className="text-[11px] font-semibold text-foreground">
                  Observações Gerais / Avarias / Peças Substituídas
                </label>
                <textarea
                  rows={3}
                  value={currentChecklist.observacoesGerais}
                  onChange={(e) =>
                    setCurrentChecklist((prev) => ({
                      ...prev,
                      observacoesGerais: e.target.value,
                    }))
                  }
                  placeholder="Registre qualquer detalhe relevante sobre o estado de conservação ou necessidades de reparo..."
                  className="w-full text-xs p-3 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                />
              </div>
            </div>

            {/* Bottom Actions Bar */}
            <div className="flex flex-wrap items-center justify-between gap-3 p-4 bg-card border border-border rounded-xl shadow-sm">
              <div className="flex flex-col sm:flex-row sm:items-center gap-1.5 sm:gap-3 text-xs text-muted-foreground">
                <div>
                  Inspetor:{" "}
                  <strong className="text-foreground">
                    {currentChecklist.operador || "Não informado"}
                  </strong>{" "}
                  • Equipamento:{" "}
                  <strong className="text-foreground">
                    {currentChecklist.equipamento || "Não informado"}
                  </strong>
                </div>
                {lastAutoSaveText && (
                  <span className="inline-flex items-center gap-1 text-[11px] text-emerald-600 dark:text-emerald-400 font-medium">
                    <ShieldCheck className="w-3.5 h-3.5" />
                    Rascunho salvo ({lastAutoSaveText})
                  </span>
                )}
              </div>

              <div className="flex flex-wrap items-center gap-2.5">
                {isChecklistMeaningful(currentChecklist) && (
                  <button
                    type="button"
                    onClick={handleNewChecklist}
                    className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-background hover:bg-muted text-muted-foreground hover:text-destructive border border-border shadow-xs transition-all"
                    title="Descartar preenchimento atual e começar novo"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Limpar</span>
                  </button>
                )}

                <button
                  type="button"
                  onClick={() => handleExportPdf(currentChecklist)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold rounded-lg bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border shadow-xs transition-all"
                  title="Baixar folha de inspeção formatada em PDF com assinaturas"
                >
                  <Download className="w-4 h-4" />
                  <span>Exportar PDF</span>
                </button>

                <button
                  type="button"
                  onClick={handleSaveChecklist}
                  disabled={isSaving}
                  className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold rounded-lg bg-primary hover:bg-primary/90 text-primary-foreground shadow-xs transition-all"
                >
                  {isSaving ? (
                    <>
                      <div className="w-3.5 h-3.5 border-2 border-primary-foreground border-t-transparent rounded-full animate-spin" />
                      <span>Salvando...</span>
                    </>
                  ) : (
                    <>
                      <Save className="w-4 h-4" />
                      <span>Finalizar & Salvar</span>
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: DASHBOARD DOS ITENS DO CHECKLIST */}
        {activeTab === "dashboard" && (
          <ChecklistDashboard
            checklists={allChecklists}
            onSelectChecklist={(item) => {
              setViewingChecklist(item);
            }}
            onUpdateChecklist={(updated) => {
              setAllChecklists((prev) => prev.map((c) => (c.id === updated.id ? updated : c)));
              setCurrentChecklist((curr) => (curr.id === updated.id ? updated : curr));
              setViewingChecklist((viewing) =>
                viewing && viewing.id === updated.id ? updated : viewing,
              );
            }}
          />
        )}

        {/* TAB 3: HISTÓRICO DE CHECKLISTS */}
        {activeTab === "historico" && (
          <div className="space-y-4">
            <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-border/60">
                <div>
                  <h3 className="font-semibold text-foreground text-sm sm:text-base">
                    Histórico de Checklists Realizados
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Consulte, visualize detalhes e exporte checklists gravados
                  </p>
                </div>

                <div className="relative w-full sm:w-72">
                  <Search className="w-3.5 h-3.5 text-muted-foreground absolute left-3 top-2.5" />
                  <input
                    type="text"
                    value={historySearch}
                    onChange={(e) => setHistorySearch(e.target.value)}
                    placeholder="Buscar por máquina, operador..."
                    className="w-full text-xs pl-8 pr-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
                  />
                </div>
              </div>

              {filteredHistory.length === 0 ? (
                <div className="text-center py-12 text-muted-foreground text-xs">
                  <ClipboardCheck className="w-10 h-10 mx-auto mb-2 opacity-50 text-primary" />
                  <p className="font-semibold text-foreground">Nenhum checklist encontrado</p>
                  <p className="mt-1">Preencha e salve seu primeiro checklist na aba ao lado.</p>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="border-b border-border text-muted-foreground font-medium">
                        <th className="py-2.5 px-3">Data / Hora</th>
                        <th className="py-2.5 px-3">Equipamento</th>
                        <th className="py-2.5 px-3">Prefixo / KM</th>
                        <th className="py-2.5 px-3">Operador</th>
                        <th className="py-2.5 px-3">Status</th>
                        <th className="py-2.5 px-3">Conformidade</th>
                        <th className="py-2.5 px-3 text-right">Ações</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-border/60">
                      {filteredHistory.map((item) => {
                        const totalChecked = (item.items || []).filter(
                          (i) => i.status !== "nao_se_aplica",
                        ).length;
                        const totalConforme = (item.items || []).filter(
                          (i) => i.status === "conforme",
                        ).length;
                        const pct =
                          totalChecked > 0 ? Math.round((totalConforme / totalChecked) * 100) : 100;

                        return (
                          <tr key={item.id} className="hover:bg-muted/30 transition-colors">
                            <td className="py-3 px-3 font-mono text-[11px] whitespace-nowrap">
                              {item.data ? item.data.split("-").reverse().join("/") : "—"}{" "}
                              <span className="text-muted-foreground">{item.hora || ""}</span>
                            </td>
                            <td className="py-3 px-3 font-semibold text-foreground whitespace-nowrap">
                              {item.equipamento || "Sem descrição"}
                            </td>
                            <td className="py-3 px-3 text-muted-foreground whitespace-nowrap text-[11px]">
                              <div className="font-mono font-semibold text-foreground">
                                {item.prefixo || "—"}
                              </div>
                              <div className="text-[10px] text-muted-foreground flex flex-wrap gap-x-1.5 font-mono">
                                {item.quilometragemCaminhao && (
                                  <span>KM: {item.quilometragemCaminhao}</span>
                                )}
                                {item.horimetroMotor && <span>Mot: {item.horimetroMotor}</span>}
                                {item.horimetroImplementoFerroviario && (
                                  <span>Impl: {item.horimetroImplementoFerroviario}</span>
                                )}
                                {item.horimetroGerador && <span>Ger: {item.horimetroGerador}</span>}
                                {!item.quilometragemCaminhao &&
                                  !item.horimetroMotor &&
                                  item.horimetroOuKm && <span>{item.horimetroOuKm}</span>}
                              </div>
                            </td>
                            <td className="py-3 px-3 text-foreground whitespace-nowrap">
                              {item.operador || "—"}
                            </td>
                            <td className="py-3 px-3 whitespace-nowrap">
                              <span
                                className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${
                                  item.statusGeral === "apto"
                                    ? "bg-emerald-500/10 text-emerald-600 border-emerald-500/30"
                                    : item.statusGeral === "atencao"
                                      ? "bg-amber-500/10 text-amber-600 border-amber-500/30"
                                      : "bg-destructive/10 text-destructive border-destructive/30"
                                }`}
                              >
                                {item.statusGeral === "apto"
                                  ? "Apto"
                                  : item.statusGeral === "atencao"
                                    ? "Atenção"
                                    : "Inapto"}
                              </span>
                            </td>
                            <td className="py-3 px-3 whitespace-nowrap">
                              <span className="font-mono font-semibold text-xs">{pct}%</span>
                              <span className="text-[10px] text-muted-foreground ml-1">
                                ({totalConforme}/{totalChecked})
                              </span>
                              {((item.fotosCount && item.fotosCount > 0) ||
                                (item.fotos && item.fotos.length > 0)) && (
                                <span
                                  className="inline-flex items-center gap-1 text-[10px] text-emerald-600 dark:text-emerald-400 bg-emerald-500/10 border border-emerald-500/20 px-1.5 py-0.5 rounded font-medium ml-1.5"
                                  title={`${item.fotosCount || item.fotos?.length} fotos anexadas`}
                                >
                                  <Camera className="w-3 h-3" />
                                  <span>{item.fotosCount || item.fotos?.length}</span>
                                </span>
                              )}
                            </td>
                            <td className="py-3 px-3 text-right whitespace-nowrap">
                              <div className="flex items-center justify-end gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setViewingChecklist(item)}
                                  className="p-1.5 rounded-md hover:bg-muted text-muted-foreground hover:text-foreground"
                                  title="Ver detalhes do checklist"
                                >
                                  <Eye className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleExportPdf(item)}
                                  className="p-1.5 rounded-md hover:bg-muted text-muted-foreground hover:text-foreground"
                                  title="Baixar em PDF"
                                >
                                  <Download className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteChecklist(item.id)}
                                  className="p-1.5 rounded-md hover:bg-destructive/10 text-muted-foreground hover:text-destructive"
                                  title="Excluir checklist"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        )}

        {/* TAB 4: PERSONALIZAR ITENS DO CHECKLIST */}
        {activeTab === "personalizar" && (
          <div className="space-y-5">
            <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-4">
              <div className="pb-3 border-b border-border/60">
                <h3 className="font-semibold text-foreground text-sm sm:text-base">
                  Personalização dos Itens de Checagem
                </h3>
                <p className="text-xs text-muted-foreground mt-0.5">
                  Você pode adicionar novos itens ou retirar itens existentes para adequar o
                  checklist ao padrão da sua ferrovia ou frota de máquinas.
                </p>
              </div>

              {/* Add item form */}
              <div className="p-3.5 rounded-xl bg-muted/20 border border-border space-y-3">
                <h4 className="font-semibold text-xs text-foreground">Adicionar Novo Item</h4>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div className="space-y-1">
                    <label className="text-[11px] font-medium text-muted-foreground">
                      Nome do Item
                    </label>
                    <input
                      type="text"
                      value={newItemName}
                      onChange={(e) => setNewItemName(e.target.value)}
                      placeholder="Ex: Nível de fluido de transmissão..."
                      className="w-full text-xs px-3 py-2 bg-background border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary"
                    />
                  </div>

                  <div className="space-y-1">
                    <label className="text-[11px] font-medium text-muted-foreground">
                      Categoria Técnica
                    </label>
                    <input
                      type="text"
                      list="categorias-sugestoes"
                      value={newItemCategory}
                      onChange={(e) => setNewItemCategory(e.target.value)}
                      placeholder="Ex: Motor e Mecânica..."
                      className="w-full text-xs px-3 py-2 bg-background border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary"
                    />
                    <datalist id="categorias-sugestoes">
                      <option value="Motor e Mecânica" />
                      <option value="Freios e Direção" />
                      <option value="Segurança e Sinalização" />
                      <option value="Sistema Elétrico e Iluminação" />
                      <option value="Estrutura e Cabine" />
                      <option value="Implementos e Ferramentas" />
                    </datalist>
                  </div>

                  <div className="flex items-end gap-2">
                    <button
                      type="button"
                      onClick={handleAddCustomItem}
                      className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-primary text-primary-foreground hover:bg-primary/90 shadow-xs transition-colors"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Inserir Item</span>
                    </button>
                  </div>
                </div>
              </div>

              {/* Items List Header with Restore Button */}
              <div className="space-y-2">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                  <h4 className="font-semibold text-xs text-foreground">
                    Itens Atualmente no Modelo ({customItems.length} itens)
                  </h4>

                  <button
                    type="button"
                    onClick={handleResetToVrfTemplate}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-semibold rounded-lg bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border transition-colors shadow-xs self-start sm:self-auto"
                    title="Restaura os 20 itens originais do formulário oficial VRF-SM PH-1016"
                  >
                    <RotateCcw className="w-3.5 h-3.5 text-primary" />
                    <span>Restaurar Padrão Oficial VRF-SM PH-1016 (20 Itens)</span>
                  </button>
                </div>

                <div className="divide-y divide-border/60 max-h-96 overflow-y-auto pr-1">
                  {customItems.map((item, idx) => (
                    <div
                      key={item.id}
                      className="py-2.5 flex items-center justify-between gap-3 text-xs"
                    >
                      <div className="flex items-center gap-2">
                        <span className="font-mono text-[10px] font-bold text-muted-foreground w-6 text-center">
                          {String(idx + 1).padStart(2, "0")}
                        </span>
                        <div>
                          <span className="font-semibold text-foreground">{item.name}</span>
                          <span className="text-[10px] text-muted-foreground ml-2 px-1.5 py-0.5 rounded bg-muted/60">
                            {item.category}
                          </span>
                          {item.required && (
                            <span className="text-[10px] text-destructive ml-1.5 font-bold">
                              (Obrigatório)
                            </span>
                          )}
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleRemoveCustomItem(item.id)}
                        className="p-1 rounded text-muted-foreground hover:text-destructive hover:bg-destructive/10"
                        title="Remover item do modelo"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Modal: Visualização Completa de um Checklist */}
      {viewingChecklist && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4 animate-in fade-in"
          onClick={() => setViewingChecklist(null)}
        >
          <div
            className="relative max-w-3xl w-full max-h-[90vh] flex flex-col bg-card border border-border rounded-xl shadow-2xl overflow-hidden"
            onClick={(e) => e.stopPropagation()}
          >
            {/* Modal Header */}
            <div className="p-4 border-b border-border flex items-center justify-between bg-muted/20">
              <div className="flex items-center gap-2">
                <ClipboardCheck className="w-4 h-4 text-primary" />
                <h4 className="font-semibold text-sm text-foreground">
                  Checklist: {viewingChecklist.equipamento} (
                  {viewingChecklist.prefixo || "Sem prefixo"})
                </h4>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => handleExportPdf(viewingChecklist)}
                  className="inline-flex items-center gap-1 text-xs px-2.5 py-1 rounded bg-secondary hover:bg-secondary/80 border border-border font-medium"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>PDF</span>
                </button>
                <button
                  type="button"
                  onClick={() => setViewingChecklist(null)}
                  className="p-1 rounded text-muted-foreground hover:text-foreground"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Modal Content */}
            <div className="p-4 overflow-y-auto space-y-4 text-xs">
              {/* Meta Grid */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-muted/20 p-3 rounded-lg border border-border/70">
                <div>
                  <span className="text-[10px] text-muted-foreground block">Data / Hora</span>
                  <span className="font-semibold font-mono">
                    {viewingChecklist.data
                      ? viewingChecklist.data.split("-").reverse().join("/")
                      : "—"}{" "}
                    {viewingChecklist.hora || ""}
                  </span>
                </div>
                <div>
                  <span className="text-[10px] text-muted-foreground block">
                    Operador / Inspetor
                  </span>
                  <span className="font-semibold">{viewingChecklist.operador || "—"}</span>
                </div>
                <div>
                  <span className="text-[10px] text-muted-foreground block">Turno / Local</span>
                  <span className="font-semibold">
                    {viewingChecklist.turno || "—"}{" "}
                    {viewingChecklist.local ? `• ${viewingChecklist.local}` : ""}
                  </span>
                </div>
                <div>
                  <span className="text-[10px] text-muted-foreground block">Status Geral</span>
                  <span
                    className={`font-bold ${
                      viewingChecklist.statusGeral === "apto"
                        ? "text-emerald-600 dark:text-emerald-400"
                        : viewingChecklist.statusGeral === "atencao"
                          ? "text-amber-600 dark:text-amber-400"
                          : "text-destructive"
                    }`}
                  >
                    {viewingChecklist.statusGeral === "apto"
                      ? "Apto p/ Operação"
                      : viewingChecklist.statusGeral === "atencao"
                        ? "Apto c/ Restrição"
                        : "Inapto / Impedido"}
                  </span>
                </div>
              </div>

              {/* Horímetros e Quilometragem no Modal */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-muted/30 p-3 rounded-lg border border-border/70">
                <div>
                  <span className="text-[10px] text-muted-foreground block">Horímetro Motor</span>
                  <span className="font-semibold font-mono">
                    {viewingChecklist.horimetroMotor || viewingChecklist.horimetroOuKm || "—"}
                  </span>
                </div>
                <div>
                  <span className="text-[10px] text-muted-foreground block">
                    H. Impl. Ferroviário
                  </span>
                  <span className="font-semibold font-mono">
                    {viewingChecklist.horimetroImplementoFerroviario || "—"}
                  </span>
                </div>
                <div>
                  <span className="text-[10px] text-muted-foreground block">Horímetro Gerador</span>
                  <span className="font-semibold font-mono">
                    {viewingChecklist.horimetroGerador || "—"}
                  </span>
                </div>
                <div>
                  <span className="text-[10px] text-muted-foreground block">KM Caminhão</span>
                  <span className="font-semibold font-mono">
                    {viewingChecklist.quilometragemCaminhao || "—"}
                  </span>
                </div>
              </div>

              {/* Items List */}
              <div className="space-y-2">
                <h5 className="font-semibold text-xs text-foreground">Itens Checados</h5>
                <div className="divide-y divide-border/60 border border-border/70 rounded-lg overflow-hidden">
                  {(viewingChecklist.items || []).map((item, idx) => (
                    <div
                      key={idx}
                      className="p-2.5 flex items-center justify-between gap-2 hover:bg-muted/20"
                    >
                      <div>
                        <span className="font-medium text-foreground">{item.name}</span>
                        <span className="text-[10px] text-muted-foreground ml-2">
                          ({item.category})
                        </span>
                        {item.notes && (
                          <p className="text-[11px] text-destructive italic mt-0.5">
                            Obs: "{item.notes}"
                          </p>
                        )}
                        {item.acaoSolucao && (
                          <p className="text-[11px] text-emerald-600 dark:text-emerald-400 font-medium mt-0.5 flex items-center gap-1">
                            <Check className="w-3 h-3 shrink-0" />
                            <span>Ação p/ solucionar: "{item.acaoSolucao}"</span>
                          </p>
                        )}
                      </div>
                      <span
                        className={`text-[10px] font-bold px-2 py-0.5 rounded ${
                          item.status === "conforme"
                            ? "bg-emerald-500/10 text-emerald-600 border border-emerald-500/30"
                            : item.status === "nao_conforme"
                              ? "bg-destructive/10 text-destructive border border-destructive/30"
                              : "bg-muted text-muted-foreground"
                        }`}
                      >
                        {item.status === "conforme"
                          ? "OK"
                          : item.status === "nao_conforme"
                            ? "NOK"
                            : "N/A"}
                      </span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Observations */}
              {viewingChecklist.observacoesGerais && (
                <div className="p-3 bg-muted/20 border border-border/70 rounded-lg">
                  <h5 className="font-semibold text-[11px] text-foreground mb-1">
                    Observações Gerais:
                  </h5>
                  <p className="text-xs text-muted-foreground italic">
                    "{viewingChecklist.observacoesGerais}"
                  </p>
                </div>
              )}

              {/* Anexo Fotográfico */}
              {Array.isArray(viewingChecklist.fotos) && viewingChecklist.fotos.length > 0 && (
                <div className="space-y-2.5 pt-2 border-t border-border/70">
                  <div className="flex items-center justify-between">
                    <h5 className="font-semibold text-xs text-foreground flex items-center gap-1.5">
                      <Camera className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                      <span>Evidências Fotográficas ({viewingChecklist.fotos.length})</span>
                    </h5>
                    <span className="text-[10px] text-muted-foreground">Clique para ampliar</span>
                  </div>
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                    {viewingChecklist.fotos.map((ph, pi) => (
                      <div
                        key={ph.id || pi}
                        onClick={() => setExpandedPhotoModalUrl(ph.dataUrl)}
                        className="group relative rounded-xl border border-border overflow-hidden bg-card shadow-2xs hover:shadow-md transition cursor-pointer flex flex-col"
                      >
                        <div className="relative aspect-4/3 overflow-hidden bg-muted/40">
                          <img
                            src={ph.dataUrl}
                            alt={ph.caption || `Foto #${pi + 1}`}
                            className="w-full h-full object-cover group-hover:scale-105 transition-transform"
                          />
                          <div className="absolute top-1.5 left-1.5 px-1.5 py-0.5 rounded bg-black/70 text-white text-[9px] font-mono">
                            #{pi + 1}
                          </div>
                        </div>
                        <div className="p-2 text-left bg-card flex-1 flex flex-col justify-between">
                          {ph.itemName && (
                            <p className="text-[10px] font-semibold text-primary truncate mb-0.5">
                              {ph.itemName}
                            </p>
                          )}
                          <p className="text-[11px] text-muted-foreground truncate">
                            {ph.caption || "Sem legenda"}
                          </p>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Lightbox para ampliação de fotos */}
      {expandedPhotoModalUrl && (
        <div
          className="fixed inset-0 z-60 flex items-center justify-center bg-black/85 p-4 animate-in fade-in"
          onClick={() => setExpandedPhotoModalUrl(null)}
        >
          <div
            className="relative max-w-4xl w-full bg-card rounded-2xl overflow-hidden shadow-2xl border border-border flex flex-col max-h-[92vh]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="p-3 border-b border-border flex items-center justify-between bg-muted/40">
              <span className="text-xs font-semibold text-foreground flex items-center gap-2">
                <Camera className="w-4 h-4 text-emerald-600" />
                Evidência Fotográfica do Checklist
              </span>
              <button
                type="button"
                onClick={() => setExpandedPhotoModalUrl(null)}
                className="p-1.5 rounded-lg hover:bg-muted text-muted-foreground hover:text-foreground"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex-1 bg-black flex items-center justify-center p-2 sm:p-4 overflow-auto min-h-[300px]">
              <img
                src={expandedPhotoModalUrl}
                alt="Foto ampliada"
                className="max-h-[78vh] w-auto object-contain rounded-lg"
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
