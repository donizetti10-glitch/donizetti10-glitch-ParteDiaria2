import { useState, useMemo } from "react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend,
} from "recharts";
import {
  type EquipmentChecklist,
  type EquipmentOperatingStatus,
  type ChecklistItem,
  saveChecklistToCloud,
  saveLocalChecklist,
} from "@/lib/checklist";
import {
  AlertTriangle,
  CheckCircle2,
  AlertCircle,
  ClipboardList,
  Filter,
  RotateCcw,
  Search,
  Wrench,
  Calendar,
  Truck,
  TrendingDown,
  ShieldCheck,
  Check,
  Save,
  Clock,
  ExternalLink,
  Loader2,
  Sparkles,
} from "lucide-react";

interface ChecklistDashboardProps {
  checklists: EquipmentChecklist[];
  onSelectChecklist?: (checklist: EquipmentChecklist) => void;
  onUpdateChecklist?: (checklist: EquipmentChecklist) => void;
}

const STATUS_COLORS: Record<EquipmentOperatingStatus, string> = {
  apto: "#16a34a", // Emerald / Green
  atencao: "#d97706", // Amber
  inapto: "#dc2626", // Red
};

export function ChecklistDashboard({
  checklists,
  onSelectChecklist,
  onUpdateChecklist,
}: ChecklistDashboardProps) {
  const [searchTerm, setSearchTerm] = useState("");
  const [periodFilter, setPeriodFilter] = useState<"todos" | "7dias" | "30dias" | "mes_atual">(
    "todos",
  );
  const [statusFilter, setStatusFilter] = useState<"todos" | EquipmentOperatingStatus>("todos");
  const [equipmentFilter, setEquipmentFilter] = useState<string>("todos");

  // State for corrective actions on non-conforming items
  const [actionDrafts, setActionDrafts] = useState<Record<string, string>>({});
  const [savingActionKey, setSavingActionKey] = useState<string | null>(null);
  const [savedActionKey, setSavedActionKey] = useState<string | null>(null);
  const [defectFilter, setDefectFilter] = useState<"todos" | "pendente" | "resolvido">("todos");
  const [defectSearch, setDefectSearch] = useState("");
  const [feedbackToast, setFeedbackToast] = useState<{
    message: string;
    type: "success" | "error";
  } | null>(null);

  // Dynamic list of unique equipment
  const equipmentOptions = useMemo(() => {
    const set = new Set<string>();
    checklists.forEach((c) => {
      const label = c.prefixo ? `${c.prefixo} - ${c.equipamento}` : c.equipamento;
      if (label.trim()) set.add(label.trim());
    });
    return Array.from(set).sort();
  }, [checklists]);

  // Filter checklists
  const filteredChecklists = useMemo(() => {
    const now = new Date();
    const todayStr = now.toISOString().slice(0, 10);

    return checklists.filter((c) => {
      // Period filter
      if (periodFilter === "7dias") {
        const cDate = new Date(c.data);
        const diffDays = (now.getTime() - cDate.getTime()) / (1000 * 3600 * 24);
        if (diffDays > 7) return false;
      } else if (periodFilter === "30dias") {
        const cDate = new Date(c.data);
        const diffDays = (now.getTime() - cDate.getTime()) / (1000 * 3600 * 24);
        if (diffDays > 30) return false;
      } else if (periodFilter === "mes_atual") {
        const cMonth = c.data?.slice(0, 7);
        const currentMonth = todayStr.slice(0, 7);
        if (cMonth !== currentMonth) return false;
      }

      // Status filter
      if (statusFilter !== "todos" && c.statusGeral !== statusFilter) {
        return false;
      }

      // Equipment filter
      if (equipmentFilter !== "todos") {
        const label = c.prefixo ? `${c.prefixo} - ${c.equipamento}` : c.equipamento;
        if (label.trim() !== equipmentFilter) return false;
      }

      // Search term
      if (searchTerm.trim()) {
        const term = searchTerm.toLowerCase();
        const matchEquip = (c.equipamento || "").toLowerCase().includes(term);
        const matchPrefixo = (c.prefixo || "").toLowerCase().includes(term);
        const matchOp = (c.operador || "").toLowerCase().includes(term);
        const matchLocal = (c.local || "").toLowerCase().includes(term);
        const matchObs = (c.observacoesGerais || "").toLowerCase().includes(term);
        if (!matchEquip && !matchPrefixo && !matchOp && !matchLocal && !matchObs) {
          return false;
        }
      }

      return true;
    });
  }, [checklists, periodFilter, statusFilter, equipmentFilter, searchTerm]);

  // KPI Calculations
  const stats = useMemo(() => {
    const total = filteredChecklists.length;
    let aptoCount = 0;
    let atencaoCount = 0;
    let inaptoCount = 0;

    let totalItemsChecked = 0;
    let totalConforme = 0;
    let totalNaoConforme = 0;
    let totalNA = 0;

    filteredChecklists.forEach((c) => {
      if (c.statusGeral === "apto") aptoCount++;
      else if (c.statusGeral === "atencao") atencaoCount++;
      else if (c.statusGeral === "inapto") inaptoCount++;

      (c.items || []).forEach((item) => {
        if (item.status === "conforme") {
          totalConforme++;
          totalItemsChecked++;
        } else if (item.status === "nao_conforme") {
          totalNaoConforme++;
          totalItemsChecked++;
        } else if (item.status === "nao_se_aplica") {
          totalNA++;
        }
      });
    });

    const conformidadeTaxa =
      totalItemsChecked > 0 ? Math.round((totalConforme / totalItemsChecked) * 100) : 100;

    const disponibilidadeTaxa =
      total > 0 ? Math.round(((aptoCount + atencaoCount) / total) * 100) : 100;

    return {
      total,
      aptoCount,
      atencaoCount,
      inaptoCount,
      totalItemsChecked,
      totalConforme,
      totalNaoConforme,
      totalNA,
      conformidadeTaxa,
      disponibilidadeTaxa,
    };
  }, [filteredChecklists]);

  // Category compliance chart data
  const categoryData = useMemo(() => {
    const catMap: Record<
      string,
      { category: string; conforme: number; naoConforme: number; total: number }
    > = {};

    filteredChecklists.forEach((c) => {
      (c.items || []).forEach((item) => {
        const cat = item.category || "Outros";
        if (!catMap[cat]) {
          catMap[cat] = { category: cat, conforme: 0, naoConforme: 0, total: 0 };
        }
        if (item.status === "conforme") {
          catMap[cat].conforme++;
          catMap[cat].total++;
        } else if (item.status === "nao_conforme") {
          catMap[cat].naoConforme++;
          catMap[cat].total++;
        }
      });
    });

    return Object.values(catMap)
      .map((item) => ({
        ...item,
        taxaConforme: item.total > 0 ? Math.round((item.conforme / item.total) * 100) : 100,
      }))
      .sort((a, b) => a.taxaConforme - b.taxaConforme); // Most problematic first
  }, [filteredChecklists]);

  // Status Pie Chart Data
  const statusPieData = useMemo(() => {
    return [
      { name: "Apto para Operação", value: stats.aptoCount, color: STATUS_COLORS.apto },
      { name: "Apto c/ Restrição", value: stats.atencaoCount, color: STATUS_COLORS.atencao },
      { name: "Inapto / Impedido", value: stats.inaptoCount, color: STATUS_COLORS.inapto },
    ].filter((item) => item.value > 0);
  }, [stats]);

  // Top non-conforming items ranking
  const topIssues = useMemo(() => {
    const issueMap: Record<
      string,
      {
        name: string;
        category: string;
        count: number;
        equipments: Set<string>;
        latestNotes: string[];
      }
    > = {};

    filteredChecklists.forEach((c) => {
      const eqLabel = c.prefixo ? `${c.prefixo} (${c.equipamento})` : c.equipamento;
      (c.items || []).forEach((item) => {
        if (item.status === "nao_conforme") {
          if (!issueMap[item.name]) {
            issueMap[item.name] = {
              name: item.name,
              category: item.category,
              count: 0,
              equipments: new Set(),
              latestNotes: [],
            };
          }
          issueMap[item.name].count++;
          if (eqLabel) issueMap[item.name].equipments.add(eqLabel);
          if (item.notes && item.notes.trim()) {
            issueMap[item.name].latestNotes.push(
              `${eqLabel || "Equipamento"}: ${item.notes.trim()}`,
            );
          }
        }
      });
    });

    return Object.values(issueMap)
      .sort((a, b) => b.count - a.count)
      .slice(0, 7);
  }, [filteredChecklists]);

  // All non-conforming items from filtered checklists with full item & action tracking
  const allDefects = useMemo(() => {
    const list: {
      key: string;
      checklistId: string;
      itemId: string;
      date: string;
      hora: string;
      equipamento: string;
      prefixo: string;
      operador: string;
      turno: string;
      local: string;
      itemName: string;
      category: string;
      notes: string;
      acaoSolucao: string;
      rawChecklist: EquipmentChecklist;
    }[] = [];

    filteredChecklists.forEach((c) => {
      (c.items || []).forEach((item, itemIdx) => {
        if (item.status === "nao_conforme") {
          const validItemId = item.id || `item-${itemIdx}`;
          list.push({
            key: `${c.id}_${validItemId}`,
            checklistId: c.id,
            itemId: validItemId,
            date: c.data,
            hora: c.hora,
            equipamento: c.equipamento,
            prefixo: c.prefixo,
            operador: c.operador,
            turno: c.turno,
            local: c.local,
            itemName: item.name,
            category: item.category,
            notes: item.notes || "Sem observação detalhada",
            acaoSolucao: item.acaoSolucao || "",
            rawChecklist: c,
          });
        }
      });
    });

    return list;
  }, [filteredChecklists]);

  // Filtered defects for the action cards/table
  const filteredDefects = useMemo(() => {
    return allDefects.filter((d) => {
      // Filter by action resolution status
      if (defectFilter === "pendente" && d.acaoSolucao.trim()) return false;
      if (defectFilter === "resolvido" && !d.acaoSolucao.trim()) return false;

      // Filter by search query
      if (defectSearch.trim()) {
        const term = defectSearch.toLowerCase();
        const matches =
          d.itemName.toLowerCase().includes(term) ||
          d.category.toLowerCase().includes(term) ||
          d.equipamento.toLowerCase().includes(term) ||
          d.prefixo.toLowerCase().includes(term) ||
          d.operador.toLowerCase().includes(term) ||
          d.notes.toLowerCase().includes(term) ||
          d.acaoSolucao.toLowerCase().includes(term);
        if (!matches) return false;
      }

      return true;
    });
  }, [allDefects, defectFilter, defectSearch]);

  const pendingDefectsCount = useMemo(
    () => allDefects.filter((d) => !d.acaoSolucao.trim()).length,
    [allDefects],
  );

  const resolvedDefectsCount = useMemo(
    () => allDefects.filter((d) => !!d.acaoSolucao.trim()).length,
    [allDefects],
  );

  // Save corrective action for a non-conforming item
  const handleSaveAction = async (
    checklistId: string,
    itemId: string,
    itemName: string,
    explicitText?: string,
  ) => {
    const key = `${checklistId}_${itemId}`;

    const targetChecklist =
      checklists.find((c) => c.id === checklistId) ||
      allDefects.find((d) => d.checklistId === checklistId)?.rawChecklist;

    if (!targetChecklist) {
      setFeedbackToast({
        message: "Checklist não encontrado para atualização.",
        type: "error",
      });
      setTimeout(() => setFeedbackToast(null), 3000);
      return;
    }

    const existingItem = (targetChecklist.items || []).find(
      (item, idx) =>
        item.id === itemId || (!item.id && `item-${idx}` === itemId) || item.name === itemName,
    );

    const actionText = (
      explicitText !== undefined
        ? explicitText
        : actionDrafts[key] !== undefined
          ? actionDrafts[key]
          : existingItem?.acaoSolucao || ""
    ).trim();

    const updatedItems: ChecklistItem[] = (targetChecklist.items || []).map((item, idx) => {
      const match =
        item.id === itemId || (!item.id && `item-${idx}` === itemId) || item.name === itemName;
      return {
        id: item.id || `item-${idx}`,
        category: item.category || "Geral",
        name: item.name || "",
        description: item.description || "",
        required: item.required ?? true,
        status: item.status || "conforme",
        notes: item.notes || "",
        acaoSolucao: match ? actionText : item.acaoSolucao || "",
      };
    });

    const updatedChecklist: EquipmentChecklist = {
      ...targetChecklist,
      id: targetChecklist.id || checklistId,
      equipamento: targetChecklist.equipamento || "",
      prefixo: targetChecklist.prefixo || "",
      operador: targetChecklist.operador || "",
      turno: targetChecklist.turno || "",
      local: targetChecklist.local || "",
      statusGeral: targetChecklist.statusGeral || "apto",
      observacoesGerais: targetChecklist.observacoesGerais || "",
      horimetroMotor: targetChecklist.horimetroMotor || "",
      horimetroImplementoFerroviario: targetChecklist.horimetroImplementoFerroviario || "",
      horimetroGerador: targetChecklist.horimetroGerador || "",
      quilometragemCaminhao: targetChecklist.quilometragemCaminhao || "",
      horimetroOuKm: targetChecklist.horimetroOuKm || "",
      items: updatedItems,
      updatedAt: Date.now(),
    };

    setSavingActionKey(key);
    try {
      const saved = await saveChecklistToCloud(updatedChecklist);
      if (onUpdateChecklist) {
        onUpdateChecklist(saved || updatedChecklist);
      }
      setActionDrafts((prev) => ({ ...prev, [key]: actionText }));
      setSavedActionKey(key);
      setFeedbackToast({
        message: actionText
          ? `✓ Ação salva com sucesso para "${itemName}"!`
          : `Ação em "${itemName}" limpa com sucesso.`,
        type: "success",
      });
      setTimeout(() => setSavedActionKey(null), 2500);
      setTimeout(() => setFeedbackToast(null), 3500);
    } catch (err) {
      console.error("Erro ao salvar ação corretiva:", err);
      saveLocalChecklist(updatedChecklist);
      if (onUpdateChecklist) {
        onUpdateChecklist(updatedChecklist);
      }
      setActionDrafts((prev) => ({ ...prev, [key]: actionText }));
      setSavedActionKey(key);
      setFeedbackToast({
        message: `✓ Ação salva localmente no dispositivo!`,
        type: "success",
      });
      setTimeout(() => setSavedActionKey(null), 2500);
      setTimeout(() => setFeedbackToast(null), 3500);
    } finally {
      setSavingActionKey(null);
    }
  };

  const handleResetFilters = () => {
    setSearchTerm("");
    setPeriodFilter("todos");
    setStatusFilter("todos");
    setEquipmentFilter("todos");
  };

  return (
    <div className="space-y-6">
      {/* Filters Bar */}
      <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-4">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 pb-3 border-b border-border/60">
          <div className="flex items-center gap-2">
            <Filter className="w-4 h-4 text-primary" />
            <h3 className="font-semibold text-sm text-foreground">
              Filtros do Dashboard de Checklist
            </h3>
          </div>
          <button
            type="button"
            onClick={handleResetFilters}
            className="text-xs text-muted-foreground hover:text-foreground flex items-center gap-1.5 self-start md:self-auto py-1 px-2.5 rounded-md hover:bg-muted/60 transition-colors"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Limpar Filtros</span>
          </button>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {/* Search */}
          <div className="space-y-1">
            <label className="text-[11px] font-medium text-muted-foreground">
              Buscar Equipamento / Operador
            </label>
            <div className="relative">
              <Search className="w-3.5 h-3.5 text-muted-foreground absolute left-3 top-2.5" />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Ex: CAT 416, EQ-01, João..."
                className="w-full text-xs pl-8 pr-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
              />
            </div>
          </div>

          {/* Period Filter */}
          <div className="space-y-1">
            <label className="text-[11px] font-medium text-muted-foreground">Período</label>
            <select
              value={periodFilter}
              onChange={(e) =>
                setPeriodFilter(e.target.value as "todos" | "7dias" | "30dias" | "mes_atual")
              }
              className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
            >
              <option value="todos">Todo o Histórico</option>
              <option value="7dias">Últimos 7 dias</option>
              <option value="30dias">Últimos 30 dias</option>
              <option value="mes_atual">Mês Atual</option>
            </select>
          </div>

          {/* Equipment Filter */}
          <div className="space-y-1">
            <label className="text-[11px] font-medium text-muted-foreground">
              Filtrar por Máquina
            </label>
            <select
              value={equipmentFilter}
              onChange={(e) => setEquipmentFilter(e.target.value)}
              className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background truncate"
            >
              <option value="todos">Todos os Equipamentos</option>
              {equipmentOptions.map((opt) => (
                <option key={opt} value={opt}>
                  {opt}
                </option>
              ))}
            </select>
          </div>

          {/* Status Filter */}
          <div className="space-y-1">
            <label className="text-[11px] font-medium text-muted-foreground">
              Status Operacional
            </label>
            <select
              value={statusFilter}
              onChange={(e) =>
                setStatusFilter(e.target.value as "todos" | EquipmentOperatingStatus)
              }
              className="w-full text-xs px-3 py-2 bg-muted/30 border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background"
            >
              <option value="todos">Todos os Status</option>
              <option value="apto">Apto p/ Operação</option>
              <option value="atencao">Apto c/ Restrição</option>
              <option value="inapto">Inapto / Impedido</option>
            </select>
          </div>
        </div>
      </div>

      {/* KPI Cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4">
        {/* Total Checklists */}
        <div className="bg-card border border-border rounded-xl p-4 shadow-xs relative overflow-hidden">
          <div className="flex items-center justify-between">
            <span className="text-[11px] font-semibold text-muted-foreground uppercase tracking-wider">
              Checklists
            </span>
            <div className="p-2 rounded-lg bg-primary/10 text-primary">
              <ClipboardList className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl sm:text-3xl font-bold font-mono text-foreground">
              {stats.total}
            </span>
            <span className="text-[11px] text-muted-foreground">inspeções</span>
          </div>
          <p className="text-[11px] text-muted-foreground mt-1">
            {stats.totalItemsChecked} itens avaliados no total
          </p>
        </div>

        {/* Taxa de Conformidade */}
        <div className="bg-card border border-border rounded-xl p-4 shadow-xs relative overflow-hidden">
          <div className="flex items-center justify-between">
            <span className="text-[11px] font-semibold text-muted-foreground uppercase tracking-wider">
              Conformidade
            </span>
            <div
              className={`p-2 rounded-lg ${
                stats.conformidadeTaxa >= 90
                  ? "bg-emerald-500/10 text-emerald-600 dark:text-emerald-400"
                  : stats.conformidadeTaxa >= 75
                    ? "bg-amber-500/10 text-amber-600 dark:text-amber-400"
                    : "bg-destructive/10 text-destructive"
              }`}
            >
              <ShieldCheck className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span
              className={`text-2xl sm:text-3xl font-bold font-mono ${
                stats.conformidadeTaxa >= 90
                  ? "text-emerald-600 dark:text-emerald-400"
                  : stats.conformidadeTaxa >= 75
                    ? "text-amber-600 dark:text-amber-400"
                    : "text-destructive"
              }`}
            >
              {stats.conformidadeTaxa}%
            </span>
            <span className="text-[11px] text-muted-foreground">de aprovação</span>
          </div>
          <div className="w-full bg-muted/60 h-1.5 rounded-full mt-2 overflow-hidden">
            <div
              className={`h-full rounded-full transition-all duration-500 ${
                stats.conformidadeTaxa >= 90
                  ? "bg-emerald-500"
                  : stats.conformidadeTaxa >= 75
                    ? "bg-amber-500"
                    : "bg-destructive"
              }`}
              style={{ width: `${stats.conformidadeTaxa}%` }}
            />
          </div>
        </div>

        {/* Equipamentos Aptos vs Impedidos */}
        <div className="bg-card border border-border rounded-xl p-4 shadow-xs relative overflow-hidden">
          <div className="flex items-center justify-between">
            <span className="text-[11px] font-semibold text-muted-foreground uppercase tracking-wider">
              Disponibilidade
            </span>
            <div className="p-2 rounded-lg bg-emerald-500/10 text-emerald-600 dark:text-emerald-400">
              <CheckCircle2 className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl sm:text-3xl font-bold font-mono text-emerald-600 dark:text-emerald-400">
              {stats.aptoCount}
            </span>
            <span className="text-xs text-muted-foreground">
              aptos / {stats.inaptoCount} impedidos
            </span>
          </div>
          <p className="text-[11px] text-muted-foreground mt-1">
            {stats.atencaoCount > 0
              ? `${stats.atencaoCount} operando com restrição`
              : "Nenhum em alerta"}
          </p>
        </div>

        {/* Não Conformidades Detectadas */}
        <div className="bg-card border border-border rounded-xl p-4 shadow-xs relative overflow-hidden">
          <div className="flex items-center justify-between">
            <span className="text-[11px] font-semibold text-muted-foreground uppercase tracking-wider">
              Defeitos Apontados
            </span>
            <div
              className={`p-2 rounded-lg ${
                stats.totalNaoConforme > 0
                  ? "bg-destructive/10 text-destructive"
                  : "bg-muted/30 text-muted-foreground"
              }`}
            >
              <AlertTriangle className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span
              className={`text-2xl sm:text-3xl font-bold font-mono ${
                stats.totalNaoConforme > 0 ? "text-destructive" : "text-foreground"
              }`}
            >
              {stats.totalNaoConforme}
            </span>
            <span className="text-[11px] text-muted-foreground">itens reprovados</span>
          </div>
          <p className="text-[11px] text-muted-foreground mt-1">
            Requerem atenção mecânica imediata
          </p>
        </div>
      </div>

      {/* Charts Grid */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        {/* Chart 1: Conformidade por Categoria */}
        <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-3">
          <div className="flex items-center justify-between border-b border-border/60 pb-3">
            <div>
              <h4 className="font-semibold text-sm text-foreground">
                Conformidade por Categoria Técnica
              </h4>
              <p className="text-xs text-muted-foreground">
                Compara aprovações vs reprovações em cada sistema
              </p>
            </div>
          </div>

          {categoryData.length === 0 ? (
            <div className="h-64 flex items-center justify-center text-xs text-muted-foreground">
              Nenhum dado disponível no período selecionado.
            </div>
          ) : (
            <div className="h-64 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={categoryData}
                  layout="vertical"
                  margin={{ left: 10, right: 20, top: 10, bottom: 10 }}
                >
                  <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
                  <XAxis type="number" domain={[0, "dataMax + 2"]} tick={{ fontSize: 11 }} />
                  <YAxis dataKey="category" type="category" tick={{ fontSize: 10 }} width={110} />
                  <Tooltip
                    formatter={(val: unknown, name: unknown) => [
                      String(val ?? 0),
                      String(name) === "conforme" ? "Conforme" : "Não Conforme",
                    ]}
                    labelStyle={{ fontWeight: "bold", fontSize: 12 }}
                  />
                  <Legend
                    formatter={(value) => (value === "conforme" ? "Conforme" : "Não Conforme")}
                    wrapperStyle={{ fontSize: 11 }}
                  />
                  <Bar dataKey="conforme" fill="#16a34a" radius={[0, 4, 4, 0]} stackId="a" />
                  <Bar dataKey="naoConforme" fill="#dc2626" radius={[0, 4, 4, 0]} stackId="a" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>

        {/* Chart 2: Status Geral dos Equipamentos */}
        <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-3">
          <div className="flex items-center justify-between border-b border-border/60 pb-3">
            <div>
              <h4 className="font-semibold text-sm text-foreground">
                Status de Liberação Operacional
              </h4>
              <p className="text-xs text-muted-foreground">
                Proporção de equipamentos aptos, com restrição e impedidos
              </p>
            </div>
          </div>

          {statusPieData.length === 0 ? (
            <div className="h-64 flex items-center justify-center text-xs text-muted-foreground">
              Nenhum checklist registrado ainda.
            </div>
          ) : (
            <div className="h-64 w-full flex flex-col items-center justify-center">
              <ResponsiveContainer width="100%" height="85%">
                <PieChart>
                  <Pie
                    data={statusPieData}
                    cx="50%"
                    cy="50%"
                    innerRadius={55}
                    outerRadius={80}
                    paddingAngle={4}
                    dataKey="value"
                  >
                    {statusPieData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip
                    formatter={(val: unknown) => [`${String(val ?? 0)} checklists`, "Quantidade"]}
                  />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                </PieChart>
              </ResponsiveContainer>
              <div className="text-[11px] text-muted-foreground text-center">
                Total de {stats.total} avaliações realizadas
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Top Defective Items (Pareto / Root Cause) */}
      <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-4">
        <div className="flex items-center justify-between border-b border-border/60 pb-3">
          <div className="flex items-center gap-2">
            <Wrench className="w-4 h-4 text-amber-500" />
            <div>
              <h4 className="font-semibold text-sm text-foreground">
                Principais Itens com Não Conformidades (Mais Frequentes)
              </h4>
              <p className="text-xs text-muted-foreground">
                Itens e componentes que mais apresentaram falhas durante as inspeções de campo
              </p>
            </div>
          </div>
        </div>

        {topIssues.length === 0 ? (
          <div className="text-center py-8 text-xs text-muted-foreground">
            <CheckCircle2 className="w-8 h-8 text-emerald-500 mx-auto mb-2 opacity-80" />
            <p className="font-medium text-foreground">Nenhuma não conformidade encontrada!</p>
            <p className="mt-1">Todos os equipamentos inspecionados estão 100% conformes.</p>
          </div>
        ) : (
          <div className="space-y-3">
            {topIssues.map((issue, idx) => (
              <div
                key={issue.name}
                className="p-3 bg-muted/20 border border-border/70 rounded-lg flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 hover:border-amber-500/50 transition-colors"
              >
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-mono font-bold bg-amber-500/10 text-amber-700 dark:text-amber-400 px-2 py-0.5 rounded border border-amber-500/20">
                      #{idx + 1}
                    </span>
                    <span className="text-xs font-semibold text-foreground">{issue.name}</span>
                    <span className="text-[10px] text-muted-foreground bg-muted/60 px-2 py-0.5 rounded">
                      {issue.category}
                    </span>
                  </div>
                  <div className="text-[11px] text-muted-foreground">
                    Equipamentos afetados:{" "}
                    <span className="text-foreground font-medium">
                      {Array.from(issue.equipments).join(", ") || "—"}
                    </span>
                  </div>
                  {issue.latestNotes.length > 0 && (
                    <div className="text-[11px] text-destructive italic">
                      Última nota: "{issue.latestNotes[0]}"
                    </div>
                  )}
                </div>

                <div className="flex items-center gap-2 shrink-0 self-end sm:self-auto">
                  <span className="text-xs font-mono font-bold px-2.5 py-1 rounded bg-destructive/10 text-destructive border border-destructive/20">
                    {issue.count} {issue.count === 1 ? "defeito" : "defeitos"}
                  </span>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Non-Conforming Items (Não OK) & Corrective Actions Section */}
      <div className="bg-card border border-border rounded-xl p-4 sm:p-5 shadow-xs space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-border/60 pb-3">
          <div className="flex items-start gap-2.5">
            <div className="p-2 rounded-lg bg-destructive/10 text-destructive mt-0.5">
              <AlertTriangle className="w-4 h-4" />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <h4 className="font-semibold text-sm sm:text-base text-foreground">
                  Itens Não Conformes (Não OK) & Ações de Solução
                </h4>
                <span className="text-[11px] font-mono font-bold px-2 py-0.5 rounded-full bg-destructive/10 text-destructive border border-destructive/20">
                  {allDefects.length}{" "}
                  {allDefects.length === 1 ? "item reprovado" : "itens reprovados"}
                </span>
              </div>
              <p className="text-xs text-muted-foreground mt-0.5">
                Insira e gerencie a ação corretiva para cada problema ou anomalia detectada nas
                inspeções
              </p>
            </div>
          </div>

          {/* KPI Summary badges */}
          <div className="flex items-center gap-2 self-start sm:self-auto shrink-0">
            <span
              className={`text-xs px-2.5 py-1 rounded-lg border font-medium flex items-center gap-1.5 ${
                pendingDefectsCount > 0
                  ? "bg-amber-500/10 text-amber-600 dark:text-amber-400 border-amber-500/30"
                  : "bg-muted text-muted-foreground border-border"
              }`}
            >
              <Clock className="w-3 h-3" />
              <span>{pendingDefectsCount} pendentes</span>
            </span>

            <span
              className={`text-xs px-2.5 py-1 rounded-lg border font-medium flex items-center gap-1.5 ${
                resolvedDefectsCount > 0
                  ? "bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border-emerald-500/30"
                  : "bg-muted text-muted-foreground border-border"
              }`}
            >
              <CheckCircle2 className="w-3 h-3" />
              <span>{resolvedDefectsCount} solucionados</span>
            </span>
          </div>
        </div>

        {/* Filter and Search Bar for Non-Conforming Items */}
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 bg-muted/20 p-3 rounded-lg border border-border/70">
          <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar">
            <button
              type="button"
              onClick={() => setDefectFilter("todos")}
              className={`text-xs font-semibold px-3 py-1.5 rounded-md transition-colors whitespace-nowrap ${
                defectFilter === "todos"
                  ? "bg-foreground text-background shadow-xs"
                  : "text-muted-foreground hover:text-foreground hover:bg-muted"
              }`}
            >
              Todos ({allDefects.length})
            </button>
            <button
              type="button"
              onClick={() => setDefectFilter("pendente")}
              className={`text-xs font-semibold px-3 py-1.5 rounded-md transition-colors whitespace-nowrap flex items-center gap-1.5 ${
                defectFilter === "pendente"
                  ? "bg-amber-600 text-white shadow-xs"
                  : "text-muted-foreground hover:text-amber-600 hover:bg-amber-500/10"
              }`}
            >
              <Clock className="w-3 h-3" />
              <span>Aguardando Ação ({pendingDefectsCount})</span>
            </button>
            <button
              type="button"
              onClick={() => setDefectFilter("resolvido")}
              className={`text-xs font-semibold px-3 py-1.5 rounded-md transition-colors whitespace-nowrap flex items-center gap-1.5 ${
                defectFilter === "resolvido"
                  ? "bg-emerald-600 text-white shadow-xs"
                  : "text-muted-foreground hover:text-emerald-600 hover:bg-emerald-500/10"
              }`}
            >
              <Check className="w-3 h-3" />
              <span>Com Ação Registrada ({resolvedDefectsCount})</span>
            </button>
          </div>

          <div className="relative w-full md:w-72">
            <Search className="w-3.5 h-3.5 text-muted-foreground absolute left-3 top-2.5" />
            <input
              type="text"
              value={defectSearch}
              onChange={(e) => setDefectSearch(e.target.value)}
              placeholder="Buscar item, defeito, equipamento..."
              className="w-full text-xs pl-8 pr-3 py-1.5 bg-background border border-input rounded-lg focus:outline-none focus:ring-1 focus:ring-primary text-foreground"
            />
          </div>
        </div>

        {/* Defects Cards List */}
        {allDefects.length === 0 ? (
          <div className="text-center py-10 text-xs text-muted-foreground space-y-1">
            <CheckCircle2 className="w-9 h-9 text-emerald-500 mx-auto opacity-80" />
            <p className="font-semibold text-sm text-foreground">
              Nenhuma não conformidade encontrada no período!
            </p>
            <p>Todos os itens inspecionados estão operando com 100% de conformidade.</p>
          </div>
        ) : filteredDefects.length === 0 ? (
          <div className="text-center py-8 text-xs text-muted-foreground space-y-2">
            <p className="font-medium text-foreground">
              Nenhum item não conforme corresponde aos filtros atuais.
            </p>
            <button
              type="button"
              onClick={() => {
                setDefectFilter("todos");
                setDefectSearch("");
              }}
              className="text-primary hover:underline font-semibold text-xs"
            >
              Limpar filtros da lista
            </button>
          </div>
        ) : (
          <div className="space-y-3.5">
            {filteredDefects.map((d) => {
              const draftValue =
                actionDrafts[d.key] !== undefined ? actionDrafts[d.key] : d.acaoSolucao;
              const isSaving = savingActionKey === d.key;
              const isSaved = savedActionKey === d.key;
              const hasSavedAction = !!d.acaoSolucao.trim();
              const hasDraftChanges =
                actionDrafts[d.key] !== undefined &&
                actionDrafts[d.key].trim() !== d.acaoSolucao.trim();

              return (
                <div
                  key={d.key}
                  className={`p-4 rounded-xl border transition-all ${
                    hasSavedAction
                      ? "bg-card border-border/80 hover:border-border"
                      : "bg-amber-500/5 border-amber-500/30 hover:border-amber-500/50"
                  }`}
                >
                  {/* Card Header */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-3 border-b border-border/60">
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="font-bold text-xs text-foreground bg-muted/60 px-2 py-0.5 rounded">
                        {d.prefixo ? `${d.prefixo} • ` : ""}
                        {d.equipamento || "Equipamento"}
                      </span>
                      <span className="text-[11px] text-muted-foreground font-mono">
                        {d.date ? d.date.split("-").reverse().join("/") : "—"} {d.hora || ""}
                      </span>
                      <span className="text-[11px] text-muted-foreground">
                        Inspetor: <strong className="text-foreground">{d.operador || "—"}</strong>
                      </span>
                      {d.turno && (
                        <span className="text-[10px] text-muted-foreground bg-muted px-1.5 py-0.5 rounded">
                          Turno {d.turno}
                        </span>
                      )}
                    </div>

                    <div className="flex items-center gap-2 self-start sm:self-auto shrink-0">
                      <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-destructive/10 text-destructive border border-destructive/20 uppercase tracking-wider">
                        Não OK (NOK)
                      </span>

                      {hasSavedAction ? (
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/30 flex items-center gap-1">
                          <Check className="w-3 h-3" />
                          <span>Ação Definida</span>
                        </span>
                      ) : (
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-amber-500/10 text-amber-600 dark:text-amber-400 border border-amber-500/30 flex items-center gap-1">
                          <Clock className="w-3 h-3" />
                          <span>Aguardando Ação</span>
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Defective Item Info */}
                  <div className="py-3 space-y-2">
                    <div className="flex items-start justify-between gap-2 flex-wrap">
                      <div>
                        <span className="text-xs font-bold text-foreground">{d.itemName}</span>
                        <span className="text-[10px] text-muted-foreground ml-2 bg-muted/50 px-2 py-0.5 rounded border border-border/50">
                          {d.category}
                        </span>
                      </div>

                      {onSelectChecklist && (
                        <button
                          type="button"
                          onClick={() => onSelectChecklist(d.rawChecklist)}
                          className="text-[11px] text-primary hover:underline inline-flex items-center gap-1 font-medium"
                          title="Abrir checklist completo"
                        >
                          <ExternalLink className="w-3 h-3" />
                          <span>Ver Checklist Completo</span>
                        </button>
                      )}
                    </div>

                    {/* Operator observation */}
                    <div className="p-2.5 rounded-lg bg-destructive/5 border border-destructive/20 text-xs flex items-start gap-2">
                      <AlertCircle className="w-3.5 h-3.5 text-destructive shrink-0 mt-0.5" />
                      <div className="text-foreground">
                        <span className="font-semibold text-destructive">
                          Observação da Inspeção:{" "}
                        </span>
                        <span className="italic">"{d.notes}"</span>
                      </div>
                    </div>

                    {/* Corrective Action Input Area */}
                    <div className="pt-2 space-y-2">
                      <div className="flex items-center justify-between">
                        <label className="text-[11px] font-semibold text-foreground flex items-center gap-1.5">
                          <Wrench className="w-3.5 h-3.5 text-primary" />
                          <span>Ação para Solucionar o Problema (Ação Corretiva):</span>
                        </label>
                        {hasSavedAction && (
                          <span className="text-[10px] text-emerald-600 dark:text-emerald-400 font-medium">
                            ✓ Registrado no checklist
                          </span>
                        )}
                      </div>

                      <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
                        <input
                          type="text"
                          value={draftValue}
                          onChange={(e) =>
                            setActionDrafts((prev) => ({ ...prev, [d.key]: e.target.value }))
                          }
                          onKeyDown={(e) => {
                            if (e.key === "Enter") {
                              e.preventDefault();
                              handleSaveAction(d.checklistId, d.itemId, d.itemName);
                            }
                          }}
                          placeholder="Descreva a ação para solucionar (ex: Completar nível com ARLA 32, Trocar mangueira, Enviar p/ oficina...)"
                          className={`flex-1 text-xs px-3 py-2 bg-background border rounded-lg focus:outline-none focus:ring-1 text-foreground transition-all ${
                            hasDraftChanges
                              ? "border-primary ring-1 ring-primary/30"
                              : "border-input focus:ring-primary"
                          }`}
                        />

                        <button
                          type="button"
                          disabled={isSaving}
                          onClick={() => handleSaveAction(d.checklistId, d.itemId, d.itemName)}
                          className={`text-xs px-4 py-2 rounded-lg font-semibold flex items-center justify-center gap-1.5 shrink-0 transition-all ${
                            isSaved
                              ? "bg-emerald-600 text-white shadow-xs"
                              : hasDraftChanges || !hasSavedAction
                                ? "bg-primary text-primary-foreground hover:bg-primary/90 shadow-xs"
                                : "bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border"
                          }`}
                          title="Salvar ação corretiva no checklist"
                        >
                          {isSaving ? (
                            <>
                              <Loader2 className="w-3.5 h-3.5 animate-spin" />
                              <span>Salvando...</span>
                            </>
                          ) : isSaved ? (
                            <>
                              <Check className="w-3.5 h-3.5" />
                              <span>Salvo!</span>
                            </>
                          ) : (
                            <>
                              <Save className="w-3.5 h-3.5" />
                              <span>Salvar Ação</span>
                            </>
                          )}
                        </button>
                      </div>

                      {/* Quick Action Suggestion Chips */}
                      <div className="flex items-center gap-1.5 flex-wrap pt-0.5">
                        <span className="text-[10px] text-muted-foreground flex items-center gap-1">
                          <Sparkles className="w-3 h-3 text-amber-500" />
                          <span>Sugestões rápidas:</span>
                        </span>
                        {[
                          d.itemName.toLowerCase().includes("arla")
                            ? "Completar nível do reservatório com ARLA 32"
                            : "Completar nível de fluido",
                          "Acionar equipe de manutenção mecânica",
                          "Substituição de peça danificada",
                          "Reaperto e regulagem operacional",
                          "Máquina recolhida para oficina",
                        ].map((sugg) => (
                          <button
                            key={sugg}
                            type="button"
                            onClick={() => {
                              setActionDrafts((prev) => ({ ...prev, [d.key]: sugg }));
                            }}
                            className="text-[10px] px-2 py-0.5 rounded-md bg-muted/60 hover:bg-muted text-muted-foreground hover:text-foreground border border-border/60 transition-colors"
                          >
                            + {sugg}
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Floating Feedback Toast */}
      {feedbackToast && (
        <div
          className={`fixed bottom-5 right-5 z-50 px-4 py-2.5 rounded-xl shadow-lg border text-xs font-semibold flex items-center gap-2 animate-in fade-in slide-in-from-bottom-2 ${
            feedbackToast.type === "success"
              ? "bg-emerald-600 text-white border-emerald-500 shadow-emerald-950/20"
              : "bg-destructive text-white border-destructive/80 shadow-destructive/20"
          }`}
        >
          {feedbackToast.type === "success" ? (
            <CheckCircle2 className="w-4 h-4 text-white" />
          ) : (
            <AlertCircle className="w-4 h-4 text-white" />
          )}
          <span>{feedbackToast.message}</span>
        </div>
      )}
    </div>
  );
}
