import { useEffect, useMemo, useState } from "react";
import {
  FileText,
  FileSpreadsheet,
  Trash2,
  Eye,
  Search,
  Filter,
  Download,
  Cloud,
  Check,
  X,
  Clock,
  User,
  MapPin,
  Building2,
  Calendar,
  Camera,
  Sparkles,
  RefreshCw,
  Database,
} from "lucide-react";
import {
  type SavedCloudReport,
  cloudReportToHeaderAndEvents,
  deleteCloudReport,
  subscribeCloudReports,
  deduplicateCloudReports,
  syncCloudReportsAcrossDevices,
} from "@/lib/cloud-reports";
import { exportAllReportsToExcel, exportSingleReportToExcel } from "@/lib/excel-export";
import { SapIw41Modal } from "@/components/SapIw41Modal";
import { SupabaseMigrationModal } from "@/components/SupabaseMigrationModal";
import { downloadReportPdf } from "@/lib/pdf-export";
import {
  formatDate,
  formatDuration,
  formatTime,
  EVENT_LABEL,
  sortAndNormalizeEvents,
  generatePdfFilename,
} from "@/lib/parte-diaria";

interface Props {
  onLoadReportToForm?: (report: SavedCloudReport) => void;
}

export function CloudReportsDatabase({ onLoadReportToForm }: Props) {
  const [reports, setReports] = useState<SavedCloudReport[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [selectedLocal, setSelectedLocal] = useState("");
  const [previewReport, setPreviewReport] = useState<SavedCloudReport | null>(null);
  const [sapModalReport, setSapModalReport] = useState<SavedCloudReport | null>(null);
  const [toastMsg, setToastMsg] = useState<string | null>(null);
  const [isDeduplicating, setIsDeduplicating] = useState(false);
  const [isManualSyncing, setIsManualSyncing] = useState(false);
  const [isSupabaseModalOpen, setIsSupabaseModalOpen] = useState(false);

  useEffect(() => {
    const unsub = subscribeCloudReports((data) => {
      setReports(data);
    });
    return unsub;
  }, []);

  const showToast = (msg: string) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(null), 3500);
  };

  const handleManualSync = async () => {
    setIsManualSyncing(true);
    try {
      const { reports: updated } = await syncCloudReportsAcrossDevices();
      setReports(updated);
      showToast(
        `✓ Sincronizado! ${updated.length} relatório(s) sincronizados entre Computador e Celular.`,
      );
    } catch (e) {
      console.warn("Manual sync error:", e);
      showToast("Aviso ao sincronizar relatórios.");
    } finally {
      setIsManualSyncing(false);
    }
  };

  const handleDeduplicate = async () => {
    setIsDeduplicating(true);
    try {
      const result = await deduplicateCloudReports();
      if (result.duplicatesRemoved > 0) {
        showToast(
          `✨ ${result.duplicatesRemoved} relatório(s) repetido(s) eliminado(s)! Mantido somente o último sincronizado de cada dia.`,
        );
      } else {
        showToast("✓ Banco otimizado: Não há duplicidades no mesmo dia.");
      }
    } catch (e) {
      console.error(e);
      showToast("Aviso ao limpar duplicados da nuvem.");
    } finally {
      setIsDeduplicating(false);
    }
  };

  // Extrair locais únicos para filtro
  const uniqueLocals = useMemo(() => {
    const locs = new Set<string>();
    reports.forEach((r) => {
      if (r.local && r.local.trim()) locs.add(r.local.trim());
    });
    return Array.from(locs).sort();
  }, [reports]);

  // Filtragem dos relatórios
  const filteredReports = useMemo(() => {
    return reports.filter((r) => {
      // Termo de busca geral
      if (searchTerm.trim()) {
        const term = searchTerm.toLowerCase().trim();
        const matches =
          (r.nome && r.nome.toLowerCase().includes(term)) ||
          (r.local && r.local.toLowerCase().includes(term)) ||
          (r.atividade && r.atividade.toLowerCase().includes(term)) ||
          (r.empresa && r.empresa.toLowerCase().includes(term)) ||
          (r.om && r.om.toLowerCase().includes(term)) ||
          (r.sinais && r.sinais.toLowerCase().includes(term)) ||
          (r.data && r.data.includes(term)) ||
          (r.colaboradores && r.colaboradores.toLowerCase().includes(term)) ||
          (r.descritivo && r.descritivo.toLowerCase().includes(term)) ||
          (r.id && r.id.toLowerCase().includes(term));

        if (!matches) return false;
      }

      // Filtro por Data Inicial
      if (startDate && r.data && r.data < startDate) return false;

      // Filtro por Data Final
      if (endDate && r.data && r.data > endDate) return false;

      // Filtro por Local
      if (selectedLocal && r.local !== selectedLocal) return false;

      return true;
    });
  }, [reports, searchTerm, startDate, endDate, selectedLocal]);

  // Estatísticas do Banco de Dados Filtrado
  const stats = useMemo(() => {
    const totalCount = filteredReports.length;
    const totalProd = filteredReports.reduce((acc, r) => acc + (r.totals?.producaoTotal || 0), 0);
    const totalDefects = filteredReports.reduce(
      (acc, r) => acc + (r.totals?.defeitosTotal || 0),
      0,
    );
    const totalJornadaMs = filteredReports.reduce((acc, r) => acc + (r.totals?.jornada || 0), 0);

    return { totalCount, totalProd, totalDefects, totalJornadaMs };
  }, [filteredReports]);

  const handleDelete = async (r: SavedCloudReport) => {
    if (
      confirm(
        `Tem certeza que deseja excluir permanentemente o relatório de ${formatDate(r.data)} (${r.nome || "sem nome"}) da nuvem?`,
      )
    ) {
      try {
        setReports((prev) => prev.filter((item) => item.id !== r.id));
        await deleteCloudReport(r.id);
        showToast("Relatório removido do banco de dados na nuvem.");
        if (previewReport?.id === r.id) setPreviewReport(null);
      } catch (err) {
        console.error("Erro ao deletar relatório:", err);
        showToast("Erro ao excluir relatório da nuvem.");
      }
    }
  };

  const handleDownloadSinglePdf = (r: SavedCloudReport) => {
    const { header, events, photos } = cloudReportToHeaderAndEvents(r);
    const filename = generatePdfFilename(header, events);
    downloadReportPdf(header, events, photos, filename);
    showToast(`Baixando PDF de ${r.nome || "Relatório"}...`);
  };

  const handleDownloadSingleExcel = (r: SavedCloudReport) => {
    exportSingleReportToExcel(r);
    showToast(`Baixando Excel de ${r.nome || "Relatório"}...`);
  };

  const handleExportAllExcel = () => {
    if (reports.length === 0) {
      alert("Não há relatórios salvos no banco de dados na nuvem.");
      return;
    }
    exportAllReportsToExcel(
      reports,
      `banco-de-dados-geral-${new Date().toISOString().slice(0, 10)}.xlsx`,
    );
    showToast("Exportando banco de dados completo para Excel...");
  };

  const handleExportFilteredExcel = () => {
    if (filteredReports.length === 0) {
      alert("Nenhum relatório encontrado para os filtros selecionados.");
      return;
    }
    exportAllReportsToExcel(
      filteredReports,
      `banco-de-dados-filtrado-${new Date().toISOString().slice(0, 10)}.xlsx`,
    );
    showToast("Exportando relatórios filtrados para Excel...");
  };

  return (
    <div className="space-y-6">
      {/* Toast alert */}
      {toastMsg && (
        <div className="fixed bottom-4 right-4 z-50 flex items-center gap-2 rounded-lg bg-emerald-600 px-4 py-3 text-sm font-medium text-white shadow-lg animate-in fade-in slide-in-from-bottom-2">
          <Check className="h-4 w-4 shrink-0" />
          <span>{toastMsg}</span>
        </div>
      )}

      {/* Modern Cloud Header */}
      <div className="rounded-xl border border-primary/20 bg-card p-5 shadow-sm">
        <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
          <div className="flex items-start gap-3">
            <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-primary/10 text-primary">
              <Cloud className="h-5 w-5" />
            </div>
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <h2 className="text-lg font-semibold tracking-tight text-foreground">
                  Banco de Dados de Relatórios na Nuvem
                </h2>
                <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/10 px-2 py-0.5 text-[10px] font-semibold text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
                  <Check className="h-3 w-3" />
                  Último sincronizado por dia
                </span>
                <span className="inline-flex items-center gap-1 rounded-full bg-blue-500/10 px-2 py-0.5 text-[10px] font-semibold text-blue-600 dark:text-blue-400 border border-blue-500/20">
                  <RefreshCw className="h-3 w-3" />
                  Sincronismo PC & Celular Ativo
                </span>
              </div>
              <p className="text-xs text-muted-foreground mt-0.5">
                Central de visualização, busca e exportação de todos os relatórios da equipe em
                formato Excel e PDF.
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setIsSupabaseModalOpen(true)}
              className="inline-flex items-center gap-1.5 rounded-lg border border-emerald-500/30 bg-emerald-500/10 hover:bg-emerald-500/20 text-emerald-700 dark:text-emerald-400 py-2 px-3 text-xs font-semibold shadow-sm transition-colors"
              title="Gerenciar migração e sincronização com o banco de dados Supabase"
            >
              <Database className="h-3.5 w-3.5 text-emerald-600 dark:text-emerald-400" />
              <span>Migração Supabase</span>
            </button>

            <button
              type="button"
              onClick={handleManualSync}
              disabled={isManualSyncing}
              className="inline-flex items-center gap-1.5 rounded-lg border border-border bg-background hover:bg-muted text-foreground py-2 px-3 text-xs font-semibold shadow-sm transition-colors disabled:opacity-50"
              title="Sincroniza instantaneamente os relatórios entre computador e celular"
            >
              <RefreshCw
                className={`h-3.5 w-3.5 ${isManualSyncing ? "animate-spin text-primary" : ""}`}
              />
              <span>{isManualSyncing ? "Sincronizando..." : "Sincronizar PC / Celular"}</span>
            </button>

            <button
              type="button"
              onClick={handleDeduplicate}
              disabled={isDeduplicating || reports.length === 0}
              className="inline-flex items-center gap-1.5 rounded-lg border border-amber-500/30 bg-amber-500/10 hover:bg-amber-500/20 text-amber-600 dark:text-amber-400 py-2 px-3 text-xs font-semibold shadow-sm transition-colors disabled:opacity-50"
              title="Garante que não haja relatórios repetidos no mesmo dia, deixando somente o último sincronizado no banco em nuvem"
            >
              <Sparkles
                className={`h-3.5 w-3.5 ${isDeduplicating ? "animate-spin text-amber-500" : ""}`}
              />
              <span>{isDeduplicating ? "Limpando..." : "Eliminar Repetidos (Mesmo Dia)"}</span>
            </button>

            <button
              type="button"
              onClick={handleExportAllExcel}
              disabled={reports.length === 0}
              className="btn-primary inline-flex items-center gap-2 py-2 px-3.5 text-xs font-semibold shadow-sm disabled:opacity-50"
              title="Baixar todos os relatórios salvos na nuvem em uma planilha Excel completa com abas de resumo e eventos"
            >
              <FileSpreadsheet className="h-4 w-4" />
              <span>Baixar Banco Geral (Excel .xlsx)</span>
            </button>

            {filteredReports.length < reports.length && (
              <button
                type="button"
                onClick={handleExportFilteredExcel}
                className="btn-secondary inline-flex items-center gap-2 py-2 px-3 text-xs"
                title="Baixar apenas os resultados filtrados em planilha Excel"
              >
                <Download className="h-3.5 w-3.5" />
                <span>Baixar Filtrados ({filteredReports.length})</span>
              </button>
            )}
          </div>
        </div>

        {/* Dynamic Metric Badges */}
        <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-4 border-t border-border pt-4 text-xs">
          <div className="rounded-lg bg-secondary/50 p-2.5">
            <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
              Total de Relatórios
            </span>
            <span className="text-base font-bold text-foreground font-mono">
              {stats.totalCount} de {reports.length}
            </span>
          </div>
          <div className="rounded-lg bg-secondary/50 p-2.5">
            <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
              Produção Acumulada
            </span>
            <span className="text-base font-bold text-emerald-600 font-mono">
              {stats.totalProd} m
            </span>
          </div>
          <div className="rounded-lg bg-secondary/50 p-2.5">
            <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
              Defeitos Acumulados
            </span>
            <span className="text-base font-bold text-amber-600 font-mono">
              {stats.totalDefects}
            </span>
          </div>
          <div className="rounded-lg bg-secondary/50 p-2.5">
            <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
              Jornada Total
            </span>
            <span className="text-base font-bold text-foreground font-mono">
              {formatDuration(stats.totalJornadaMs)}
            </span>
          </div>
        </div>
      </div>

      {/* Filter and Search Panel */}
      <div className="rounded-xl border border-border bg-card p-4 shadow-sm space-y-3">
        <div className="flex flex-col md:flex-row items-center gap-3">
          {/* Main Search Input */}
          <div className="relative flex-1 w-full">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Pesquisar por responsável, local, atividade, empresa, OM..."
              className="input pl-9 pr-8 text-xs w-full"
            />
            {searchTerm && (
              <button
                type="button"
                onClick={() => setSearchTerm("")}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            )}
          </div>

          {/* Quick Filters */}
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 w-full md:w-auto shrink-0">
            <div>
              <select
                value={selectedLocal}
                onChange={(e) => setSelectedLocal(e.target.value)}
                className="input text-xs cursor-pointer w-full"
              >
                <option value="">-- Todos os Locais --</option>
                {uniqueLocals.map((loc) => (
                  <option key={loc} value={loc}>
                    {loc}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex items-center gap-1">
              <input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="input text-xs w-full"
                title="Data inicial"
              />
            </div>

            <div className="flex items-center gap-1 col-span-2 sm:col-span-1">
              <input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="input text-xs w-full"
                title="Data final"
              />
            </div>
          </div>
        </div>

        {/* Active Filters Bar */}
        {(searchTerm || startDate || endDate || selectedLocal) && (
          <div className="flex items-center justify-between gap-2 pt-1 border-t border-border/60 text-xs text-muted-foreground">
            <div className="flex flex-wrap items-center gap-1.5">
              <Filter className="h-3 w-3 text-primary shrink-0" />
              <span>Filtros ativos:</span>
              {searchTerm && (
                <span className="rounded bg-accent px-2 py-0.5 font-medium text-foreground">
                  "{searchTerm}"
                </span>
              )}
              {selectedLocal && (
                <span className="rounded bg-accent px-2 py-0.5 font-medium text-foreground">
                  Local: {selectedLocal}
                </span>
              )}
              {startDate && (
                <span className="rounded bg-accent px-2 py-0.5 font-medium text-foreground">
                  De: {formatDate(startDate)}
                </span>
              )}
              {endDate && (
                <span className="rounded bg-accent px-2 py-0.5 font-medium text-foreground">
                  Até: {formatDate(endDate)}
                </span>
              )}
            </div>

            <button
              type="button"
              onClick={() => {
                setSearchTerm("");
                setStartDate("");
                setEndDate("");
                setSelectedLocal("");
              }}
              className="text-xs text-primary hover:underline font-medium shrink-0"
            >
              Limpar Filtros
            </button>
          </div>
        )}
      </div>

      {/* Reports Table / List */}
      <div className="rounded-xl border border-border bg-card shadow-sm overflow-hidden">
        {filteredReports.length === 0 ? (
          <div className="p-8 text-center text-muted-foreground space-y-2">
            <Cloud className="h-10 w-10 mx-auto text-muted-foreground/40 stroke-1" />
            <p className="text-sm font-medium">Nenhum relatório encontrado na nuvem.</p>
            <p className="text-xs text-muted-foreground max-w-md mx-auto">
              {reports.length === 0
                ? "Preencha o formulário e clique em 'Salvar Relatório na Nuvem' para registrar os dados."
                : "Tente alterar os termos da busca ou limpar os filtros de data e local."}
            </p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead className="bg-secondary/60 text-muted-foreground border-b border-border font-semibold uppercase tracking-wider text-[10px]">
                <tr>
                  <th className="py-3 px-3">Data / Envio</th>
                  <th className="py-3 px-3">Responsável</th>
                  <th className="py-3 px-3">Atividade / Empresa</th>
                  <th className="py-3 px-3">Local / OM</th>
                  <th className="py-3 px-3 text-center">Produção</th>
                  <th className="py-3 px-3 text-center">Defeitos</th>
                  <th className="py-3 px-3 text-center">Jornada</th>
                  <th className="py-3 px-3 text-right">Ações (Excel / PDF)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border/60">
                {filteredReports.map((r) => (
                  <tr key={r.id} className="hover:bg-accent/40 transition-colors">
                    <td className="py-3 px-3 font-medium text-foreground whitespace-nowrap">
                      <div className="flex items-center gap-1.5 font-mono text-xs">
                        <Calendar className="h-3.5 w-3.5 text-primary shrink-0" />
                        <span>{formatDate(r.data)}</span>
                      </div>
                      <div className="text-[10px] text-muted-foreground font-mono mt-0.5">
                        {new Date(r.createdAt).toLocaleTimeString("pt-BR", {
                          hour: "2-digit",
                          minute: "2-digit",
                        })}
                      </div>
                    </td>

                    <td className="py-3 px-3">
                      <div className="font-semibold text-foreground flex items-center gap-1">
                        <User className="h-3 w-3 text-muted-foreground shrink-0" />
                        <span className="truncate max-w-[140px]">{r.nome || "Não informado"}</span>
                      </div>
                      {r.colaboradores && (
                        <div
                          className="text-[10px] text-muted-foreground truncate max-w-[140px]"
                          title={r.colaboradores}
                        >
                          Equipe: {r.colaboradores}
                        </div>
                      )}
                    </td>

                    <td className="py-3 px-3">
                      <div
                        className="font-medium text-foreground truncate max-w-[180px] flex items-center gap-1.5"
                        title={r.atividade}
                      >
                        <span className="truncate">{r.atividade || "Sem nome de atividade"}</span>
                        {r.photos && r.photos.length > 0 && (
                          <span
                            className="inline-flex items-center gap-0.5 text-[10px] font-medium bg-primary/10 text-primary px-1.5 py-0.2 rounded shrink-0"
                            title={`${r.photos.length} foto(s) anexada(s)`}
                          >
                            <Camera className="h-3 w-3" />
                            <span>{r.photos.length}</span>
                          </span>
                        )}
                      </div>
                      {r.empresa && (
                        <div className="text-[10px] text-muted-foreground flex items-center gap-1 mt-0.5">
                          <Building2 className="h-3 w-3 shrink-0" />
                          <span className="truncate max-w-[160px]">{r.empresa}</span>
                        </div>
                      )}
                    </td>

                    <td className="py-3 px-3">
                      <div className="font-medium text-foreground flex items-center gap-1">
                        <MapPin className="h-3 w-3 text-amber-500 shrink-0" />
                        <span>{r.local || "Não inf."}</span>
                      </div>
                      <div className="text-[10px] text-muted-foreground font-mono mt-0.5">
                        {[
                          r.linha ? `Linha: ${r.linha}` : null,
                          r.via ? `Via: ${r.via}` : null,
                          r.om ? `OM: ${r.om}` : null,
                        ]
                          .filter(Boolean)
                          .join(" • ")}
                      </div>
                    </td>

                    <td className="py-3 px-3 text-center font-mono font-bold text-emerald-600">
                      {r.totals?.producaoTotal || 0} m
                    </td>

                    <td className="py-3 px-3 text-center font-mono font-bold text-amber-600">
                      {r.totals?.defeitosTotal || 0}
                    </td>

                    <td className="py-3 px-3 text-center font-mono text-muted-foreground whitespace-nowrap">
                      {formatDuration(r.totals?.jornada || 0)}
                    </td>

                    <td className="py-3 px-3 text-right whitespace-nowrap">
                      <div className="flex items-center justify-end gap-1.5">
                        {/* Download Single Excel */}
                        <button
                          type="button"
                          onClick={() => handleDownloadSingleExcel(r)}
                          className="p-1.5 rounded-md border border-emerald-500/30 bg-emerald-50 text-emerald-700 hover:bg-emerald-100 dark:bg-emerald-950/40 dark:text-emerald-300 transition-colors"
                          title="Baixar Relatório em Excel (.xlsx)"
                        >
                          <FileSpreadsheet className="h-3.5 w-3.5" />
                        </button>

                        {/* Download SAP IW41 */}
                        <button
                          type="button"
                          onClick={() => setSapModalReport(r)}
                          className="p-1.5 rounded-md border border-emerald-600/40 bg-emerald-600/15 text-emerald-800 dark:text-emerald-300 hover:bg-emerald-600 hover:text-white transition-colors"
                          title="Gerar planilha de apropriação SAP IW41 (.xlsx)"
                        >
                          <FileSpreadsheet className="h-3.5 w-3.5" />
                        </button>

                        {/* Download Single PDF */}
                        <button
                          type="button"
                          onClick={() => handleDownloadSinglePdf(r)}
                          className="p-1.5 rounded-md border border-sky-500/30 bg-sky-50 text-sky-700 hover:bg-sky-100 dark:bg-sky-950/40 dark:text-sky-300 transition-colors"
                          title="Baixar Relatório em PDF"
                        >
                          <FileText className="h-3.5 w-3.5" />
                        </button>

                        {/* Preview detail */}
                        <button
                          type="button"
                          onClick={() => setPreviewReport(r)}
                          className="p-1.5 rounded-md border border-border bg-secondary hover:bg-accent text-foreground transition-colors"
                          title="Visualizar Detalhes"
                        >
                          <Eye className="h-3.5 w-3.5" />
                        </button>

                        {/* Load into form */}
                        {onLoadReportToForm && (
                          <button
                            type="button"
                            onClick={() => {
                              onLoadReportToForm(r);
                              showToast("Relatório carregado na tela principal!");
                            }}
                            className="px-2 py-1 rounded-md text-[10px] font-semibold bg-primary/10 text-primary hover:bg-primary hover:text-white transition-colors"
                            title="Carregar dados deste relatório na tela principal"
                          >
                            Editar
                          </button>
                        )}

                        {/* Delete */}
                        <button
                          type="button"
                          onClick={() => handleDelete(r)}
                          className="p-1.5 rounded-md border border-destructive/30 bg-destructive/10 text-destructive hover:bg-destructive hover:text-white transition-colors"
                          title="Excluir da Nuvem"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Preview Modal */}
      {previewReport && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 backdrop-blur-sm animate-in fade-in">
          <div className="relative w-full max-w-3xl max-h-[90vh] overflow-y-auto rounded-xl border border-border bg-card p-6 shadow-2xl space-y-5">
            <div className="flex items-center justify-between border-b border-border pb-3">
              <div className="flex items-center gap-2">
                <Cloud className="h-5 w-5 text-primary" />
                <h3 className="text-base font-bold text-foreground">
                  Detalhes do Relatório em Nuvem
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setPreviewReport(null)}
                className="rounded-lg p-1.5 text-muted-foreground hover:bg-accent hover:text-foreground"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            {/* Header info grid */}
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 bg-secondary/40 p-4 rounded-lg text-xs">
              <div>
                <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
                  Data do Relatório
                </span>
                <span className="font-bold text-foreground font-mono">
                  {formatDate(previewReport.data)}
                </span>
              </div>
              <div>
                <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
                  Responsável
                </span>
                <span className="font-bold text-foreground">{previewReport.nome || "-"}</span>
              </div>
              <div>
                <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
                  Local
                </span>
                <span className="font-bold text-foreground">{previewReport.local || "-"}</span>
              </div>
              <div>
                <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
                  Linha / Via
                </span>
                <span className="font-medium text-foreground">
                  {previewReport.linha || "-"} / {previewReport.via || "-"}
                </span>
              </div>
              <div>
                <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
                  GDA / OM
                </span>
                <span className="font-medium text-foreground">
                  {previewReport.gda || "-"} / {previewReport.om || "-"}
                </span>
              </div>
              <div>
                <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
                  Empresa
                </span>
                <span className="font-medium text-foreground">{previewReport.empresa || "-"}</span>
              </div>
              <div className="col-span-2 sm:col-span-3">
                <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
                  Nome da Atividade
                </span>
                <span className="font-bold text-primary text-sm">
                  {previewReport.atividade || "-"}
                </span>
              </div>
              {previewReport.colaboradores && (
                <div className="col-span-2 sm:col-span-3">
                  <span className="text-muted-foreground block text-[10px] uppercase font-semibold">
                    Colaboradores
                  </span>
                  <span className="text-foreground">{previewReport.colaboradores}</span>
                </div>
              )}
            </div>

            {/* Event Timeline Table */}
            <div>
              <h4 className="text-xs font-semibold uppercase tracking-wider text-muted-foreground mb-2">
                Linha do Tempo de Eventos ({previewReport.events?.length || 0})
              </h4>
              <div className="rounded-lg border border-border overflow-hidden text-xs max-h-60 overflow-y-auto">
                <table className="w-full text-left">
                  <thead className="bg-secondary text-muted-foreground text-[10px] uppercase font-semibold">
                    <tr>
                      <th className="py-2 px-3">Início</th>
                      <th className="py-2 px-3">Fim</th>
                      <th className="py-2 px-3">Duração</th>
                      <th className="py-2 px-3">Evento</th>
                      <th className="py-2 px-3">Km / Prod.</th>
                      <th className="py-2 px-3">Descrição / Obs</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border">
                    {(() => {
                      const sorted = sortAndNormalizeEvents(previewReport.events || []);
                      return sorted.map((ev, i) => {
                        const next = sorted[i + 1];
                        const startStr = formatTime(ev.timestamp);
                        let endStr = "";
                        let durStr = "";
                        if (next) {
                          endStr = formatTime(next.timestamp);
                          durStr = formatDuration(Math.max(0, next.timestamp - ev.timestamp));
                        } else if (ev.kind === "fim_jornada") {
                          endStr = formatTime(ev.timestamp);
                          durStr = "00:00:00";
                        } else {
                          endStr = "Em andamento";
                          durStr = "(em andamento)";
                        }
                        return (
                          <tr key={ev.id} className="hover:bg-accent/30">
                            <td className="py-2 px-3 font-mono text-[11px] whitespace-nowrap">
                              {startStr}
                            </td>
                            <td className="py-2 px-3 font-mono text-[11px] whitespace-nowrap text-muted-foreground">
                              {endStr}
                            </td>
                            <td className="py-2 px-3 font-mono text-[11px] whitespace-nowrap font-semibold">
                              {durStr}
                            </td>
                            <td className="py-2 px-3 font-medium">
                              {EVENT_LABEL[ev.kind] || ev.kind}
                            </td>
                            <td className="py-2 px-3 font-mono">
                              {[
                                ev.kmInicial ? `km ${ev.kmInicial}` : null,
                                ev.producao ? `${ev.producao}m` : null,
                                ev.defeitos ? `${ev.defeitos}def` : null,
                              ]
                                .filter(Boolean)
                                .join(" • ") || "-"}
                            </td>
                            <td className="py-2 px-3 text-muted-foreground">
                              {[ev.descricao, ev.observacoes].filter(Boolean).join(" — ") || "-"}
                            </td>
                          </tr>
                        );
                      });
                    })()}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Photos Section in Modal */}
            {previewReport.photos && previewReport.photos.length > 0 && (
              <div className="space-y-2">
                <div className="flex items-center gap-2">
                  <Camera className="h-4 w-4 text-primary" />
                  <h4 className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                    Anexo Fotográfico ({previewReport.photos.length}{" "}
                    {previewReport.photos.length === 1 ? "foto" : "fotos"})
                  </h4>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
                  {previewReport.photos.map((p, idx) => (
                    <div
                      key={p.id || idx}
                      className="border border-border rounded-lg overflow-hidden bg-muted/30 p-1"
                    >
                      <div className="aspect-4/3 relative overflow-hidden rounded bg-black/5 flex items-center justify-center">
                        <img
                          src={p.dataUrl}
                          alt={`Foto ${idx + 1}`}
                          className="w-full h-full object-contain"
                          referrerPolicy="no-referrer"
                        />
                        <span className="absolute top-1 left-1 bg-background/90 text-foreground font-bold text-[9px] px-1.5 py-0.5 rounded shadow-xs">
                          Foto {(idx + 1).toString().padStart(2, "0")}
                        </span>
                      </div>
                      {p.caption && (
                        <p
                          className="text-[10px] text-muted-foreground p-1 truncate"
                          title={p.caption}
                        >
                          {p.caption}
                        </p>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Action buttons */}
            <div className="flex flex-wrap items-center justify-between gap-3 border-t border-border pt-4">
              <div className="text-xs text-muted-foreground font-mono">ID: {previewReport.id}</div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={() => setSapModalReport(previewReport)}
                  className="btn-secondary inline-flex items-center gap-1.5 text-xs py-1.5 px-3 font-semibold text-emerald-700 dark:text-emerald-400 border-emerald-500/40"
                  title="Gerar planilha de apropriação de mão de obra para SAP IW41"
                >
                  <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                  <span>Apropriação SAP (IW41)</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleDownloadSingleExcel(previewReport)}
                  className="btn-secondary inline-flex items-center gap-1.5 text-xs py-1.5 px-3"
                >
                  <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                  <span>Baixar Excel</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleDownloadSinglePdf(previewReport)}
                  className="btn-primary inline-flex items-center gap-1.5 text-xs py-1.5 px-3"
                >
                  <FileText className="h-4 w-4" />
                  <span>Baixar PDF</span>
                </button>

                <button
                  type="button"
                  onClick={() => setPreviewReport(null)}
                  className="btn-ghost text-xs py-1.5 px-3"
                >
                  Fechar
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      {/* SAP IW41 Modal */}
      {sapModalReport && (
        <SapIw41Modal
          isOpen={!!sapModalReport}
          onClose={() => setSapModalReport(null)}
          header={cloudReportToHeaderAndEvents(sapModalReport).header}
          events={cloudReportToHeaderAndEvents(sapModalReport).events}
          onSuccess={(filename) => showToast(`✓ Planilha SAP IW41 gerada com sucesso: ${filename}`)}
        />
      )}

      {/* Supabase Migration Modal */}
      <SupabaseMigrationModal
        isOpen={isSupabaseModalOpen}
        onClose={() => setIsSupabaseModalOpen(false)}
        onMigrationComplete={() => {
          showToast("✓ Migração para o Supabase concluída com sucesso!");
          handleManualSync();
        }}
      />
    </div>
  );
}
