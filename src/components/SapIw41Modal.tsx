import React, { useState, useMemo } from "react";
import {
  FileSpreadsheet,
  X,
  Users,
  Clock,
  Briefcase,
  Check,
  Plus,
  Play,
  FileCheck,
  Loader2,
  Sparkles,
} from "lucide-react";
import {
  OFFICIAL_SAP_COLLABORATORS,
  SapInterval,
  buildSapIntervalsFromEvents,
  exportSapIw41ToExcel,
  formatSapDate,
  formatSapDateDashed,
} from "@/lib/sap-iw41-export";
import { WorkEvent } from "@/lib/parte-diaria";

interface SapIw41ModalProps {
  isOpen: boolean;
  onClose: () => void;
  header: {
    om?: string;
    data?: string;
    colaboradores?: string;
    atividade?: string;
    local?: string;
  };
  events?: WorkEvent[];
  onSuccess?: (filename: string) => void;
}

export function SapIw41Modal({
  isOpen,
  onClose,
  header,
  events = [],
  onSuccess,
}: SapIw41ModalProps) {
  // Estado para OM e OP
  const [om, setOm] = useState(header.om || "");
  const [op, setOp] = useState("10");
  const [isExporting, setIsExporting] = useState(false);

  // Colaboradores e matrículas
  const initialCollaborators = useMemo(() => {
    const rawColabs = (header.colaboradores || "").toLowerCase();
    const hasSpecificSelection = rawColabs.trim().length > 0;

    return OFFICIAL_SAP_COLLABORATORS.map((c) => {
      // Se houver colaboradores preenchidos no cabeçalho, seleciona apenas os que constam
      let isSelected = true;
      if (hasSpecificSelection) {
        const firstName = c.nome.split(" ")[0].toLowerCase();
        const lastName = c.nome.split(" ").slice(-1)[0].toLowerCase();
        isSelected =
          rawColabs.includes(c.nome.toLowerCase()) ||
          rawColabs.includes(firstName) ||
          rawColabs.includes(lastName);
      }
      return {
        ...c,
        selecionado: isSelected,
      };
    });
  }, [header.colaboradores]);

  const [collaborators, setCollaborators] = useState(initialCollaborators);
  const [newColabName, setNewColabName] = useState("");
  const [newColabMatricula, setNewColabMatricula] = useState("");

  // Intervalos operacionais detectados dos eventos ou padrão
  const [intervals, setIntervals] = useState<SapInterval[]>(() =>
    buildSapIntervalsFromEvents(events, header.data),
  );

  // Recalcular quando o modal abrir ou eventos mudarem
  React.useEffect(() => {
    if (isOpen) {
      setOm(header.om || "");
      setOp("10");
      setCollaborators(initialCollaborators);
      setIntervals(buildSapIntervalsFromEvents(events, header.data));
      setIsExporting(false);
    }
  }, [isOpen, header.om, header.data, initialCollaborators, events]);

  if (!isOpen) return null;

  const toggleCollaborator = (matricula: string) => {
    setCollaborators((prev) =>
      prev.map((c) => (c.matricula === matricula ? { ...c, selecionado: !c.selecionado } : c)),
    );
  };

  const selectAllCollaborators = () => {
    setCollaborators((prev) => prev.map((c) => ({ ...c, selecionado: true })));
  };

  const deselectAllCollaborators = () => {
    setCollaborators((prev) => prev.map((c) => ({ ...c, selecionado: false })));
  };

  const handleAddCustomCollaborator = (e: React.FormEvent) => {
    e.preventDefault();
    const nome = newColabName.trim();
    const matricula = newColabMatricula.trim().replace(/\D/g, "");
    if (!nome || !matricula) return;

    setCollaborators((prev) => [...prev, { nome, matricula, selecionado: true }]);
    setNewColabName("");
    setNewColabMatricula("");
  };

  const selectedCollaborators = collaborators.filter((c) => c.selecionado);
  const totalRows = selectedCollaborators.length * intervals.length;
  const expectedFilename = `${formatSapDateDashed(header.data)}.xlsm`;

  const handleExport = async () => {
    if (selectedCollaborators.length === 0) {
      alert("Selecione ao menos um colaborador para gerar a apropriação.");
      return;
    }
    if (intervals.length === 0) {
      alert("Não há intervalos de atividades para gerar.");
      return;
    }

    try {
      setIsExporting(true);
      const filename = await exportSapIw41ToExcel({
        om,
        op,
        reportDate: header.data,
        collaborators: selectedCollaborators,
        intervals,
      });

      if (onSuccess) {
        onSuccess(filename);
      }
      onClose();
    } catch (err) {
      console.error("Erro ao gerar arquivo Excel com Macro:", err);
      alert("Ocorreu um erro ao gerar a planilha com macro. Tente novamente.");
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="bg-card text-card-foreground border border-border rounded-xl shadow-2xl w-full max-w-3xl max-h-[92vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-5 py-4 border-b border-border bg-muted/40">
          <div className="flex items-center gap-3">
            <div className="h-10 w-10 rounded-lg bg-emerald-600/15 text-emerald-600 dark:text-emerald-400 flex items-center justify-center font-bold">
              <FileSpreadsheet className="h-5 w-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base font-bold text-foreground">
                  Apropriação de Mão de Obra — SAP IW41
                </h3>
                <span className="bg-purple-600/10 text-purple-700 dark:text-purple-400 border border-purple-500/20 text-[10px] font-bold px-2 py-0.5 rounded-full uppercase tracking-wider flex items-center gap-1">
                  <Play className="h-2.5 w-2.5 fill-current" />
                  Excel .xlsm (Macro VBA)
                </span>
              </div>
              <p className="text-xs text-muted-foreground">
                Arquivo nomeado conforme a data da inspeção ({expectedFilename}) com a macro e botão
                de execução no SAP
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="rounded-lg p-1.5 text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* Content Body */}
        <div className="flex-1 overflow-y-auto p-5 space-y-5 text-sm">
          {/* Parâmetros da Ordem e Operação */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 p-3.5 bg-muted/30 border border-border rounded-lg">
            <div>
              <label className="text-xs font-semibold text-foreground flex items-center gap-1.5 mb-1">
                <Briefcase className="h-3.5 w-3.5 text-primary" />
                <span>Ordem de Manutenção (OM)</span>
              </label>
              <input
                type="text"
                value={om}
                onChange={(e) => setOm(e.target.value)}
                placeholder="Deixar em branco para preencher no Excel"
                className="w-full text-xs font-mono bg-background border border-input rounded-md px-2.5 py-1.5 focus:ring-1 focus:ring-primary outline-hidden"
              />
              <span className="text-[10px] text-muted-foreground mt-0.5 block">
                {om
                  ? "Utilizando OM informada no cabeçalho"
                  : "Coluna Ordem ficará vazia para preenchimento manual na planilha"}
              </span>
            </div>

            <div>
              <label className="text-xs font-semibold text-foreground flex items-center gap-1.5 mb-1">
                <span>Operação (OP)</span>
              </label>
              <input
                type="text"
                value={op}
                onChange={(e) => setOp(e.target.value)}
                placeholder="10"
                className="w-full text-xs font-mono bg-background border border-input rounded-md px-2.5 py-1.5 focus:ring-1 focus:ring-primary outline-hidden"
              />
              <span className="text-[10px] text-muted-foreground mt-0.5 block">
                Padrão SAP: operação 10
              </span>
            </div>

            <div>
              <label className="text-xs font-semibold text-foreground flex items-center gap-1.5 mb-1">
                <Clock className="h-3.5 w-3.5 text-primary" />
                <span>Nome do Arquivo & Aba</span>
              </label>
              <div className="w-full text-xs font-mono bg-muted/60 border border-border rounded-md px-2.5 py-1.5 text-foreground flex items-center justify-between">
                <span className="font-bold text-emerald-600 dark:text-emerald-400 truncate">
                  {expectedFilename}
                </span>
                <span className="text-[10px] text-muted-foreground font-sans shrink-0 ml-1">
                  (com traço)
                </span>
              </div>
              <span className="text-[10px] text-muted-foreground mt-0.5 block">
                Nome da planilha conforme a data da inspeção
              </span>
            </div>
          </div>

          {/* Destaque da Macro Embutida e Atalho */}
          <div className="p-3.5 bg-purple-500/10 border border-purple-500/20 rounded-lg flex items-start gap-3">
            <div className="h-8 w-8 rounded-lg bg-purple-600/20 text-purple-600 dark:text-purple-400 flex items-center justify-center shrink-0 mt-0.5">
              <Play className="h-4 w-4 fill-current ml-0.5" />
            </div>
            <div className="text-xs space-y-1">
              <div className="flex items-center gap-2">
                <span className="font-bold text-foreground">
                  Macro SAP IW41 Embutida com Atalho e Botão na Planilha
                </span>
                <span className="inline-flex items-center gap-1 bg-purple-600/20 text-purple-700 dark:text-purple-300 font-mono text-[10px] px-1.5 py-0.5 rounded font-bold">
                  Ctrl + Shift + R
                </span>
              </div>
              <p className="text-muted-foreground leading-relaxed">
                A planilha gerada inclui a macro <strong>Macro_IW44_TIC</strong> integrada, um botão
                interativo <strong>[▶ EXECUTAR NO SAP]</strong> diretamente na planilha e atalho de
                teclado automático para conectar ao SAP GUI e lançar o HH sem esforço manual.
              </p>
            </div>
          </div>

          {/* Equipe e Matrículas */}
          <div className="space-y-2">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Users className="h-4 w-4 text-primary" />
                <h4 className="text-xs font-bold uppercase tracking-wider text-foreground">
                  Colaboradores e Matrículas ({selectedCollaborators.length} de{" "}
                  {collaborators.length} selecionados)
                </h4>
              </div>
              <div className="flex items-center gap-2 text-xs">
                <button
                  type="button"
                  onClick={selectAllCollaborators}
                  className="text-primary hover:underline font-medium"
                >
                  Selecionar Todos
                </button>
                <span className="text-muted-foreground">•</span>
                <button
                  type="button"
                  onClick={deselectAllCollaborators}
                  className="text-muted-foreground hover:underline"
                >
                  Desmarcar Todos
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              {collaborators.map((colab) => (
                <label
                  key={colab.matricula}
                  className={`flex items-center justify-between p-2.5 rounded-lg border text-xs cursor-pointer transition-colors ${
                    colab.selecionado
                      ? "bg-primary/5 border-primary/40 shadow-2xs"
                      : "bg-background border-border hover:bg-muted/40 opacity-70"
                  }`}
                >
                  <div className="flex items-center gap-2.5 min-w-0">
                    <input
                      type="checkbox"
                      checked={colab.selecionado}
                      onChange={() => toggleCollaborator(colab.matricula)}
                      className="rounded border-input text-primary focus:ring-primary h-4 w-4"
                    />
                    <div className="truncate">
                      <p className="font-semibold text-foreground truncate">{colab.nome}</p>
                      <p className="text-[11px] text-muted-foreground font-mono">
                        Matrícula: <strong className="text-foreground">{colab.matricula}</strong>
                      </p>
                    </div>
                  </div>
                  {colab.selecionado && <Check className="h-4 w-4 text-primary shrink-0" />}
                </label>
              ))}
            </div>

            {/* Adicionar colaborador avulso */}
            <form
              onSubmit={handleAddCustomCollaborator}
              className="flex items-center gap-2 pt-1 text-xs"
            >
              <input
                type="text"
                value={newColabName}
                onChange={(e) => setNewColabName(e.target.value)}
                placeholder="Adicionar outro colaborador..."
                className="flex-1 bg-background border border-input rounded px-2.5 py-1 text-xs"
              />
              <input
                type="text"
                value={newColabMatricula}
                onChange={(e) => setNewColabMatricula(e.target.value)}
                placeholder="Matrícula (ex: 10059999)"
                className="w-40 bg-background border border-input rounded px-2.5 py-1 text-xs font-mono"
              />
              <button
                type="submit"
                disabled={!newColabName.trim() || !newColabMatricula.trim()}
                className="btn-secondary py-1 px-2.5 text-xs inline-flex items-center gap-1 disabled:opacity-40"
              >
                <Plus className="h-3 w-3" />
                <span>Adicionar</span>
              </button>
            </form>
          </div>

          {/* Atividades e Horários */}
          <div className="space-y-2">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Clock className="h-4 w-4 text-primary" />
                <h4 className="text-xs font-bold uppercase tracking-wider text-foreground">
                  Etapas e Intervalos Operacionais ({intervals.length} etapas)
                </h4>
              </div>
              <span className="text-[11px] text-muted-foreground">
                {events.length > 0
                  ? "Calculado a partir dos eventos da linha do tempo"
                  : "Sequência padrão do modelo SAP (6 etapas)"}
              </span>
            </div>

            <div className="border border-border rounded-lg overflow-hidden">
              <div className="max-h-44 overflow-y-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-muted/80 text-muted-foreground font-semibold sticky top-0 border-b border-border">
                    <tr>
                      <th className="py-2 px-3">Atividade</th>
                      <th className="py-2 px-3">Início</th>
                      <th className="py-2 px-3">Fim</th>
                      <th className="py-2 px-3">Desvio</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border font-mono text-[11px]">
                    {intervals.map((intv, idx) => (
                      <tr key={intv.id || idx} className="hover:bg-muted/20">
                        <td className="py-1.5 px-3 font-sans font-medium text-foreground">
                          {intv.atividade}
                        </td>
                        <td className="py-1.5 px-3 text-muted-foreground">
                          {intv.dataInicio} {intv.horaInicio}
                        </td>
                        <td className="py-1.5 px-3 text-muted-foreground">
                          {intv.dataFim} {intv.horaFim}
                        </td>
                        <td className="py-1.5 px-3 text-amber-600 dark:text-amber-400 font-sans">
                          {intv.eventoDesvio || "-"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          {/* Resumo da Exportação */}
          <div className="p-3 bg-emerald-500/10 border border-emerald-500/20 rounded-lg flex items-start gap-3">
            <FileCheck className="h-5 w-5 text-emerald-600 dark:text-emerald-400 shrink-0 mt-0.5" />
            <div className="text-xs space-y-1">
              <p className="font-semibold text-foreground">
                Total de {totalRows} linhas serão geradas no arquivo {expectedFilename}
              </p>
              <p className="text-muted-foreground">
                A planilha segue as colunas e formato da transação <strong>IW41</strong>:{" "}
                {formatSapDate(header.data)}, Ordem, OP, Matricula, DATA INICIO, HORA INICIO, DATA
                FIM, HORA FIM, Evento Desvio e Confirmação Final.
              </p>
              <p className="text-[11px] text-emerald-700 dark:text-emerald-300 font-medium">
                ✓ Coluna "Confirmação Final ?" conterá "x" na última linha e valor booleano
                configurado na coluna Z para a automação SAP.
              </p>
            </div>
          </div>
        </div>

        {/* Footer Actions */}
        <div className="flex items-center justify-between px-5 py-4 border-t border-border bg-muted/20">
          <button
            type="button"
            onClick={onClose}
            disabled={isExporting}
            className="btn-ghost text-xs py-2 px-3 text-muted-foreground hover:text-foreground disabled:opacity-50"
          >
            Cancelar
          </button>

          <button
            type="button"
            onClick={handleExport}
            disabled={selectedCollaborators.length === 0 || intervals.length === 0 || isExporting}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-medium py-2 px-4 rounded-md transition-colors flex items-center gap-2 shadow-xs disabled:opacity-50 text-xs sm:text-sm"
          >
            {isExporting ? (
              <>
                <Loader2 className="h-4 w-4 animate-spin" />
                <span>Gerando .xlsm com Macro...</span>
              </>
            ) : (
              <>
                <FileSpreadsheet className="h-4 w-4" />
                <span>Gerar Planilha IW41 ({expectedFilename})</span>
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
}
