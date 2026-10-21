import { useState, useEffect } from "react";
import {
  Database,
  CheckCircle2,
  AlertCircle,
  Copy,
  Check,
  RefreshCw,
  ExternalLink,
  X,
  ArrowRight,
  ShieldCheck,
  Zap,
} from "lucide-react";
import {
  checkSupabaseTables,
  migrateAllDataToSupabase,
  type MigrationProgress,
  type MigrationStatus,
  SUPABASE_MIGRATION_SQL,
} from "@/lib/supabase-migration";

interface Props {
  isOpen: boolean;
  onClose: () => void;
  onMigrationComplete?: () => void;
}

export function SupabaseMigrationModal({ isOpen, onClose, onMigrationComplete }: Props) {
  const [isChecking, setIsChecking] = useState(false);
  const [isMigrating, setIsMigrating] = useState(false);
  const [copiedSql, setCopiedSql] = useState(false);
  const [tableStatus, setTableStatus] = useState<MigrationStatus | null>(null);
  const [progress, setProgress] = useState<MigrationProgress | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const checkStatus = async () => {
    setIsChecking(true);
    setErrorMsg(null);
    try {
      const status = await checkSupabaseTables();
      setTableStatus(status);
    } catch (e: unknown) {
      console.error("Error checking Supabase tables:", e);
      setErrorMsg("Erro ao verificar tabelas do Supabase. Verifique sua conexão.");
    } finally {
      setIsChecking(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      checkStatus();
    }
  }, [isOpen]);

  if (!isOpen) return null;

  const handleCopySql = () => {
    navigator.clipboard.writeText(SUPABASE_MIGRATION_SQL);
    setCopiedSql(true);
    setTimeout(() => setCopiedSql(false), 2500);
  };

  const handleStartMigration = async () => {
    setIsMigrating(true);
    setErrorMsg(null);
    try {
      const result = await migrateAllDataToSupabase({
        onProgress: (p) => {
          setProgress(p);
        },
      });
      if (result.success) {
        onMigrationComplete?.();
      } else if (result.errors?.length) {
        setErrorMsg(result.errors.join(", "));
      }
    } catch (e: unknown) {
      console.error("Migration error:", e);
      const msg =
        e instanceof Error ? e.message : "Ocorreu um erro durante a migração para o Supabase.";
      setErrorMsg(msg);
    } finally {
      setIsMigrating(false);
    }
  };

  const supabaseSqlEditorUrl =
    "https://supabase.com/dashboard/project/ivzutgehaxmvoreypssm/sql/new";

  return (
    <div
      id="supabase-migration-modal-backdrop"
      className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200"
    >
      <div
        id="supabase-migration-modal-card"
        className="relative w-full max-w-2xl bg-card border border-border text-card-foreground rounded-2xl shadow-2xl flex flex-col max-h-[90vh] overflow-hidden"
      >
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-border bg-muted/30">
          <div className="flex items-center gap-2.5">
            <div className="p-2 rounded-xl bg-emerald-500/10 text-emerald-600 dark:text-emerald-400">
              <Database className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-semibold text-foreground flex items-center gap-2">
                Migração para Banco de Dados Supabase
                <span className="text-[11px] font-normal px-2 py-0.5 rounded-full bg-emerald-500/10 text-emerald-700 dark:text-emerald-300 border border-emerald-500/20">
                  Sem limite de cota
                </span>
              </h2>
              <p className="text-xs text-muted-foreground">
                Transfira todos os relatórios e checklists do Firebase para o Supabase
              </p>
            </div>
          </div>
          <button
            id="close-supabase-modal-btn"
            onClick={onClose}
            className="p-1.5 text-muted-foreground hover:text-foreground rounded-lg hover:bg-muted transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Body */}
        <div className="px-6 py-5 overflow-y-auto space-y-5 text-xs text-foreground">
          {/* Status Banner */}
          <div className="p-3.5 rounded-xl border bg-muted/40 flex items-start gap-3">
            <ShieldCheck className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <div className="font-semibold text-foreground flex items-center gap-2">
                Projeto Supabase Conectado
                <code className="text-[10px] px-1.5 py-0.5 rounded bg-muted border font-mono">
                  ivzutgehaxmvoreypssm
                </code>
              </div>
              <p className="text-muted-foreground leading-relaxed">
                Ao migrar para o Supabase, seu aplicativo ganha sincronização instantânea em tempo
                real, sem bloqueios de cota diária de leitura do Firebase, funcionando perfeitamente
                em todos os seus dispositivos.
              </p>
            </div>
          </div>

          {/* Table Readiness */}
          <div className="border border-border rounded-xl p-4 space-y-3 bg-card">
            <div className="flex items-center justify-between">
              <span className="font-semibold text-foreground text-xs uppercase tracking-wider">
                1. Status das Tabelas no Supabase
              </span>
              <button
                id="check-supabase-tables-btn"
                onClick={checkStatus}
                disabled={isChecking}
                className="flex items-center gap-1.5 text-xs text-primary hover:underline disabled:opacity-50"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isChecking ? "animate-spin" : ""}`} />
                {isChecking ? "Verificando..." : "Verificar Tabelas"}
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 pt-1">
              <div
                className={`p-3 rounded-lg border flex items-center justify-between ${
                  tableStatus?.relatoriosTableReady
                    ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-800 dark:text-emerald-300"
                    : "bg-amber-500/10 border-amber-500/30 text-amber-800 dark:text-amber-300"
                }`}
              >
                <span className="font-medium">Tabela &apos;relatorios&apos;</span>
                {tableStatus?.relatoriosTableReady ? (
                  <span className="flex items-center gap-1 text-[11px] font-semibold text-emerald-600">
                    <CheckCircle2 className="w-4 h-4" /> Pronta
                  </span>
                ) : (
                  <span className="flex items-center gap-1 text-[11px] font-semibold text-amber-600">
                    <AlertCircle className="w-4 h-4" /> Pendente
                  </span>
                )}
              </div>

              <div
                className={`p-3 rounded-lg border flex items-center justify-between ${
                  tableStatus?.checklistsTableReady
                    ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-800 dark:text-emerald-300"
                    : "bg-amber-500/10 border-amber-500/30 text-amber-800 dark:text-amber-300"
                }`}
              >
                <span className="font-medium">Tabela &apos;checklists_equipamento&apos;</span>
                {tableStatus?.checklistsTableReady ? (
                  <span className="flex items-center gap-1 text-[11px] font-semibold text-emerald-600">
                    <CheckCircle2 className="w-4 h-4" /> Pronta
                  </span>
                ) : (
                  <span className="flex items-center gap-1 text-[11px] font-semibold text-amber-600">
                    <AlertCircle className="w-4 h-4" /> Pendente
                  </span>
                )}
              </div>
            </div>

            {/* SQL instruction if tables not ready */}
            {(!tableStatus?.tablesReady || !tableStatus) && (
              <div className="pt-2 border-t border-border mt-3 space-y-2">
                <p className="text-muted-foreground leading-relaxed">
                  Para criar as tabelas no Supabase com segurança, copie o script SQL abaixo e cole
                  no <strong>SQL Editor</strong> do seu painel do Supabase:
                </p>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    id="copy-supabase-sql-btn"
                    onClick={handleCopySql}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-primary text-primary-foreground text-xs font-medium hover:bg-primary/90 transition-colors shadow-xs"
                  >
                    {copiedSql ? (
                      <>
                        <Check className="w-3.5 h-3.5" /> SQL Copiado!
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" /> Copiar Código SQL
                      </>
                    )}
                  </button>
                  <a
                    id="open-supabase-dashboard-link"
                    href={supabaseSqlEditorUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center gap-1 px-3 py-1.5 rounded-lg border border-border hover:bg-muted text-xs transition-colors"
                  >
                    Abrir SQL Editor no Supabase <ExternalLink className="w-3.5 h-3.5 ml-0.5" />
                  </a>
                </div>
              </div>
            )}
          </div>

          {/* Migration Action Section */}
          <div className="border border-border rounded-xl p-4 space-y-3 bg-card">
            <span className="font-semibold text-foreground text-xs uppercase tracking-wider block">
              2. Transferência de Dados
            </span>
            <p className="text-muted-foreground">
              Esta ação coletará todos os relatórios diários de obra e checklists armazenados no seu
              dispositivo e servidor, formatando-os e gravando-os em lote no Supabase.
            </p>

            {progress && (
              <div className="p-3.5 rounded-lg bg-muted/60 border border-border space-y-2">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-medium text-foreground">
                    {progress.message || progress.stage}
                  </span>
                  <span className="text-muted-foreground">
                    {progress.phase === "completed" ? "Concluído" : "Em andamento..."}
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-2 text-xs pt-1">
                  <div className="bg-card p-2 rounded border">
                    <span className="text-muted-foreground block text-[11px]">
                      Relatórios Migrados:
                    </span>
                    <span className="text-sm font-semibold text-primary">
                      {progress.reportsMigrated ?? progress.current}
                    </span>
                  </div>
                  <div className="bg-card p-2 rounded border">
                    <span className="text-muted-foreground block text-[11px]">
                      Checklists Migrados:
                    </span>
                    <span className="text-sm font-semibold text-primary">
                      {progress.checklistsMigrated ?? 0}
                    </span>
                  </div>
                </div>
              </div>
            )}

            {errorMsg && (
              <div className="p-3 rounded-lg bg-destructive/10 border border-destructive/30 text-destructive text-xs flex items-center gap-2">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>{errorMsg}</span>
              </div>
            )}

            <div className="pt-2 flex items-center justify-between gap-3">
              <button
                id="start-supabase-migration-btn"
                onClick={handleStartMigration}
                disabled={isMigrating}
                className="w-full sm:w-auto flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-medium text-xs shadow-xs transition-colors disabled:opacity-50"
              >
                {isMigrating ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin" />
                    Migrando Dados para Supabase...
                  </>
                ) : (
                  <>
                    <Zap className="w-4 h-4" />
                    Iniciar Migração dos Dados Agora
                    <ArrowRight className="w-4 h-4" />
                  </>
                )}
              </button>

              {progress?.phase === "completed" && (
                <span className="text-xs text-emerald-600 dark:text-emerald-400 font-medium flex items-center gap-1 shrink-0">
                  <CheckCircle2 className="w-4 h-4" /> Migração Finalizada!
                </span>
              )}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-3.5 border-t border-border bg-muted/20 flex items-center justify-between text-xs text-muted-foreground">
          <span>O aplicativo já continuará salvando automaticamente no Supabase.</span>
          <button
            id="close-modal-bottom-btn"
            onClick={onClose}
            className="px-4 py-1.5 rounded-lg border border-border hover:bg-muted text-foreground transition-colors"
          >
            Fechar
          </button>
        </div>
      </div>
    </div>
  );
}
