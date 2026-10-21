import { supabase } from "@/integrations/supabase/client";
import {
  type SavedCloudReport,
  getLocalCache as getLocalReportsCache,
  setLocalCache as setLocalReportsCache,
  mergeReports,
} from "./cloud-reports";
import { type EquipmentChecklist, getLocalChecklists, saveLocalChecklist } from "./checklist";

export const SUPABASE_PROJECT_ID = "ivzutgehaxmvoreypssm";
export const SUPABASE_SQL_EDITOR_URL = `https://supabase.com/dashboard/project/${SUPABASE_PROJECT_ID}/sql/new`;

export const SQL_MIGRATION_SCRIPT = `-- ==========================================
-- SCRIPT DE MIGRAÇÃO: RELATÓRIOS E CHECKLISTS
-- Execute este script no SQL Editor do Supabase:
-- https://supabase.com/dashboard/project/${SUPABASE_PROJECT_ID}/sql/new
-- ==========================================

-- 1. Criação da tabela de Relatórios
CREATE TABLE IF NOT EXISTS public.relatorios (
  id TEXT PRIMARY KEY,
  data TEXT NOT NULL,
  nome TEXT NOT NULL DEFAULT '',
  local TEXT DEFAULT '',
  sinais TEXT DEFAULT '',
  linha TEXT DEFAULT '',
  via TEXT DEFAULT '',
  gda TEXT DEFAULT '',
  om TEXT DEFAULT '',
  empresa TEXT DEFAULT '',
  atividade TEXT DEFAULT '',
  colaboradores TEXT DEFAULT '',
  descritivo TEXT DEFAULT '',
  "kmProgramado" TEXT DEFAULT '',
  executado TEXT DEFAULT '',
  "observacaoGeral" TEXT DEFAULT '',
  telefone TEXT DEFAULT '',
  events JSONB DEFAULT '[]'::jsonb,
  photos JSONB DEFAULT '[]'::jsonb,
  totals JSONB DEFAULT '{}'::jsonb,
  "createdAt" BIGINT NOT NULL,
  "updatedAt" BIGINT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- 2. Criação da tabela de Checklists de Equipamentos
CREATE TABLE IF NOT EXISTS public.checklists_equipamento (
  id TEXT PRIMARY KEY,
  data TEXT NOT NULL,
  hora TEXT DEFAULT '',
  equipamento TEXT NOT NULL DEFAULT '',
  prefixo TEXT DEFAULT '',
  "horimetroMotor" TEXT DEFAULT '',
  "horimetroImplementoFerroviario" TEXT DEFAULT '',
  "horimetroGerador" TEXT DEFAULT '',
  "quilometragemCaminhao" TEXT DEFAULT '',
  "horimetroOuKm" TEXT DEFAULT '',
  operador TEXT NOT NULL DEFAULT '',
  turno TEXT DEFAULT '',
  local TEXT DEFAULT '',
  "statusGeral" TEXT DEFAULT 'apto',
  "observacoesGerais" TEXT DEFAULT '',
  items JSONB DEFAULT '[]'::jsonb,
  "fotosCount" INTEGER DEFAULT 0,
  "createdAt" BIGINT NOT NULL,
  "updatedAt" BIGINT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- 3. Habilita RLS (Row Level Security)
ALTER TABLE public.relatorios ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.checklists_equipamento ENABLE ROW LEVEL SECURITY;

-- 4. Cria Políticas de Acesso Público para leitura e escrita
DROP POLICY IF EXISTS "Allow all on relatorios" ON public.relatorios;
CREATE POLICY "Allow all on relatorios" ON public.relatorios
  FOR ALL
  TO public
  USING (true)
  WITH CHECK (true);

DROP POLICY IF EXISTS "Allow all on checklists_equipamento" ON public.checklists_equipamento;
CREATE POLICY "Allow all on checklists_equipamento" ON public.checklists_equipamento
  FOR ALL
  TO public
  USING (true)
  WITH CHECK (true);

-- 5. Concede permissões aos papéis
GRANT ALL ON public.relatorios TO anon, authenticated, service_role;
GRANT ALL ON public.checklists_equipamento TO anon, authenticated, service_role;

-- 6. Habilita Replicação em Tempo Real (Realtime)
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables 
    WHERE pubname = 'supabase_realtime' AND tablename = 'relatorios'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE public.relatorios;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables 
    WHERE pubname = 'supabase_realtime' AND tablename = 'checklists_equipamento'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE public.checklists_equipamento;
  END IF;
END $$;
`;

export const SUPABASE_MIGRATION_SQL = SQL_MIGRATION_SCRIPT;

export interface MigrationProgress {
  stage: string;
  current: number;
  total: number;
  message?: string;
  phase?: "in_progress" | "completed" | "error";
  reportsMigrated?: number;
  checklistsMigrated?: number;
}

export interface MigrationStatus {
  tablesReady: boolean;
  relatoriosTableReady: boolean;
  checklistsTableReady: boolean;
  reportsCountInSupabase: number;
  checklistsCountInSupabase: number;
  lastChecked: number;
  error?: string | null;
}

let cachedStatus: MigrationStatus = {
  tablesReady: false,
  relatoriosTableReady: false,
  checklistsTableReady: false,
  reportsCountInSupabase: 0,
  checklistsCountInSupabase: 0,
  lastChecked: 0,
  error: null,
};

const statusListeners = new Set<(status: MigrationStatus) => void>();

export function subscribeMigrationStatus(cb: (status: MigrationStatus) => void): () => void {
  statusListeners.add(cb);
  cb(cachedStatus);
  return () => {
    statusListeners.delete(cb);
  };
}

function notifyStatus(status: MigrationStatus) {
  cachedStatus = status;
  statusListeners.forEach((fn) => {
    try {
      fn(status);
    } catch (e) {
      console.warn("Migration listener err:", e);
    }
  });
}

/**
 * Verifica se as tabelas relatorios e checklists_equipamento existem no Supabase
 */
export async function checkSupabaseTables(): Promise<MigrationStatus> {
  let relatoriosReady = false;
  let checklistsReady = false;
  let reportsCount = 0;
  let checklistsCount = 0;
  let errMsg: string | null = null;

  try {
    const { count, error } = await supabase
      .from("relatorios")
      .select("id", { count: "exact", head: true });

    if (!error) {
      relatoriosReady = true;
      reportsCount = count || 0;
    } else {
      errMsg = error.message;
    }
  } catch (err) {
    errMsg = err instanceof Error ? err.message : String(err);
  }

  try {
    const { count, error } = await supabase
      .from("checklists_equipamento")
      .select("id", { count: "exact", head: true });

    if (!error) {
      checklistsReady = true;
      checklistsCount = count || 0;
    } else if (!errMsg) {
      errMsg = error.message;
    }
  } catch (err) {
    if (!errMsg) errMsg = err instanceof Error ? err.message : String(err);
  }

  const status: MigrationStatus = {
    tablesReady: relatoriosReady && checklistsReady,
    relatoriosTableReady: relatoriosReady,
    checklistsTableReady: checklistsReady,
    reportsCountInSupabase: reportsCount,
    checklistsCountInSupabase: checklistsCount,
    lastChecked: Date.now(),
    error: errMsg,
  };

  notifyStatus(status);
  return status;
}

/**
 * Converte um SavedCloudReport para o formato de inserção na tabela relatorios do Supabase
 */
export function formatReportForSupabase(report: SavedCloudReport) {
  return {
    id: report.id,
    data: report.data || new Date().toISOString().split("T")[0],
    nome: report.nome || "",
    local: report.local || "",
    sinais: report.sinais || "",
    linha: report.linha || "",
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
    events: report.events || [],
    photos: report.photos || [],
    totals: report.totals || {
      jornada: 0,
      atividade: 0,
      preparacao: 0,
      parada: 0,
      producaoTotal: 0,
      defeitosTotal: 0,
    },
    createdAt: Number(report.createdAt || Date.now()),
    updatedAt: Number(report.updatedAt || Date.now()),
  };
}

/**
 * Converte um EquipmentChecklist para o formato de inserção na tabela checklists_equipamento
 */
export function formatChecklistForSupabase(chk: EquipmentChecklist) {
  return {
    id: chk.id,
    data: chk.data || new Date().toISOString().split("T")[0],
    hora: chk.hora || "",
    equipamento: chk.equipamento || "",
    prefixo: chk.prefixo || "",
    horimetroMotor: chk.horimetroMotor || "",
    horimetroImplementoFerroviario: chk.horimetroImplementoFerroviario || "",
    horimetroGerador: chk.horimetroGerador || "",
    quilometragemCaminhao: chk.quilometragemCaminhao || "",
    horimetroOuKm: chk.horimetroOuKm || "",
    operador: chk.operador || "",
    turno: chk.turno || "",
    local: chk.local || "",
    statusGeral: chk.statusGeral || "apto",
    observacoesGerais: chk.observacoesGerais || "",
    items: chk.items || [],
    fotosCount: chk.fotosCount || 0,
    createdAt: Number(chk.createdAt || Date.now()),
    updatedAt: Number(chk.updatedAt || Date.now()),
  };
}

/**
 * Executa a migração de todos os dados locais/servidor para o Supabase
 */
export async function migrateAllDataToSupabase(options?: {
  onProgress?: (progress: MigrationProgress) => void;
}): Promise<{
  success: boolean;
  migratedReports: number;
  migratedChecklists: number;
  errors: string[];
}> {
  const errors: string[] = [];
  let migratedReports = 0;
  let migratedChecklists = 0;

  // 1. Verifica se as tabelas existem
  const status = await checkSupabaseTables();
  if (!status.tablesReady) {
    return {
      success: false,
      migratedReports: 0,
      migratedChecklists: 0,
      errors: [
        "As tabelas 'relatorios' e/ou 'checklists_equipamento' ainda não foram criadas no Supabase. Por favor, execute o script SQL no SQL Editor do Supabase antes de migrar.",
      ],
    };
  }

  // 2. Coleta todos os relatórios disponíveis (local cache + servidor)
  let allReports: SavedCloudReport[] = getLocalReportsCache();
  try {
    const res = await fetch("/api/cloud-reports");
    if (res.ok) {
      const serverReports = (await res.json()) as SavedCloudReport[];
      if (Array.isArray(serverReports) && serverReports.length > 0) {
        allReports = mergeReports(serverReports, allReports);
      }
    }
  } catch (e) {
    console.warn("Could not fetch server reports during migration:", e);
  }

  // 3. Envia relatórios para o Supabase em lotes de 10
  if (allReports.length > 0) {
    options?.onProgress?.({
      stage: "Migrando relatórios diários...",
      current: 0,
      total: allReports.length,
    });

    for (let i = 0; i < allReports.length; i += 10) {
      const batch = allReports.slice(i, i + 10).map(formatReportForSupabase);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { error } = await supabase
        .from("relatorios")
        .upsert(batch as any, { onConflict: "id" });
      if (error) {
        errors.push(
          `Erro ao migrar lote de relatórios (${i}-${i + batch.length}): ${error.message}`,
        );
      } else {
        migratedReports += batch.length;
      }
      options?.onProgress?.({
        stage: "Migrando relatórios diários...",
        current: Math.min(i + 10, allReports.length),
        total: allReports.length,
      });
    }
  }

  // 4. Coleta todos os checklists de equipamento
  const localChecklists = getLocalChecklists();
  if (localChecklists.length > 0) {
    options?.onProgress?.({
      stage: "Migrando checklists de equipamento...",
      current: 0,
      total: localChecklists.length,
    });

    for (let i = 0; i < localChecklists.length; i += 10) {
      const batch = localChecklists.slice(i, i + 10).map(formatChecklistForSupabase);
      const { error } = await supabase
        .from("checklists_equipamento")
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .upsert(batch as any, { onConflict: "id" });
      if (error) {
        errors.push(
          `Erro ao migrar lote de checklists (${i}-${i + batch.length}): ${error.message}`,
        );
      } else {
        migratedChecklists += batch.length;
      }
      options?.onProgress?.({
        stage: "Migrando checklists de equipamento...",
        current: Math.min(i + 10, localChecklists.length),
        total: localChecklists.length,
      });
    }
  }

  // Atualiza status final
  await checkSupabaseTables();

  options?.onProgress?.({
    stage: "Concluído",
    current: 100,
    total: 100,
    message: "Migração concluída com sucesso!",
    phase: "completed",
    reportsMigrated: migratedReports,
    checklistsMigrated: migratedChecklists,
  });

  return {
    success: errors.length === 0,
    migratedReports,
    migratedChecklists,
    errors,
  };
}
