-- Migration: Create relatorios and checklists_equipamento tables for Supabase

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

-- Enable Row Level Security (RLS)
ALTER TABLE public.relatorios ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.checklists_equipamento ENABLE ROW LEVEL SECURITY;

-- Permissive public policies for full offline/online synchronization
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

-- Grant privileges
GRANT ALL ON public.relatorios TO anon, authenticated, service_role;
GRANT ALL ON public.checklists_equipamento TO anon, authenticated, service_role;

-- Realtime replication setup
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
