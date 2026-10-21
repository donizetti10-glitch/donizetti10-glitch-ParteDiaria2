
CREATE TABLE public.shared_state (
  key TEXT PRIMARY KEY,
  value TEXT,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

GRANT SELECT, INSERT, UPDATE ON public.shared_state TO anon, authenticated;
GRANT ALL ON public.shared_state TO service_role;

ALTER TABLE public.shared_state ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone can read shared_state" ON public.shared_state FOR SELECT USING (true);
CREATE POLICY "Anyone can insert shared_state" ON public.shared_state FOR INSERT WITH CHECK (true);
CREATE POLICY "Anyone can update shared_state" ON public.shared_state FOR UPDATE USING (true) WITH CHECK (true);

INSERT INTO public.shared_state (key, value) VALUES ('last_defect', '') ON CONFLICT DO NOTHING;

ALTER PUBLICATION supabase_realtime ADD TABLE public.shared_state;
