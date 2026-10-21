
-- Restrict shared_state to a whitelist of known keys with bounded value length,
-- replacing overly-permissive USING/WITH CHECK (true) policies.

ALTER TABLE public.shared_state
  ADD CONSTRAINT shared_state_key_whitelist
  CHECK (key IN ('last_defect'));

ALTER TABLE public.shared_state
  ADD CONSTRAINT shared_state_value_length
  CHECK (value IS NULL OR char_length(value) <= 100);

DROP POLICY IF EXISTS "Anyone can insert shared_state" ON public.shared_state;
DROP POLICY IF EXISTS "Anyone can update shared_state" ON public.shared_state;
DROP POLICY IF EXISTS "Anyone can read shared_state" ON public.shared_state;

CREATE POLICY "Read allowed shared_state keys"
  ON public.shared_state FOR SELECT
  USING (key IN ('last_defect'));

CREATE POLICY "Insert allowed shared_state keys"
  ON public.shared_state FOR INSERT
  WITH CHECK (
    key IN ('last_defect')
    AND (value IS NULL OR char_length(value) <= 100)
  );

CREATE POLICY "Update allowed shared_state keys"
  ON public.shared_state FOR UPDATE
  USING (key IN ('last_defect'))
  WITH CHECK (
    key IN ('last_defect')
    AND (value IS NULL OR char_length(value) <= 100)
  );
