-- Long-term user context, kept separate from short-term chat history.
CREATE TABLE IF NOT EXISTS user_preferences (
  session_id TEXT NOT NULL,
  preference_key TEXT NOT NULL,
  preference_value JSONB NOT NULL DEFAULT '{}'::jsonb,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  PRIMARY KEY (session_id, preference_key)
);

CREATE TABLE IF NOT EXISTS context_memories (
  id BIGSERIAL PRIMARY KEY,
  session_id TEXT NOT NULL,
  memory_key TEXT NOT NULL,
  memory_value JSONB NOT NULL DEFAULT '{}'::jsonb,
  memory_type TEXT NOT NULL DEFAULT 'fact',
  confidence NUMERIC(4,3),
  last_used_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (session_id, memory_key)
);

CREATE INDEX IF NOT EXISTS context_memories_session_idx ON context_memories(session_id, updated_at DESC);
