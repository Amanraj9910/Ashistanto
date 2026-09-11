CREATE TABLE IF NOT EXISTS user_tokens (
  session_id TEXT PRIMARY KEY,
  access_token TEXT,
  account JSONB,
  expires_at BIGINT,
  email TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS pending_actions (
  action_id TEXT PRIMARY KEY,
  data JSONB NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending',
  timestamp TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS conversation_sessions (
  session_id TEXT PRIMARY KEY,
  history JSONB NOT NULL DEFAULT '[]'::jsonb,
  context_summary TEXT NOT NULL DEFAULT '',
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS msal_cache (
  id TEXT PRIMARY KEY,
  cache_data TEXT NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS pending_actions_status_idx ON pending_actions(status);
CREATE INDEX IF NOT EXISTS conversation_sessions_updated_idx ON conversation_sessions(updated_at);
