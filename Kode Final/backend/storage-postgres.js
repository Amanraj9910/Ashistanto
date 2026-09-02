const { Pool } = require('pg');

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_SSL === 'false' ? false : { rejectUnauthorized: false },
  max: Number(process.env.DATABASE_POOL_MAX || 10),
  idleTimeoutMillis: 30000,
  connectionTimeoutMillis: 10000
});

const query = (text, values) => pool.query(text, values);
const json = (value, fallback) => value == null ? fallback : (typeof value === 'string' ? JSON.parse(value) : value);

const userTokenStore = {
  async set(sessionId, data) { await query(`INSERT INTO user_tokens(session_id,access_token,account,expires_at,email) VALUES($1,$2,$3,$4,$5) ON CONFLICT(session_id) DO UPDATE SET access_token=EXCLUDED.access_token,account=EXCLUDED.account,expires_at=EXCLUDED.expires_at,email=EXCLUDED.email,updated_at=NOW()`, [sessionId, data.accessToken, data.account || null, data.expiresAt || null, data.email || null]); },
  async get(sessionId) { const { rows } = await query('SELECT * FROM user_tokens WHERE session_id=$1', [sessionId]); const row = rows[0]; return row ? { accessToken: row.access_token, account: row.account, expiresAt: row.expires_at, email: row.email } : null; },
  async has(sessionId) { const { rows } = await query('SELECT 1 FROM user_tokens WHERE session_id=$1', [sessionId]); return Boolean(rows[0]); },
  async delete(sessionId) { await query('DELETE FROM user_tokens WHERE session_id=$1', [sessionId]); }
};

const pendingActionsStore = {
  async set(actionId, data) { await query(`INSERT INTO pending_actions(action_id,data,status,timestamp) VALUES($1,$2,$3,$4) ON CONFLICT(action_id) DO UPDATE SET data=EXCLUDED.data,status=EXCLUDED.status,timestamp=EXCLUDED.timestamp`, [actionId, data, data.status || 'pending', data.timestamp || new Date()]); },
  async get(actionId) { const { rows } = await query('SELECT data FROM pending_actions WHERE action_id=$1', [actionId]); return rows[0] ? json(rows[0].data, {}) : null; },
  async delete(actionId) { await query('DELETE FROM pending_actions WHERE action_id=$1', [actionId]); },
  async has(actionId) { const { rows } = await query('SELECT 1 FROM pending_actions WHERE action_id=$1', [actionId]); return Boolean(rows[0]); },
  async entries() { const { rows } = await query('SELECT action_id,data FROM pending_actions'); return rows.map(row => [row.action_id, json(row.data, {})]); },
  async values() { const { rows } = await query('SELECT data FROM pending_actions'); return rows.map(row => json(row.data, {})); }
};

const conversationSessions = {
  async set(sessionId, history, summary = '') { await query(`INSERT INTO conversation_sessions(session_id,history,context_summary) VALUES($1,$2,$3) ON CONFLICT(session_id) DO UPDATE SET history=EXCLUDED.history,context_summary=EXCLUDED.context_summary,updated_at=NOW()`, [sessionId, history || [], summary || '']); },
  async get(sessionId) { const { rows } = await query('SELECT history,context_summary FROM conversation_sessions WHERE session_id=$1', [sessionId]); const row = rows[0]; return row ? { history: json(row.history, []), summary: row.context_summary || '' } : null; },
  async has(sessionId) { const { rows } = await query('SELECT 1 FROM conversation_sessions WHERE session_id=$1', [sessionId]); return Boolean(rows[0]); },
  async delete(sessionId) { await query('DELETE FROM conversation_sessions WHERE session_id=$1', [sessionId]); },
  async entries() { const { rows } = await query('SELECT session_id,history,context_summary FROM conversation_sessions'); return rows.map(row => [row.session_id, { history: json(row.history, []), summary: row.context_summary || '' }]); }
};

const userPreferences = {
  async getAll(sessionId) { const { rows } = await query('SELECT preference_key, preference_value FROM user_preferences WHERE session_id=$1 ORDER BY preference_key', [sessionId]); return Object.fromEntries(rows.map(row => [row.preference_key, json(row.preference_value, null)])); },
  async set(sessionId, key, value) { await query(`INSERT INTO user_preferences(session_id,preference_key,preference_value) VALUES($1,$2,$3) ON CONFLICT(session_id,preference_key) DO UPDATE SET preference_value=EXCLUDED.preference_value,updated_at=NOW()`, [sessionId, key, value]); },
  async delete(sessionId, key) { await query('DELETE FROM user_preferences WHERE session_id=$1 AND preference_key=$2', [sessionId, key]); }
};

const contextMemories = {
  async list(sessionId) { const { rows } = await query('SELECT memory_key,memory_value,memory_type,confidence FROM context_memories WHERE session_id=$1 ORDER BY updated_at DESC', [sessionId]); return rows.map(row => ({ key: row.memory_key, value: json(row.memory_value, null), type: row.memory_type, confidence: row.confidence })); },
  async upsert(sessionId, key, value, type = 'fact', confidence = null) { await query(`INSERT INTO context_memories(session_id,memory_key,memory_value,memory_type,confidence,last_used_at) VALUES($1,$2,$3,$4,$5,NOW()) ON CONFLICT(session_id,memory_key) DO UPDATE SET memory_value=EXCLUDED.memory_value,memory_type=EXCLUDED.memory_type,confidence=EXCLUDED.confidence,last_used_at=NOW(),updated_at=NOW()`, [sessionId, key, value, type, confidence]); },
  async delete(sessionId, key) { await query('DELETE FROM context_memories WHERE session_id=$1 AND memory_key=$2', [sessionId, key]); }
};

const msalCachePlugin = {
  async beforeCacheAccess(ctx) { const { rows } = await query("SELECT cache_data FROM msal_cache WHERE id='default'"); if (rows[0]) ctx.tokenCache.deserialize(rows[0].cache_data); },
  async afterCacheAccess(ctx) { if (ctx.cacheHasChanged) await query("INSERT INTO msal_cache(id,cache_data) VALUES('default',$1) ON CONFLICT(id) DO UPDATE SET cache_data=EXCLUDED.cache_data,updated_at=NOW()", [ctx.tokenCache.serialize()]); }
};

async function close() { await pool.end(); }
module.exports = { userTokenStore, pendingActionsStore, conversationSessions, userPreferences, contextMemories, msalCachePlugin, pool, close };
