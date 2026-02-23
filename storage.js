const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

const defaultDbPath = path.resolve(__dirname, 'data.sqlite');
const dbPath = process.env.DB_PATH || defaultDbPath;

console.log(`[Storage] Initializing SQLite database at: ${dbPath}`);

// Attempt to create the directory if it doesn't exist
const dbDir = path.dirname(dbPath);
try {
    if (!fs.existsSync(dbDir)) {
        console.log(`[Storage] Directory ${dbDir} does not exist. Attempting to create it...`);
        fs.mkdirSync(dbDir, { recursive: true });
    }
} catch (err) {
    console.error(`[Storage] ⚠️ Warning: Failed to create directory ${dbDir}. Permission denied?`);
}

const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error(`\n======================================================`);
        console.error(`❌ FATAL ERRROR: SQLite unable to open database file`);
        console.error(`Path: ${dbPath}`);
        console.error(`Error details: ${err.message}`);
        console.error(`\nPossible causes in Azure App Service:`);
        console.error(`1. If DB_PATH is set to /home/... but WEBSITES_ENABLE_APP_SERVICE_STORAGE is not true.`);
        console.error(`2. The Docker container runs as the 'node' user (from Dockerfile), which lacks write permissions to the mounted /home directory.`);
        console.error(`\nTo fix immediately: Remove the DB_PATH environment variable in Azure Configuration to use ephemeral local storage.`);
        console.error(`======================================================\n`);
    } else {
        console.log(`[Storage] ✅ SQLite database connected successfully.`);
    }
});

// Initialize database schema
db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS user_tokens (
        session_id TEXT PRIMARY KEY,
        access_token TEXT,
        account TEXT,
        expires_at INTEGER,
        email TEXT
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS pending_actions (
        action_id TEXT PRIMARY KEY,
        data TEXT,
        status TEXT,
        timestamp TEXT
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS conversation_sessions (
        session_id TEXT PRIMARY KEY,
        history TEXT
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS msal_cache (
        id TEXT PRIMARY KEY,
        cache_data TEXT
    )`);
});

// Helper for Promisifying SQLite queries
const runQuery = (sql, params = []) => new Promise((resolve, reject) => {
    db.run(sql, params, function (err) {
        if (err) reject(err);
        else resolve({ lastID: this.lastID, changes: this.changes });
    });
});

const getQuery = (sql, params = []) => new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => {
        if (err) reject(err);
        else resolve(row);
    });
});

const allQuery = (sql, params = []) => new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
    });
});

// --- User Token Store Wrapper ---
const userTokenStore = {
    async set(sessionId, data) {
        const { accessToken, account, expiresAt, email } = data;
        const accountStr = account ? JSON.stringify(account) : null;
        await runQuery(
            `INSERT OR REPLACE INTO user_tokens (session_id, access_token, account, expires_at, email) VALUES (?, ?, ?, ?, ?)`,
            [sessionId, accessToken, accountStr, expiresAt, email]
        );
    },
    async get(sessionId) {
        const row = await getQuery(`SELECT * FROM user_tokens WHERE session_id = ?`, [sessionId]);
        if (!row) return null;
        return {
            accessToken: row.access_token,
            account: row.account ? JSON.parse(row.account) : null,
            expiresAt: row.expires_at,
            email: row.email
        };
    },
    async has(sessionId) {
        const row = await getQuery(`SELECT 1 FROM user_tokens WHERE session_id = ?`, [sessionId]);
        return !!row;
    },
    async delete(sessionId) {
        await runQuery(`DELETE FROM user_tokens WHERE session_id = ?`, [sessionId]);
    }
};

// --- Pending Actions Store Wrapper ---
const pendingActionsStore = {
    async set(actionId, data) {
        const dataStr = JSON.stringify(data);
        const status = data.status || 'pending';
        const timestamp = data.timestamp || new Date().toISOString();
        await runQuery(
            `INSERT OR REPLACE INTO pending_actions (action_id, data, status, timestamp) VALUES (?, ?, ?, ?)`,
            [actionId, dataStr, status, timestamp]
        );
    },
    async get(actionId) {
        const row = await getQuery(`SELECT data FROM pending_actions WHERE action_id = ?`, [actionId]);
        if (!row) return null;
        return JSON.parse(row.data);
    },
    async delete(actionId) {
        await runQuery(`DELETE FROM pending_actions WHERE action_id = ?`, [actionId]);
    },
    async has(actionId) {
        const row = await getQuery(`SELECT 1 FROM pending_actions WHERE action_id = ?`, [actionId]);
        return !!row;
    },
    async entries() {
        const rows = await allQuery(`SELECT action_id, data FROM pending_actions`);
        return rows.map(r => [r.action_id, JSON.parse(r.data)]);
    },
    async values() {
        const rows = await allQuery(`SELECT data FROM pending_actions`);
        return rows.map(r => JSON.parse(r.data));
    }
};

// --- Conversation Sessions Wrapper ---
const conversationSessions = {
    async set(sessionId, historyArr) {
        const historyStr = JSON.stringify(historyArr);
        await runQuery(
            `INSERT OR REPLACE INTO conversation_sessions (session_id, history) VALUES (?, ?)`,
            [sessionId, historyStr]
        );
    },
    async get(sessionId) {
        const row = await getQuery(`SELECT history FROM conversation_sessions WHERE session_id = ?`, [sessionId]);
        if (!row) return undefined;
        return JSON.parse(row.history);
    },
    async has(sessionId) {
        const row = await getQuery(`SELECT 1 FROM conversation_sessions WHERE session_id = ?`, [sessionId]);
        return !!row;
    },
    async delete(sessionId) {
        await runQuery(`DELETE FROM conversation_sessions WHERE session_id = ?`, [sessionId]);
    },
    async entries() {
        const rows = await allQuery(`SELECT session_id, history FROM conversation_sessions`);
        return rows.map(r => [r.session_id, JSON.parse(r.history)]);
    }
};

// --- MSAL Cache Plugin ---
const msalCachePlugin = {
    async beforeCacheAccess(cacheContext) {
        const row = await getQuery(`SELECT cache_data FROM msal_cache WHERE id = 'default'`);
        if (row && row.cache_data) {
            cacheContext.tokenCache.deserialize(row.cache_data);
        }
    },
    async afterCacheAccess(cacheContext) {
        if (cacheContext.cacheHasChanged) {
            const cacheData = cacheContext.tokenCache.serialize();
            await runQuery(`INSERT OR REPLACE INTO msal_cache (id, cache_data) VALUES ('default', ?)`, [cacheData]);
        }
    }
};

module.exports = {
    userTokenStore,
    pendingActionsStore,
    conversationSessions,
    msalCachePlugin
};
