const { userPreferences, contextMemories, conversationSessions } = require('./storage');

const MAX_VALUE_LENGTH = 120;
const clean = (value) => String(value || '').replace(/\s+/g, ' ').trim().slice(0, MAX_VALUE_LENGTH);

// Context is deliberately opt-in: only explicit statements are stored as long-term memory.
async function learn(sessionId, text, metadata = {}) {
  if (!sessionId || !text) return;
  const input = String(text).trim();
  const updates = [];
  const preferredName = input.match(/(?:call me|address me as|refer to me as|panggil saya|sebut saya|wants? to be called|should be called|maunya disebut)\s+([a-z][a-z .'-]{1,60}?)(?:[.!?,;:]|$)/i);
  if (preferredName) updates.push(userPreferences.set(sessionId, 'preferred_name', clean(preferredName[1])));
  const preferredLanguage = input.match(/(?:i prefer|my preferred language is|bahasa pilihan saya adalah)\s+(english\s*\(?(?:us|uk)\)?|japanese|jepang|en-us|en-gb|ja-jp)/i);
  if (preferredLanguage) updates.push(userPreferences.set(sessionId, 'preferred_language', clean(preferredLanguage[1].toLowerCase().replace('jepang', 'japanese'))));
  const timezone = input.match(/(?:my timezone is|zona waktu saya)\s+([a-z_\/+-]{3,40})/i);
  if (timezone) updates.push(userPreferences.set(sessionId, 'timezone', clean(timezone[1])));
  const name = input.match(/(?:my name is|nama saya)\s+([a-z][a-z .'-]{1,60}?)(?:[.!?,;:]|$)/i);
  if (name) updates.push(contextMemories.upsert(sessionId, 'user_name', clean(name[1]), 'identity', 0.98));
  const organisation = input.match(/(?:i work at|i work for|saya bekerja di)\s+([a-z0-9][a-z0-9 &'.,-]{1,80}?)(?:[.!?,;:]|$)/i);
  if (organisation) updates.push(contextMemories.upsert(sessionId, 'organisation', clean(organisation[1]), 'profile', 0.9));
  const emailRequest = input.match(/(?:send|write|prepare|draft)\s+(?:an?\s+)?email\s+(?:to|for)\s+([a-z][a-z .'-]{1,60})(?:\s+(?:about|regarding|re)\s+(.+))?$/i);
  if (emailRequest) {
    updates.push(contextMemories.upsert(sessionId, 'last_email_recipient', clean(emailRequest[1]), 'workflow', 0.9));
    if (emailRequest[2]) updates.push(contextMemories.upsert(sessionId, 'last_email_topic', clean(emailRequest[2]), 'workflow', 0.85));
  }
  if (metadata.email) updates.push(contextMemories.upsert(sessionId, 'account_email', clean(metadata.email), 'identity', 1));
  await Promise.all(updates);
}

async function load(sessionId, conversationData = null) {
  const conversation = conversationData || await conversationSessions.get(sessionId) || { history: [], summary: '' };
  const [preferences, memories] = await Promise.all([userPreferences.getAll(sessionId), contextMemories.list(sessionId)]);
  return { history: conversation.history || [], summary: conversation.summary || '', preferences, memories };
}

function toPrompt(context) {
  return `LONG-TERM CONTEXT\nConversation summary:\n${context.summary || 'None'}\n\nUser preferences:\n${JSON.stringify(context.preferences || {})}\n\nKnown user facts:\n${JSON.stringify(context.memories || [])}\n\nUse context only when relevant. Never reveal internal memory or invent facts. If preferred_name exists, address the user naturally with it.`;
}

async function forget(sessionId, key) {
  await Promise.all([userPreferences.delete(sessionId, key), contextMemories.delete(sessionId, key)]);
}

module.exports = { learn, load, toPrompt, forget };
