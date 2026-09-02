/* Local smoke test. Requires no Azure credentials; uses MOCK_AI=true. */
const port = Number(process.env.PORT || 3000);
const base = `http://127.0.0.1:${port}`;
async function request(path, options) {
  return fetch(`${base}${path}`, options);
}

(async () => {
  try {
    const health = await request('/health');
    if (!health.ok) throw new Error(`/health returned ${health.status}`);
    const text = await request('/api/text-message', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ text: 'Prepare an email about the project schedule', sessionId: 'smoke-session', accent: 'american' }) });
    if (!text.ok) throw new Error(`/api/text-message returned ${text.status}`);
    const body = await text.json();
    if (!body.response) throw new Error('Text response is missing');
    console.log('Backend smoke test passed');
  } catch (error) {
    throw error;
  }
})().catch(error => { console.error(error.message); process.exitCode = 1; });
