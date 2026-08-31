const express = require('express');
const router = express.Router();
const {
  getAuthUrl,
  getAccessTokenByAuthCode
} = require('./graph-tools');
const { userTokenStore } = require('./storage');

// Step 1: Redirect to Microsoft login
router.get('/login', async (req, res) => {
  const url = await getAuthUrl();
  res.redirect(url);
});

// Logout endpoint - clear session
router.post('/logout', async (req, res) => {
  const { sessionId } = req.body;

  if (sessionId && await userTokenStore.has(sessionId)) {
    await userTokenStore.delete(sessionId);
    console.log('✅ Session cleared:', sessionId);
  }

  res.json({ success: true, message: 'Logged out successfully' });
});

// Logout endpoint - GET version for frontend redirect
router.get('/logout', async (req, res) => {
  const sessionId = req.query.sessionId;
  
  if (sessionId && await userTokenStore.has(sessionId)) {
    await userTokenStore.delete(sessionId);
    console.log('✅ Session cleared (GET):', sessionId);
  }
  
  // Redirect to login page
  res.redirect('/login/');
});

// Alternative route for /redirect (alias for /login)
router.get('/redirect', async (req, res) => {
  const url = await getAuthUrl();
  res.redirect(url);
});

// Step 2: Handle Microsoft redirect
router.get('/callback', async (req, res) => {
  try {
    const code = req.query.code;
    const tokens = await getAccessTokenByAuthCode(code);

    // Generate a session ID for this user 
    const sessionId = `session_${Date.now()}`;

    // Store tokens with expiration metadata and MSAL account for automatic silent refresh
    // Note: MSAL manages refresh tokens internally in its cache - we don't need to store them
    // We store the `account` object which is required by acquireTokenSilent()
    const email = tokens.account.username;
    await userTokenStore.set(sessionId, {
      accessToken: tokens.accessToken,
      account: tokens.account,  // CRITICAL: needed for acquireTokenSilent
      expiresAt: Date.now() + ((tokens.expiresIn || 3600) * 1000), // Default 1 hour if not provided
      email: email
    });

    console.log('✅ User logged in:', email);
    console.log('📌 Session ID:', sessionId);
    console.log('🔑 MSAL account stored for silent refresh:', !!tokens.account);
    console.log('⏰ Token expires at:', new Date(Date.now() + ((tokens.expiresIn || 3600) * 1000)).toISOString());

    res.redirect(`/auth/success?sessionId=${sessionId}`);
  } catch (err) {
    console.error('❌ Login failed:', err);
    res.status(500).send('Login failed.');
  }
});

// Step 3: Confirmation page with auto-redirect
// NOTE: /auth/success is intentionally NOT handled here.
// It is served by the Next.js static export at frontend/out/auth/success/index.html, which
// reads ?sessionId= client-side and stores it. The /callback above redirects there.
// Do not re-add a route for it: express.static is mounted before this router in server.js,
// so a duplicate would be silently shadowed and become confusing dead code.

router.get('/user', async (req, res) => {
  const sessionId = req.query.sessionId;
  const sessionData = sessionId ? await userTokenStore.get(sessionId) : null;
  if (!sessionData) return res.status(401).send('User not logged in');
  res.json({
    sessionId: sessionId,
    accessToken: sessionData.accessToken,
    email: sessionData.email
  });
});

router.get('/session-token/:sessionId', async (req, res) => {
  const { sessionId } = req.params;
  const tokenData = await userTokenStore.get(sessionId);

  if (!tokenData) {
    return res.status(404).json({ error: 'Session not found or token expired' });
  }

  // Return access token from the token object
  res.json({ accessToken: tokenData.accessToken });
});

module.exports = { router };