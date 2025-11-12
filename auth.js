const express = require('express');
const router = express.Router();
const {
  getAuthUrl,
  getAccessTokenByAuthCode
} = require('./graph-tools');

let loggedInUser = null; // store user session in memory
// Export a map to store user tokens by session ID
const userTokenStore = new Map();

// Step 1: Redirect to Microsoft login
router.get('/login', async (req, res) => {
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
    
    loggedInUser = {
      sessionId: sessionId,
      accessToken: tokens.accessToken,
      refreshToken: tokens.refreshToken,
      email: tokens.account.username
    };
    
    // Store token for later use in voice processing
    userTokenStore.set(sessionId, tokens.accessToken);
    
    console.log('✅ User logged in:', loggedInUser.email);
    console.log('📌 Session ID:', sessionId);
    
    res.redirect(`/auth/success?sessionId=${sessionId}`);
  } catch (err) {
    console.error('❌ Login failed:', err);
    res.status(500).send('Login failed.');
  }
});

// Step 3: Confirmation page
router.get('/success', (req, res) => {
  const sessionId = req.query.sessionId;
  if (!loggedInUser) return res.send('No active session');
  res.send(`
    <h2>✅ Logged in as ${loggedInUser.email}</h2>
    <p>Session ID: <strong>${sessionId}</strong></p>
    <p>You can now use the Voice Bot. Close this tab and return to the app.</p>
    <script>
      // Store session ID in localStorage for the web app
      localStorage.setItem('userSessionId', '${sessionId}');
    </script>
  `);
});

router.get('/user', (req, res) => {
  if (!loggedInUser) return res.status(401).send('User not logged in');
  res.json(loggedInUser);
});

router.get('/session-token/:sessionId', (req, res) => {
  const { sessionId } = req.params;
  const token = userTokenStore.get(sessionId);
  
  if (!token) {
    return res.status(404).json({ error: 'Session not found or token expired' });
  }
  
  res.json({ accessToken: token });
});

module.exports = { router, loggedInUser, userTokenStore };
