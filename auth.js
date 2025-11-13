const express = require('express');
const router = express.Router();
const {
  getAuthUrl,
  getAccessTokenByAuthCode
} = require('./graph-tools');

let loggedInUser = null; 
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

    const sessionId = `session_${Date.now()}`;

    loggedInUser = {
      sessionId: sessionId,
      accessToken: tokens.accessToken,
      refreshToken: tokens.refreshToken,
      email: tokens.account.username
    };

    userTokenStore.set(sessionId, tokens.accessToken);

    console.log('✅ User logged in:', loggedInUser.email);
    console.log('📌 Session ID:', sessionId);

    // ⭐⭐⭐ FIX: store session ID in localStorage (never fails)
    return res.send(`
      <html>
        <body style="font-family: Arial; text-align:center; padding-top:40px;">
          <h2>Logging you in…</h2>
          <p>Please wait…</p>
          <script>
            // Save session ID
            localStorage.setItem('sessionId', '${sessionId}');
            
            // Redirect to homepage
            window.location.href = "https://microsoft-agent-aubbhefsbzagdhha.eastus-01.azurewebsites.net/";
          </script>
        </body>
      </html>
    `);

  } catch (err) {
    console.error('❌ Login failed:', err);
    return res.status(500).send('Login failed.');
  }
});

// Step 3: API – Get logged in user
router.get('/user', (req, res) => {
  if (!loggedInUser) return res.status(401).send('User not logged in');
  res.json(loggedInUser);
});

// Step 4: API – Get access token using session ID
router.get('/session-token/:sessionId', (req, res) => {
  const { sessionId } = req.params;
  const token = userTokenStore.get(sessionId);

  if (!token) {
    return res.status(404).json({ error: 'Session not found or token expired' });
  }

  res.json({ accessToken: token });
});

module.exports = { router, loggedInUser, userTokenStore };
