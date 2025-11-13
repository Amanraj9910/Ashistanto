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

    // ⭐ Store session ID in a cookie (frontend can read it)
    res.cookie("sessionId", sessionId, {
      httpOnly: false,   // Allow frontend to read it
      secure: true,      // Required for HTTPS
      sameSite: "Lax"
    });

    // ⭐ Redirect user back to homepage
    return res.redirect("https://microsoft-agent-aubbhefsbzagdhha.eastus-01.azurewebsites.net/");

  } catch (err) {
    console.error('❌ Login failed:', err);
    res.status(500).send('Login failed.');
  }
});

// Step 3: Get logged in user
router.get('/user', (req, res) => {
  if (!loggedInUser) return res.status(401).send('User not logged in');
  res.json(loggedInUser);
});

// Step 4: Get token by session ID (frontend will call this)
router.get('/session-token/:sessionId', (req, res) => {
  const { sessionId } = req.params;
  const token = userTokenStore.get(sessionId);

  if (!token) {
    return res.status(404).json({ error: 'Session not found or token expired' });
  }

  res.json({ accessToken: token });
});

module.exports = { router, loggedInUser, userTokenStore };
