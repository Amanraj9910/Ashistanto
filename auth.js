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

// Step 3: Confirmation page with auto-redirect
router.get('/success', (req, res) => {
  const sessionId = req.query.sessionId;
  if (!loggedInUser) return res.send('No active session');
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Login Successful</title>
      <style>
        body {
          font-family: Arial, sans-serif;
          display: flex;
          justify-content: center;
          align-items: center;
          height: 100vh;
          margin: 0;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
        }
        .container {
          text-align: center;
          background: rgba(255, 255, 255, 0.1);
          padding: 40px;
          border-radius: 15px;
          backdrop-filter: blur(10px);
        }
        .spinner {
          border: 4px solid rgba(255, 255, 255, 0.3);
          border-top: 4px solid white;
          border-radius: 50%;
          width: 40px;
          height: 40px;
          animation: spin 1s linear infinite;
          margin: 20px auto;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>✅ Logged in as ${loggedInUser.email}</h2>
        <p>Session ID: <strong>${sessionId}</strong></p>
        <div class="spinner"></div>
        <p>Redirecting you back to the app...</p>
      </div>
      <script>
        // Store session ID in localStorage
        localStorage.setItem('userSessionId', '${sessionId}');
        
        // Redirect to home page after 2 seconds
        setTimeout(() => {
          window.location.href = '/';
        }, 2000);
      </script>
    </body>
    </html>
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