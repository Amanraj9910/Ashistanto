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

// Logout endpoint - clear session
router.post('/logout', (req, res) => {
  const { sessionId } = req.body;

  if (sessionId && userTokenStore.has(sessionId)) {
    userTokenStore.delete(sessionId);
    console.log('✅ Session cleared:', sessionId);
  }

  // Clear the logged in user
  loggedInUser = null;

  res.json({ success: true, message: 'Logged out successfully' });
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

    loggedInUser = {
      sessionId: sessionId,
      accessToken: tokens.accessToken,
      refreshToken: tokens.refreshToken,
      email: tokens.account.username
    };

    // Store tokens with expiration metadata for automatic refresh
    userTokenStore.set(sessionId, {
      accessToken: tokens.accessToken,
      refreshToken: tokens.refreshToken,
      expiresAt: Date.now() + ((tokens.expiresIn || 3600) * 1000), // Default 1 hour if not provided
      email: tokens.account.username
    });

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
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Login Successful - Ashistanto</title>
      <link rel="icon" href="/img/favicon.ico.png" type="image/png">
      <script src="https://cdn.tailwindcss.com"></script>
      <style>
        @keyframes checkmark {
          0% { transform: scale(0) rotate(45deg); }
          50% { transform: scale(1.2) rotate(45deg); }
          100% { transform: scale(1) rotate(45deg); }
        }
        
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(20px); }
          to { opacity: 1; transform: translateY(0); }
        }
        
        .checkmark-circle {
          animation: fadeIn 0.5s ease-out;
        }
        
        .checkmark {
          animation: checkmark 0.6s ease-out 0.3s both;
        }
        
        .fade-in {
          animation: fadeIn 0.8s ease-out 0.5s both;
        }
      </style>
    </head>
    <body class="min-h-screen bg-gradient-to-br from-gray-50 via-gray-100 to-gray-200">
      <!-- Background decorative elements -->
      <div class="fixed inset-0 overflow-hidden pointer-events-none">
        <div class="absolute top-20 left-10 w-72 h-72 bg-red-200 rounded-full mix-blend-multiply filter blur-xl opacity-20 animate-pulse"></div>
        <div class="absolute top-40 right-10 w-72 h-72 bg-red-300 rounded-full mix-blend-multiply filter blur-xl opacity-20 animate-pulse" style="animation-delay: 2s;"></div>
        <div class="absolute bottom-20 left-1/2 w-72 h-72 bg-red-400 rounded-full mix-blend-multiply filter blur-xl opacity-20 animate-pulse" style="animation-delay: 4s;"></div>
      </div>

      <div class="relative min-h-screen flex items-center justify-center px-4">
        <div class="max-w-md w-full">
          <!-- Success card -->
          <div class="bg-white rounded-2xl shadow-2xl p-8 backdrop-blur-sm bg-opacity-95 text-center">
            <!-- Success checkmark -->
            <div class="checkmark-circle mb-6">
              <div class="w-24 h-24 mx-auto bg-gradient-to-br from-green-400 to-green-600 rounded-full flex items-center justify-center shadow-lg">
                <svg class="checkmark w-12 h-12 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M5 13l4 4L19 7"></path>
                </svg>
              </div>
            </div>

            <!-- Success message -->
            <div class="fade-in">
              <h2 class="text-3xl font-bold text-gray-800 mb-3">Login Successful!</h2>
              <p class="text-gray-600 mb-2">Welcome back,</p>
              <p class="text-lg font-semibold text-transparent bg-clip-text bg-gradient-to-r from-red-500 to-red-600 mb-6">
                ${loggedInUser.email}
              </p>

              <!-- Loading indicator -->
              <div class="flex items-center justify-center gap-2 text-gray-500 mb-4">
                <svg class="animate-spin h-5 w-5 text-red-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                  <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                  <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
                <span class="text-sm">Redirecting to your dashboard...</span>
              </div>

              <!-- Progress bar -->
              <div class="w-full bg-gray-200 rounded-full h-2 overflow-hidden">
                <div class="bg-gradient-to-r from-red-500 to-red-600 h-2 rounded-full transition-all duration-2000" style="width: 0%; animation: progress 2s ease-out forwards;"></div>
              </div>
            </div>
          </div>

          <!-- Footer -->
          <div class="text-center mt-6 text-gray-600 text-sm fade-in">
            <p>🔒 Secure session established</p>
          </div>
        </div>
      </div>

      <style>
        @keyframes progress {
          from { width: 0%; }
          to { width: 100%; }
        }
      </style>

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
  const tokenData = userTokenStore.get(sessionId);

  if (!tokenData) {
    return res.status(404).json({ error: 'Session not found or token expired' });
  }

  // Return access token from the token object
  res.json({ accessToken: tokenData.accessToken });
});

module.exports = { router, loggedInUser, userTokenStore };