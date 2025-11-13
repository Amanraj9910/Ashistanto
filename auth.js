const express = require('express');
const session = require('express-session');
const helmet = require('helmet');
const { getAuthUrl, getAccessTokenByAuthCode } = require('./graph-tools');
const app = express();
const router = express.Router();

// Use helmet for security headers
app.use(helmet());

// Session config (use RedisStore in production)
app.use(session({
  secret: process.env.SESSION_SECRET || 'yourSecretHere',
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: true,      // HTTPS only
    httpOnly: true,    // Prevent JS access (XSS)
    sameSite: 'lax',   // CSRF protection
    maxAge: 3600000    // 1 hour
  }
}));

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

    // Save tokens in server-side session
    req.session.isAuthenticated = true;
    req.session.account = {
      email: tokens.account.username,
      accessToken: tokens.accessToken,
      refreshToken: tokens.refreshToken,
      tokenExpiry: Date.now() + (tokens.expiresIn * 1000)
    };

    // Redirect back to main app URL
    return res.redirect('https://microsoft-agent-aubbhefsbzagdhha.eastus-01.azurewebsites.net');

  } catch (err) {
    console.error('❌ Login failed:', err);
    return res.status(500).send('Login failed.');
  }
});

// Auth check middleware
function isAuthenticated(req, res, next) {
  if (!req.session.isAuthenticated) {
    return res.status(401).json({ error: 'User not logged in' });
  }
  next();
}

// Step 3: API – Get logged in user
router.get('/user', isAuthenticated, (req, res) => {
  res.json(req.session.account);
});

// Step 4: API – Get access token (protected)
router.get('/access-token', isAuthenticated, (req, res) => {
  res.json({ accessToken: req.session.account.accessToken });
});

// Attach router to app root
app.use('/', router);

module.exports = app;
