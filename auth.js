const { google } = require('googleapis');
const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const SCOPES = ['https://www.googleapis.com/auth/gmail.modify'];
const TOKEN_PATH = path.join(__dirname, 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
const PORT = 3000;
const REDIRECT_URI = `http://localhost:${PORT}`;

function isInvalidGrant(err) {
  const msg = String(err.message || err).toLowerCase();
  if (msg.includes('invalid_grant')) return true;
  const data = err.response && err.response.data;
  return data && data.error === 'invalid_grant';
}

async function authenticate() {
  const raw = fs.readFileSync(CREDENTIALS_PATH, 'utf8');
  const credentials = JSON.parse(raw);
  const { client_secret, client_id } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, REDIRECT_URI);

  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
    oAuth2Client.setCredentials(token);
    try {
      await oAuth2Client.getAccessToken();
      return oAuth2Client;
    } catch (err) {
      if (!isInvalidGrant(err)) throw err;
      console.log('Gmail token expired or revoked — signing in again...');
      fs.unlinkSync(TOKEN_PATH);
    }
  }

  return getNewToken(oAuth2Client);
}

async function getNewToken(oAuth2Client) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: SCOPES,
  });

  const code = await new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const url = new URL(req.url, REDIRECT_URI);
      const code = url.searchParams.get('code');
      const error = url.searchParams.get('error');

      res.writeHead(200, { 'Content-Type': 'text/html' });
      if (code) {
        res.end('<h2>Auth complete — you can close this tab.</h2>');
        server.close();
        resolve(code);
      } else {
        res.end(`<h2>Auth failed: ${error}</h2>`);
        server.close();
        reject(new Error(`OAuth error: ${error}`));
      }
    });

    server.listen(PORT, () => {
      console.log(`Opening browser for Google auth...`);
      const open = process.platform === 'win32'
        ? `start "" "${authUrl}"`
        : process.platform === 'darwin'
        ? `open "${authUrl}"`
        : `xdg-open "${authUrl}"`;
      exec(open);
    });
  });

  const { tokens } = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokens);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
  console.log('Token saved to', TOKEN_PATH);
  return oAuth2Client;
}

module.exports = { authenticate };
