require('dotenv').config({ override: true });
const { google } = require('googleapis');
const { authenticate } = require('./auth');

async function fetchAllUnread(gmail) {
  const threads = [];
  let pageToken = undefined;
  do {
    const res = await gmail.users.threads.list({
      userId: 'me',
      q: 'is:unread in:inbox -label:CleanupQueue',
      maxResults: 500,
      ...(pageToken && { pageToken }),
    });
    threads.push(...(res.data.threads || []));
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return threads;
}

async function main() {
  const auth = await authenticate();
  const gmail = google.gmail({ version: 'v1', auth });

  console.log('Fetching remaining unread threads...\n');
  const threads = await fetchAllUnread(gmail);
  console.log(`Found ${threads.length} threads.\n`);
  console.log('─'.repeat(80));

  for (const thread of threads) {
    const res = await gmail.users.threads.get({
      userId: 'me',
      id: thread.id,
      format: 'metadata',
      metadataHeaders: ['From', 'Subject'],
    });
    const msg = res.data.messages?.[0];
    if (!msg) continue;
    const get = name => (msg.payload.headers.find(h => h.name.toLowerCase() === name) || {}).value || '';
    console.log(`From:    ${get('from')}`);
    console.log(`Subject: ${get('subject')}`);
    console.log('─'.repeat(80));
  }
}

main().catch(err => { console.error('Fatal:', err.message); process.exit(1); });
