require('dotenv').config({ override: true });
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const { authenticate } = require('./auth');
const { ruleClassify, aiBatchClassify } = require('./classify');

const DECISIONS_LOG = path.join(__dirname, 'decisions.log');
const BATCH_SIZE = 50;

async function fetchAllUnreadThreads(gmail) {
  const threads = [];
  let pageToken = undefined;

  do {
    const res = await gmail.users.threads.list({
      userId: 'me',
      q: 'is:unread in:inbox -label:CleanupQueue',
      maxResults: BATCH_SIZE,
      ...(pageToken && { pageToken }),
    });
    threads.push(...(res.data.threads || []));
    pageToken = res.data.nextPageToken;
  } while (pageToken);

  return threads;
}

async function getThreadInfo(gmail, threadId) {
  const res = await gmail.users.threads.get({
    userId: 'me',
    id: threadId,
    format: 'metadata',
    metadataHeaders: ['From', 'Subject', 'Date', 'List-Unsubscribe', 'List-Id', 'Precedence'],
  });

  const messages = res.data.messages || [];
  if (!messages.length) return null;

  const msg = messages[0];
  const headers = msg.payload.headers || [];
  const get = name => (headers.find(h => h.name.toLowerCase() === name.toLowerCase()) || {}).value || '';
  const listUnsubscribe = get('list-unsubscribe').trim();

  return {
    id: threadId,
    from: get('from'),
    subject: get('subject'),
    date: get('date'),
    snippet: (msg.snippet || '').slice(0, 800),
    messageCount: messages.length,
    labelIds: res.data.labelIds || [],
    listUnsubscribe: !!listUnsubscribe,
    listId: get('list-id').trim(),
    precedence: get('precedence').trim(),
  };
}

function logDecision(info, result) {
  const entry = {
    ts: new Date().toISOString(),
    id: info.id,
    from: info.from,
    subject: info.subject,
    date: info.date,
    decision: result.decision,
    reason: result.reason,
    method: result.method,
  };
  fs.appendFileSync(DECISIONS_LOG, JSON.stringify(entry) + '\n');

  const tag = result.method === 'rule' ? '[rule]' : '[ai]  ';
  const verdict = result.decision === 'delete' ? 'DELETE' : 'KEEP  ';
  console.log(`${verdict} ${tag} ${info.subject || '(no subject)'}`);
  console.log(`         From: ${info.from}`);
  console.log(`         Why:  ${result.reason}`);
}

async function getOrCreateLabel(gmail, name) {
  const res = await gmail.users.labels.list({ userId: 'me' });
  const existing = (res.data.labels || []).find(l => l.name === name);
  if (existing) return existing.id;

  const created = await gmail.users.labels.create({
    userId: 'me',
    requestBody: { name, labelListVisibility: 'labelShow', messageListVisibility: 'show' },
  });
  console.log(`Created label: ${name}`);
  return created.data.id;
}

async function queueForCleanup(gmail, threadId, labelId) {
  await gmail.users.threads.modify({
    userId: 'me',
    id: threadId,
    requestBody: {
      addLabelIds: [labelId],
      removeLabelIds: ['INBOX'],
    },
  });
}

async function main() {
  const provider = (process.env.AI_PROVIDER || 'gemini').toLowerCase();
  if (provider === 'gemini' && !process.env.GEMINI_API_KEY) {
    console.error('Error: GEMINI_API_KEY is not set. Add it to your .env file.');
    console.error('Get a key at https://aistudio.google.com/apikey');
    process.exit(1);
  }
  if (provider === 'anthropic' && !process.env.ANTHROPIC_API_KEY) {
    console.error('Error: ANTHROPIC_API_KEY is not set. Add it to your .env file.');
    process.exit(1);
  }
  if (provider !== 'gemini' && provider !== 'anthropic') {
    console.error(`Error: Unknown AI_PROVIDER "${provider}". Use "gemini" or "anthropic".`);
    process.exit(1);
  }

  console.log(`AI provider: ${provider}\n`);

  const auth = await authenticate();
  const gmail = google.gmail({ version: 'v1', auth });

  const cleanupLabelId = await getOrCreateLabel(gmail, 'CleanupQueue');

  console.log('Fetching all unread inbox threads...');
  const allThreads = await fetchAllUnreadThreads(gmail);

  if (!allThreads.length) {
    console.log('No unread threads found.');
    return;
  }

  console.log(`Found ${allThreads.length} threads total. Processing in batches of ${BATCH_SIZE}...\n`);

  let totalKept = 0, totalQueued = 0, totalErrors = 0, batch = 1;

  for (let i = 0; i < allThreads.length; i += BATCH_SIZE) {
    const threads = allThreads.slice(i, i + BATCH_SIZE);
    console.log(`\n── Batch ${batch} (${i + 1}–${Math.min(i + BATCH_SIZE, allThreads.length)} of ${allThreads.length}) ──────────────────`);
    console.log(`Found ${threads.length} threads. Classifying...\n`);

  // Pass 1: fetch metadata + rule classify
  const infos = [];
  const ruleResults = new Map();
  const unclassified = [];

  for (const thread of threads) {
    try {
      const info = await getThreadInfo(gmail, thread.id);
      if (!info) continue;
      infos.push(info);

      const rule = ruleClassify(info);
      if (rule) {
        ruleResults.set(info.id, rule);
      } else {
        unclassified.push(info);
      }
    } catch (err) {
      console.error(`Error fetching thread ${thread.id}: ${err.message}`);
    }
  }

  // Pass 2: batch AI for unclassified
  const aiResults = new Map();
  if (unclassified.length) {
    console.log(`Pass 1: ${ruleResults.size} rule-classified, ${unclassified.length} sent to AI...\n`);
    try {
      const decisions = await aiBatchClassify(unclassified);
      for (const d of decisions) {
        const info = unclassified[d.index];
        if (info) aiResults.set(info.id, { decision: d.decision, reason: d.reason, method: 'ai' });
      }
    } catch (err) {
      console.error(`AI batch failed: ${err.message} — defaulting unclassified to keep`);
      for (const info of unclassified) {
        aiResults.set(info.id, { decision: 'keep', reason: 'AI error — defaulting to keep', method: 'ai' });
      }
    }
  }

  // Apply decisions
  let kept = 0, queued = 0, errors = 0;

  for (const info of infos) {
    const result = ruleResults.get(info.id) || aiResults.get(info.id) || { decision: 'keep', reason: 'no result', method: 'ai' };
    logDecision(info, result);
    try {
      if (result.decision === 'delete') {
        await queueForCleanup(gmail, info.id, cleanupLabelId);
        queued++;
      } else {
        kept++;
      }
    } catch (err) {
      console.error(`Error applying decision for ${info.id}: ${err.message}`);
      errors++;
    }
  }

  totalKept += kept;
  totalQueued += queued;
  totalErrors += errors;
  batch++;

  console.log(`\nBatch summary — Keep: ${kept}, CleanupQueue: ${queued}${errors ? `, Errors: ${errors}` : ''}`);
  }

  console.log('\n══ Final Summary ═════════════');
  console.log(`Keep:          ${totalKept}`);
  console.log(`CleanupQueue:  ${totalQueued}`);
  if (totalErrors) console.log(`Errors:        ${totalErrors}`);
  console.log(`\nDecisions logged to: ${DECISIONS_LOG}`);
}

main().catch(err => {
  console.error('Fatal:', err.message);
  process.exit(1);
});
