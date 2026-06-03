// Gmail Cleanup Agent v4 (Gemini multi-label)
// Pass 1: rule-based (classifyInbox)
// Pass 2: AI assigns your existing Gmail labels (classifyUnclassified)
// Run order: resetLabel -> setupLabel -> classifyInbox -> classifyUnclassified
//
// Script properties:
//   GEMINI_API_KEY              — https://aistudio.google.com/apikey
//   GEMINI_MODEL                — optional, default gemini-2.5-flash
//   DASHBOARD_SPREADSHEET_ID    — optional; auto-created on first log if unset
//   WEBAPP_API_SECRET           — required for Web App API (doPost); use a long random string
//   SCHEDULE_CONFIG               — JSON: { taskTypes, frequency, targetDays, targetHour }

const LABEL_NAME = 'CleanupQueue';
const SPECIAL_KEEP = 'Keep';
const SPECIAL_DELETE = 'Delete';
const BATCH_SIZE = 200;
const AI_BATCH_SIZE = 50;
const READ_RULE_BATCH_SIZE = 10;
const READ_AI_BATCH_SIZE = 5;
const READ_EMAILS_SEARCH_QUERY = 'is:read in:inbox -label:CleanupQueue';
const AI_SNIPPET_CHARS = 800;
const DEFAULT_GEMINI_MODEL = 'gemini-2.5-flash';
const MAX_SCHEMA_ENUM_LABELS = 128;

// --- Dashboard logging (must be in this file for classifyInbox / classifyUnclassified) ---

const DASHBOARD_LOG_SHEET = 'Classification Log';
const DASHBOARD_SPREADSHEET_TITLE = 'Gmail Cleanup Dashboard';
const DASHBOARD_HEADERS = ['Timestamp', 'Source', 'Sender', 'Subject', 'Action', 'Reason', 'Applied', 'Thread ID'];
const LOG_COL_SOURCE = 1;
const LOG_COL_ACTION = 4;
const LOG_COL_REASON = 5;
const LOG_COL_APPLIED = 6;
const LOG_COL_THREAD_ID = 7;
const SCHEDULE_SOURCE = 'Scheduler';
const LAST_CLEANUP_RUN_KEY = 'LAST_CLEANUP_RUN_AT';

/**
 * Resolves the log spreadsheet: Script property → bound sheet → auto-create.
 * Run setupDashboard() once if you want to create/configure it manually.
 */
function ensureDashboardSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DASHBOARD_SPREADSHEET_ID');

  if (id) {
    try {
      return SpreadsheetApp.openById(id.trim());
    } catch (e) {
      Logger.log('Invalid DASHBOARD_SPREADSHEET_ID — will create a new dashboard.');
      props.deleteProperty('DASHBOARD_SPREADSHEET_ID');
    }
  }

  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      props.setProperty('DASHBOARD_SPREADSHEET_ID', active.getId());
      Logger.log('Dashboard bound to active spreadsheet: ' + active.getUrl());
      return active;
    }
  } catch (e) {
    // Standalone script — no container spreadsheet.
  }

  var created = SpreadsheetApp.create(DASHBOARD_SPREADSHEET_TITLE);
  props.setProperty('DASHBOARD_SPREADSHEET_ID', created.getId());
  Logger.log('Created dashboard spreadsheet: ' + created.getUrl());
  Logger.log('Saved DASHBOARD_SPREADSHEET_ID in Script Properties.');
  return created;
}

/** One-time setup: create or bind the dashboard and ensure the log tab exists. */
function setupDashboard() {
  var ss = ensureDashboardSpreadsheet();
  getDashboardLogSheet(ss);
  Logger.log('Dashboard ready: ' + ss.getUrl());
  Logger.log('Log tab: "' + DASHBOARD_LOG_SHEET + '"');
  return ss.getUrl();
}

function getDashboardLogSheet(ss) {
  var sheet = ss.getSheetByName(DASHBOARD_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_LOG_SHEET);
  }
  ensureDashboardLogHeaders_(sheet);
  return sheet;
}

function ensureDashboardLogHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(DASHBOARD_HEADERS);
    sheet.setFrozenRows(1);
    return;
  }
  var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), DASHBOARD_HEADERS.length)).getValues()[0];
  if (String(existing[LOG_COL_APPLIED] || '') !== 'Applied' ||
      String(existing[LOG_COL_THREAD_ID] || '') !== 'Thread ID') {
    sheet.getRange(1, 1, 1, DASHBOARD_HEADERS.length).setValues([DASHBOARD_HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function buildDashboardLogRow_(source, sender, subject, action, reason, thread, applied) {
  var threadId = thread && thread.getId ? thread.getId() : '';
  return [
    new Date(),
    source,
    sender,
    subject,
    action,
    reason,
    applied ? 'yes' : 'no',
    threadId,
  ];
}

function isKeepAction_(action) {
  return String(action || '').trim().toLowerCase() === 'keep';
}

function isSkippedAction_(action) {
  return String(action || '').trim().toLowerCase() === 'skipped';
}

function isLogAppliedFlag_(value) {
  var flag = String(value || '').trim().toLowerCase();
  return flag === 'yes' || flag === 'no';
}

/** Thread ID column, or legacy column G before Applied was added. */
function getLogThreadId_(row) {
  var threadId = String(row[LOG_COL_THREAD_ID] || '').trim();
  if (threadId) {
    return threadId;
  }
  var legacy = row[LOG_COL_APPLIED];
  if (!isLogAppliedFlag_(legacy)) {
    return String(legacy || '').trim();
  }
  return '';
}

/**
 * Whether a log row represents a label the script actually applied (archive-eligible).
 * Uses Applied column when present; falls back for legacy rows without it.
 */
function isScriptLabelAppliedAction_(action, source, applied) {
  if (isLogAppliedFlag_(applied)) {
    return String(applied || '').trim().toLowerCase() === 'yes';
  }

  if (isKeepAction_(action) || isSkippedAction_(action)) {
    return false;
  }

  var normalized = String(action || '').trim();
  if (!normalized) {
    return false;
  }

  if (normalized === LABEL_NAME || normalized.toLowerCase() === 'delete') {
    return true;
  }

  var sourceName = String(source || '').trim();
  if (sourceName === SCHEDULE_SOURCE) {
    return false;
  }

  // Rule-Based pass only logs Keep or CleanupQueue.
  if (sourceName === 'Rule-Based') {
    return normalized === LABEL_NAME;
  }

  // Gemini pass: trust non-Keep script rows (covers labels later deleted from Gmail).
  if (sourceName === 'Gemini AI') {
    return true;
  }

  var labels = getExistingLabels();
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].toLowerCase() === normalized.toLowerCase()) {
      return true;
    }
  }
  return false;
}

/**
 * Appends classification rows to the dashboard sheet.
 * Each entry: [Date, source, sender, subject, action, reason, applied, threadId]
 * @param {Array<Array>} logEntries
 */
function logToDashboard(logEntries) {
  if (!logEntries || logEntries.length === 0) {
    return;
  }

  var ss;
  try {
    ss = ensureDashboardSpreadsheet();
  } catch (e) {
    Logger.log('logToDashboard failed: ' + e.message);
    return;
  }

  var sheet = getDashboardLogSheet(ss);

  var rows = logEntries.map(function(entry) {
    return [
      entry[0] instanceof Date ? entry[0] : new Date(entry[0]),
      String(entry[1] || ''),
      String(entry[2] || ''),
      String(entry[3] || ''),
      String(entry[4] || ''),
      String(entry[5] || ''),
      String(entry[6] || ''),
      String(entry[7] || ''),
    ];
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, DASHBOARD_HEADERS.length).setValues(rows);
  Logger.log('Dashboard: appended ' + rows.length + ' row(s) → ' + ss.getUrl());
}

function logSchedulerEvent_(action, reason) {
  logToDashboard([buildDashboardLogRow_(
    SCHEDULE_SOURCE,
    'System',
    'Automated task',
    String(action || 'Event'),
    String(reason || ''),
    null,
    false
  )]);
}

const DELETE_DOMAINS = [
  'substack.com', 'beehiiv.com', 'mailchimp.com', 'sendgrid.net',
  'constantcontact.com', 'klaviyo.com', 'campaigns.amazon.com',
  'marketing.amazon.com', 'news.linkedin.com', 'e.linkedin.com',
  'em.linkedin.com', 'bounce.linkedin.com', 'notifications.google.com',
  'mailer.robinhood.com', 'email.robinhood.com',
];

const DELETE_SENDERS = [
  'no-reply@news.htzone.co.il', 'lenovo@ecomm.lenovo.com',
  'hi@mail.theresanaiforthat.com', 'jobalerts-noreply@linkedin.com',
  'notifications-noreply@linkedin.com', 'messages-noreply@linkedin.com',
  'groups-noreply@linkedin.com', 'jobs-listings@linkedin.com',
  'newsletters-noreply@linkedin.com', 'kayak@msg.kayak.com',
  'hello@duolingo.com', 'robot@letterboxd.com', 'news@letterboxd.com',
  'noreply@lovable.dev', 'events@marketing.descript.com',
  'newsletter@marketing.descript.com', 'shop@sensibo.com',
  'no-reply@marketing.base44.com', 'support@team.promo.com',
  'no-reply@marketing.lyftmail.com', 'notifications@mail.postman.com',
  'news@m.elements.envato.com', 'your@insights.veed.io',
  'insights@insights.veed.io', 'team@newsletter.artlist.io',
  'nextplayso@substack.com', 'nextplayso+should-you-join@substack.com',
  'cheapsoftwarestocks@substack.com', 'info@e.tikvah.org', 'no-reply@substack.com',
  'support@beenverified-newsletter.com', 'info.asif.group@send.vpcontact.com',
  'books@greenbrothers.co.il', 'info@delta.co.il', 'Info@alinasleep.co.il',
  'info@kapra.co.il', 'info@kitchenware.co.il', 'info@teimim.com',
  'info@tevabari.co.il', 'info@zap.co.il', 'do-not-reply@pandazzz.co.il',
  'itay@ets-hasade.com', 'info@ortokal.co.il', 'info@marathonisrael.co.il',
  'noam@marketing.shukcity.co.il', 'newsletter@cwc.co.il',
  'info@drzivav.com', 'briaa@briaa-gilboa.co.il',
  'Benny@ai-israel.minisite.ms', 'amitayboneh@substack.com',
  'qed@nfx.com', 'ship@info.vercel.com', 'hello@apify.com',
  'welcome@supabase.com', 'zeno@updates.resend.com',
  'no-reply@mail.nordvpn.com', 'privacy@transactional.life360.com',
  'v_chloe.ropa@oneforma.com', 'v_jenelyn.acob@oneforma.com',
  'sam.sornito@oneforma.com', 'newsletter@easeus.com',
  'info@email.trainingpeaks.com', 'run@coopah.com',
  'hello@mail.grammarly.com', 'premium@academia-mail.com',
  'bounced@midgampanel.com', 'dtc-email@discounttire-email.com',
  'office@barakhass.co.il', 'ilan.haran@hacara.org.il',
  'communications@mail.aircanada.com', 'galb@log-on.com',
  'yossi@civi.co.il', 'wehave@skeelz.co.il', 'noreply@mc.sba.org.il',
  'tanach.israeli.gmail.com@send.vpcontact.com', 'andy@andyfrisella.com',
  'marketing@pilat.co.il', 'noreply@foodsdictionary.com',
  'opportunities@careeralerts.careers.hpe.com',
  'rinaspap2-jobnotification@noreply55.jobs2web.com',
  'Contact.play-back.co.il@send.vpcontact.com',
  'messaging-digest-noreply@linkedin.com',
  'tommy.barav@magical.team', 'hello@letterboxd.com',
  'yehonatan@ti-space.com', 'hello@gamma.app',
  'email@mail.pmi.org', 'newsletter@mail.bubble.io',
];

const DELETE_SUBJECT_KEYWORDS = [
  'פרסומת', 'פרסום', 'מבצע', 'הנחה',
  'sale', 'unsubscribe', 'newsletter', 'weekly digest',
  'weekly roundup', 'monthly digest', 'special offer',
  'limited time', 'exclusive deal', 'flash sale',
  '% off', 'promo code', 'coupon', 'save big',
  'black friday', 'cyber monday', "you're invited",
  "don't miss out", 'last chance', 'job alert', 'jobs at', 'is hiring',
];

const KEEP_DOMAINS = ['gmail.com'];

const KEEP_SENDERS = [
  'rantzabar1@gmail.com', 't-namal@univ.haifa.ac.il',
  'tayaruthaifa@univ.haifa.ac.il', 'm_avshalom@avshalom-inst.co.il',
  'mailer@2sign.co.il', 'HarelInsurance@harel-group.co.il',
  'Donotreply2@aig.co.il', 'noreply@ionos.com',
  'invoice+statements@lovable.dev', 'invoice+statements@supabase.com',
  'invoice+statements@mail.anthropic.com', 'invoice+statements@captions.ai',
  'donotreply@rivhit.co.il', 'noreply@info.iherb.com',
  'noReply@postil.co.il', 'mailer@e-shops.co.il',
  'noreply@iw.menoramivt.co.il', 'no-reply@ycombinator.com',
  'drive-shares-dm-noreply@google.com',
  'store+79639380258@t.shopifyemail.com',
  'no-reply@email.claude.com',
];

// --- Rule lookup indexes (built once at load) ---

function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function toLowerSet(items) {
  var set = {};
  items.forEach(function(item) {
    set[item.toLowerCase()] = true;
  });
  return set;
}

function sortByLengthDesc(items) {
  return items.slice().sort(function(a, b) { return b.length - a.length; });
}

function buildDomainIndex(domains) {
  var lowered = domains.map(function(d) { return d.toLowerCase(); });
  return {
    exact: toLowerSet(lowered),
    suffixes: sortByLengthDesc(lowered),
  };
}

function buildSubjectKeywordRegex(keywords) {
  if (!keywords.length) return null;
  var pattern = keywords.map(function(kw) { return escapeRegExp(kw); }).join('|');
  return new RegExp(pattern, 'i');
}

var DELETE_SENDERS_SET = toLowerSet(DELETE_SENDERS);
var KEEP_SENDERS_SET = toLowerSet(KEEP_SENDERS);
var DELETE_SENDERS_DESC = sortByLengthDesc(DELETE_SENDERS.map(function(s) { return s.toLowerCase(); }));
var KEEP_SENDERS_DESC = sortByLengthDesc(KEEP_SENDERS.map(function(s) { return s.toLowerCase(); }));
var DELETE_DOMAIN_INDEX = buildDomainIndex(DELETE_DOMAINS);
var KEEP_DOMAIN_INDEX = buildDomainIndex(KEEP_DOMAINS);
var DELETE_SUBJECT_REGEX = buildSubjectKeywordRegex(DELETE_SUBJECT_KEYWORDS);

/** Gmail system category labels that map to Promotions / Social tabs. */
var GMAIL_AUTO_DELETE_CATEGORIES = {
  'CATEGORY_PROMOTIONS': true,
  'CATEGORY_SOCIAL': true,
};

// --- Helpers ---

function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function getGeminiModel() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || DEFAULT_GEMINI_MODEL;
}

function extractEmailAddress(from) {
  var angle = from.match(/<([^>]+)>/);
  if (angle) return angle[1].trim().toLowerCase();
  var bare = from.match(/[\w.+-]+@[\w.-]+/);
  return bare ? bare[0].toLowerCase() : from.toLowerCase().trim();
}

function extractDomain(from) {
  var email = extractEmailAddress(from);
  var at = email.lastIndexOf('@');
  return at === -1 ? '' : email.slice(at + 1);
}

function domainMatchesIndex(domain, index) {
  if (!domain) return false;
  if (index.exact[domain]) return true;
  var suffixes = index.suffixes;
  for (var i = 0; i < suffixes.length; i++) {
    var pattern = suffixes[i];
    var suffix = '.' + pattern;
    if (domain.length > suffix.length &&
        domain.substring(domain.length - suffix.length) === suffix) {
      return true;
    }
  }
  return false;
}

function senderMatches(fromLower, email, exactSet, listDesc) {
  if (email && exactSet[email]) return true;
  for (var i = 0; i < listDesc.length; i++) {
    if (fromLower.indexOf(listDesc[i]) !== -1) return true;
  }
  return false;
}

/**
 * Parses the header block from a raw RFC 822 message string.
 * @param {string} raw
 * @returns {Object.<string, string>}
 */
function parseRawMessageHeaders_(raw) {
  var headers = {};
  if (!raw) {
    return headers;
  }
  var headerBlock = String(raw).split(/\r?\n\r?\n/)[0] || '';
  var lines = headerBlock.split(/\r?\n/);
  var currentName = '';
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (/^[\s\t]/.test(line) && currentName) {
      headers[currentName] += ' ' + line.trim();
      continue;
    }
    var colon = line.indexOf(':');
    if (colon === -1) {
      continue;
    }
    currentName = line.slice(0, colon).trim().toLowerCase();
    headers[currentName] = line.slice(colon + 1).trim();
  }
  return headers;
}

/**
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 * @returns {Object.<string, string>}
 */
function getMessageHeaders_(message) {
  if (!message) {
    return {};
  }
  try {
    return parseRawMessageHeaders_(message.getRawContent());
  } catch (e) {
    Logger.log('Header parse failed: ' + e.message);
    return {};
  }
}

function messageHasListUnsubscribe_(message) {
  var value = getMessageHeaders_(message)['list-unsubscribe'] || '';
  return !!String(value).trim();
}

/**
 * @param {GoogleAppsScript.Gmail.GmailThread} thread
 * @returns {string[]}
 */
function getThreadGmailCategoryNames_(thread) {
  var names = [];
  if (!thread || !thread.getLabels) {
    return names;
  }
  try {
    thread.getLabels().forEach(function(label) {
      var name = label.getName();
      if (GMAIL_AUTO_DELETE_CATEGORIES[name]) {
        names.push(GMAIL_AUTO_DELETE_CATEGORIES[name]);
      }
    });
  } catch (e) {
    Logger.log('Category label read failed: ' + e.message);
  }
  return names;
}

/**
 * @param {GoogleAppsScript.Gmail.GmailThread} thread
 * @returns {boolean}
 */
function threadHasAutoDeleteGmailCategory_(thread) {
  return getThreadGmailCategoryNames_(thread).length > 0;
}

/**
 * @param {string} sender
 * @param {string} subject
 * @param {GoogleAppsScript.Gmail.GmailMessage=} message
 * @param {GoogleAppsScript.Gmail.GmailThread=} thread
 * @returns {{verdict: string, reason: string}|null}
 */
function classifyByRules(sender, subject, message, thread) {
  var fromLower = sender.toLowerCase();
  var email = extractEmailAddress(sender);
  var domain = extractDomain(sender);
  var subjectLower = subject.toLowerCase();

  if (senderMatches(fromLower, email, KEEP_SENDERS_SET, KEEP_SENDERS_DESC)) {
    return { verdict: 'keep', reason: 'Matched keep sender' };
  }
  if (domainMatchesIndex(domain, KEEP_DOMAIN_INDEX)) {
    return { verdict: 'keep', reason: 'Matched keep domain' };
  }

  if (message && messageHasListUnsubscribe_(message)) {
    return { verdict: 'delete', reason: 'List-Unsubscribe header' };
  }

  if (thread && threadHasAutoDeleteGmailCategory_(thread)) {
    return {
      verdict: 'delete',
      reason: 'Gmail category: ' + getThreadGmailCategoryNames_(thread).join(', '),
    };
  }

  if (senderMatches(fromLower, email, DELETE_SENDERS_SET, DELETE_SENDERS_DESC)) {
    return { verdict: 'delete', reason: 'Matched delete sender' };
  }
  if (domainMatchesIndex(domain, DELETE_DOMAIN_INDEX)) {
    return { verdict: 'delete', reason: 'Matched delete domain' };
  }
  if (DELETE_SUBJECT_REGEX && DELETE_SUBJECT_REGEX.test(subjectLower)) {
    return { verdict: 'delete', reason: 'Matched delete subject keyword' };
  }
  return null;
}

function ruleShouldDelete(sender, subject, message, thread) {
  var result = classifyByRules(sender, subject, message, thread);
  return !!(result && result.verdict === 'delete');
}

/**
 * Fetches custom user label names (excludes CleanupQueue).
 * @returns {string[]}
 */
function getExistingLabels() {
  var names = [];
  GmailApp.getUserLabels().forEach(function(label) {
    var name = label.getName();
    if (name !== LABEL_NAME) {
      names.push(name);
    }
  });
  names.sort();
  return names;
}

/**
 * Case-insensitive map: lowercase name -> { name, label }.
 */
function buildLabelLookup(existingLabels) {
  var lookup = {};
  existingLabels.forEach(function(name) {
    var label = GmailApp.getUserLabelByName(name);
    if (label) {
      lookup[name.toLowerCase()] = { name: name, label: label };
    }
  });
  return lookup;
}

function buildMultiLabelSystemPrompt(existingLabels) {
  var labelLines = existingLabels.length
    ? existingLabels.map(function(l) { return '- ' + l; }).join('\n')
    : '(none configured — use Keep or Delete only)';

  return (
    'You are an email triage assistant. You receive batches of messages with metadata and body snippets.\n' +
    'Analyze each batch on its own: infer who sends bulk mail, which items are newsletters or marketing, ' +
    'which are direct personal or operational mail, and which automated notices look actionable.\n' +
    'Do not assume any fixed job, industry, or inbox profile — reason only from batch evidence and the label list below.\n\n' +
    'USER GMAIL LABELS (exact strings — your only custom taxonomy):\n' +
    labelLines + '\n\n' +
    'SPECIAL VALUES:\n' +
    '- "' + SPECIAL_KEEP + '": Leave in inbox when no user label fits and the message appears individually important ' +
    '(direct human correspondence, actionable requests, reservations or receipts you may need, security or account alerts, ' +
    'time-sensitive coordination). When unsure between a custom label and ' + SPECIAL_KEEP + ', choose ' + SPECIAL_KEEP +
    ' only if importance is clear from the content.\n' +
    '- "' + SPECIAL_DELETE + '": Junk or low-value bulk. Use for promotional blasts, newsletters, drip campaigns, ' +
    'marketing automation, social or network digests, retailer deals, product-update spam, and non-actionable bulk notifications. ' +
    'Be decisive when metadata or copy signals mass mail (List-Unsubscribe, List-Id, marketing footers, percent-off language, etc.).\n\n' +
    'Rules:\n' +
    '- Never invent label names — only strings from the user label list or the two special values.\n' +
    '- Compare each email to patterns across the batch, not stereotypical personas.\n' +
    '- Assign exactly one outcome per email.\n' +
    '- Prefer ' + SPECIAL_DELETE + ' for clear marketing and bulk mail.\n\n' +
    'Respond with ONLY a JSON array — no markdown:\n' +
    '[{"index": 1, "suggestedLabel": "LabelName", "reason": "one short phrase"}, ...]'
  );
}

/**
 * @param {number} index 1-based
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 * @param {GoogleAppsScript.Gmail.GmailThread} thread
 * @returns {string}
 */
function formatEmailForGemini_(index, message, thread) {
  var parts = [
    index + '. From: ' + message.getFrom(),
    'Subject: ' + message.getSubject(),
  ];
  var headers = getMessageHeaders_(message);
  if (headers['list-unsubscribe']) {
    parts.push('List-Unsubscribe: yes');
  }
  if (headers['list-id']) {
    parts.push('List-Id: present');
  }
  if (headers['precedence']) {
    parts.push('Precedence: ' + headers['precedence']);
  }
  var categories = getThreadGmailCategoryNames_(thread);
  if (categories.length) {
    parts.push('Gmail categories: ' + categories.join(', '));
  }
  parts.push('Snippet: ' + (message.getPlainBody() || '').slice(0, AI_SNIPPET_CHARS));
  return parts.join(' | ');
}

function buildGeminiBatchUserPrompt_(threads) {
  var lines = threads.map(function(thread, i) {
    var msg = thread.getMessages()[0];
    return formatEmailForGemini_(i + 1, msg, thread);
  });
  return 'Emails to classify:\n' + lines.join('\n');
}

function buildMultiLabelResponseSchema(existingLabels) {
  var enumValues = existingLabels.slice();
  enumValues.push(SPECIAL_KEEP, SPECIAL_DELETE);

  var suggestedLabelSchema = { type: 'STRING' };
  if (enumValues.length > 0 && enumValues.length <= MAX_SCHEMA_ENUM_LABELS) {
    suggestedLabelSchema.enum = enumValues;
  }

  return {
    type: 'ARRAY',
    items: {
      type: 'OBJECT',
      properties: {
        index: { type: 'INTEGER' },
        suggestedLabel: suggestedLabelSchema,
        reason: { type: 'STRING' },
      },
      required: ['index', 'suggestedLabel', 'reason'],
    },
  };
}

/**
 * Maps AI output to SPECIAL_KEEP, SPECIAL_DELETE, or an exact user label name.
 * @returns {string|null}
 */
function resolveSuggestedLabel(raw, existingLabels) {
  if (raw === undefined || raw === null) return null;
  var normalized = String(raw).trim();
  if (!normalized) return null;

  var lower = normalized.toLowerCase();
  if (lower === 'keep' || lower === 'inbox/keep' || lower === 'inbox') {
    return SPECIAL_KEEP;
  }
  if (lower === 'delete' || lower === 'archive/unclassified' || lower === 'archive') {
    return SPECIAL_DELETE;
  }

  for (var i = 0; i < existingLabels.length; i++) {
    if (existingLabels[i].toLowerCase() === lower) {
      return existingLabels[i];
    }
  }
  return null;
}

/**
 * Call Gemini generateContent; returns response text (JSON array).
 * @param {string} userPrompt
 * @param {string[]} existingLabels
 */
function callGemini(userPrompt, existingLabels) {
  existingLabels = existingLabels || [];

  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is not set in Script Properties');
  }

  var model = getGeminiModel();
  var url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model + ':generateContent?key=' + encodeURIComponent(apiKey);

  var payload = {
    systemInstruction: {
      parts: [{ text: buildMultiLabelSystemPrompt(existingLabels) }],
    },
    contents: [{
      role: 'user',
      parts: [{ text: userPrompt }],
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: buildMultiLabelResponseSchema(existingLabels),
    },
  };

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(payload),
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) {
    Logger.log('Gemini HTTP ' + code + ': ' + body);
    throw new Error('Gemini API error: HTTP ' + code);
  }

  var data = JSON.parse(body);
  var text =
    data.candidates &&
    data.candidates[0] &&
    data.candidates[0].content &&
    data.candidates[0].content.parts &&
    data.candidates[0].content.parts[0] &&
    data.candidates[0].content.parts[0].text;

  if (!text) {
    Logger.log('Gemini response (no text): ' + body);
    throw new Error('No text in Gemini response');
  }
  return text;
}

function parseLabelSuggestionsFromAi(rawText) {
  var match = rawText.match(/\[[\s\S]*\]/);
  if (!match) {
    throw new Error('No JSON array in Gemini response: ' + rawText.slice(0, 200));
  }
  return JSON.parse(match[0]);
}

// --- Pass 1: Rule-based ---

function setupLabel() {
  if (!GmailApp.getUserLabelByName(LABEL_NAME)) {
    GmailApp.createLabel(LABEL_NAME);
    Logger.log('Label "CleanupQueue" created');
  } else {
    Logger.log('Label already exists');
  }
}

function resetLabel() {
  var old = GmailApp.getUserLabelByName('To Delete');
  if (old) {
    old.deleteLabel();
    Logger.log('Removed old "To Delete" label');
  } else {
    Logger.log('No "To Delete" label found');
  }
}

function buildInboxSearchQuery_(includeRead) {
  var query = 'in:inbox -label:CleanupQueue';
  if (!includeRead) {
    query = 'is:unread ' + query;
  }
  return query;
}

function applyCleanupQueueAction_(thread, label, wasUnread) {
  thread.addLabel(label);
  if (wasUnread) {
    thread.moveToArchive();
  }
}

function classifyInbox(includeRead) {
  var logEntries = [];
  var label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    Logger.log('Run setupLabel() first!');
    return;
  }

  var threads = GmailApp.search(buildInboxSearchQuery_(includeRead), 0, BATCH_SIZE);
  var tagged = 0, kept = 0, unclassified = 0;

  threads.forEach(function(thread) {
    var wasUnread = thread.isUnread();
    var msg = thread.getMessages()[0];
    var sender = msg.getFrom();
    var subject = msg.getSubject();
    var ruleResult = classifyByRules(sender, subject, msg, thread);

    if (ruleResult && ruleResult.verdict === 'keep') {
      logEntries.push(buildDashboardLogRow_('Rule-Based', sender, subject, 'Keep', ruleResult.reason, thread, false));
      kept++;
      return;
    }

    if (ruleResult && ruleResult.verdict === 'delete') {
      applyCleanupQueueAction_(thread, label, wasUnread);
      logEntries.push(buildDashboardLogRow_('Rule-Based', sender, subject, LABEL_NAME, ruleResult.reason, thread, true));
      Logger.log('-> CleanupQueue' + (wasUnread ? ' (archived)' : ' (left in inbox)') + ': ' + subject);
      tagged++;
    } else {
      Logger.log('? Unclassified: ' + subject + ' | ' + sender);
      unclassified++;
    }
  });

  logToDashboard(logEntries);

  Logger.log('--- Summary ---');
  Logger.log('Tagged    : ' + tagged);
  Logger.log('Kept      : ' + kept);
  Logger.log('Unclear   : ' + unclassified);
  Logger.log('Total     : ' + threads.length);
}

// --- Pass 2: AI multi-label classification (Gemini) ---

function classifyUnclassified(includeRead) {
  var logEntries = [];
  var cleanupLabel = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!cleanupLabel) {
    Logger.log('Run setupLabel() first!');
    return;
  }
  if (!getGeminiApiKey()) {
    Logger.log('Add GEMINI_API_KEY to Script Properties first!');
    Logger.log('Get a key at https://aistudio.google.com/apikey');
    return;
  }

  var existingLabels = getExistingLabels();
  var labelLookup = buildLabelLookup(existingLabels);

  Logger.log('User labels for AI (' + existingLabels.length + '): ' +
    (existingLabels.length ? existingLabels.join(', ') : '(none)'));

  var threads = GmailApp.search(buildInboxSearchQuery_(includeRead), 0, AI_BATCH_SIZE);
  if (threads.length === 0) {
    Logger.log('No emails remaining' + (includeRead ? '' : ' (unread only)') + '!');
    logToDashboard(logEntries);
    return;
  }

  var userPrompt = buildGeminiBatchUserPrompt_(threads);

  var rawText;
  try {
    rawText = callGemini(userPrompt, existingLabels);
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    return;
  }

  Logger.log('Gemini raw response: ' + rawText);

  var suggestions;
  try {
    suggestions = parseLabelSuggestionsFromAi(rawText);
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    return;
  }

  var labeled = 0, deleted = 0, kept = 0, skipped = 0;

  suggestions.forEach(function(d) {
    var thread = threads[d.index - 1];
    if (!thread) return;

    var wasUnread = thread.isUnread();
    var msg = thread.getMessages()[0];
    var sender = msg.getFrom();
    var subject = msg.getSubject();
    var rawLabel = d.suggestedLabel !== undefined ? d.suggestedLabel : d.decision;
    var decision = d.decision !== undefined ? d.decision : rawLabel;
    var resolved = resolveSuggestedLabel(rawLabel, existingLabels);
    var reason = d.reason || '';

    if (resolved === SPECIAL_DELETE) {
      applyCleanupQueueAction_(thread, cleanupLabel, wasUnread);
      logEntries.push(buildDashboardLogRow_('Gemini AI', sender, subject, LABEL_NAME, reason, thread, true));
      Logger.log('-> ' + LABEL_NAME + ' [Delete' + (wasUnread ? ', archived' : ', left in inbox') + ']: ' +
        subject + ' — ' + reason);
      deleted++;
      return;
    }

    if (resolved === SPECIAL_KEEP) {
      logEntries.push(buildDashboardLogRow_('Gemini AI', sender, subject, 'Keep', reason, thread, false));
      Logger.log('Keep [AI]: ' + subject + ' — ' + reason);
      kept++;
      return;
    }

    var entry = labelLookup[resolved.toLowerCase()];
    if (entry) {
      thread.addLabel(entry.label);
      logEntries.push(buildDashboardLogRow_('Gemini AI', sender, subject, entry.name, reason, thread, true));
      Logger.log('-> Label "' + entry.name + '": ' + subject + ' — ' + reason);
      labeled++;
      return;
    }

    logEntries.push(buildDashboardLogRow_(
      'Gemini AI',
      sender,
      subject,
      'Skipped',
      'Unrecognized label: ' + rawLabel,
      thread,
      false
    ));
    Logger.log('? Unknown label "' + rawLabel + '", left in inbox: ' + subject);
    skipped++;
  });

  logToDashboard(logEntries);

  Logger.log('--- AI Pass Summary ---');
  Logger.log('Labeled    : ' + labeled);
  Logger.log('Deleted    : ' + deleted);
  Logger.log('Kept       : ' + kept);
  Logger.log('Unrecognized: ' + skipped);
  Logger.log('Total      : ' + threads.length);
  Logger.log('Model      : ' + getGeminiModel());
}

/**
 * Read-only inbox classification: rules then AI on read mail only.
 * Query is fixed to is:read in:inbox -label:CleanupQueue — never touches unread mail.
 * Labels only; never archives (threads stay in inbox for review).
 */
function classifyReadEmails() {
  setupLabel();

  var cleanupLabel = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!cleanupLabel) {
    Logger.log('Run setupLabel() first!');
    return;
  }

  markCleanupRunStart_();

  // --- Pass 1: rule-based (read inbox only) ---
  var logEntries = [];
  var tagged = 0, kept = 0, unclassified = 0;
  var ruleThreads = GmailApp.search(READ_EMAILS_SEARCH_QUERY, 0, READ_RULE_BATCH_SIZE);

  ruleThreads.forEach(function(thread) {
    var msg = thread.getMessages()[0];
    var sender = msg.getFrom();
    var subject = msg.getSubject();
    var ruleResult = classifyByRules(sender, subject, msg, thread);

    if (ruleResult && ruleResult.verdict === 'keep') {
      logEntries.push(buildDashboardLogRow_('Rule-Based', sender, subject, 'Keep', ruleResult.reason, thread, false));
      kept++;
      return;
    }

    if (ruleResult && ruleResult.verdict === 'delete') {
      thread.addLabel(cleanupLabel);
      logEntries.push(buildDashboardLogRow_('Rule-Based', sender, subject, LABEL_NAME, ruleResult.reason, thread, true));
      Logger.log('-> CleanupQueue [read, inbox]: ' + subject);
      tagged++;
      return;
    }

    Logger.log('? Unclassified [read]: ' + subject + ' | ' + sender);
    unclassified++;
  });

  logToDashboard(logEntries);

  Logger.log('--- Read Mail Rule Pass ---');
  Logger.log('Tagged    : ' + tagged);
  Logger.log('Kept      : ' + kept);
  Logger.log('Unclear   : ' + unclassified);
  Logger.log('Total     : ' + ruleThreads.length);

  // --- Pass 2: AI multi-label (read inbox only) ---
  if (!getGeminiApiKey()) {
    Logger.log('Skipping AI pass — add GEMINI_API_KEY to Script Properties.');
    return;
  }

  var aiLogEntries = [];
  var existingLabels = getExistingLabels();
  var labelLookup = buildLabelLookup(existingLabels);

  Logger.log('User labels for AI (' + existingLabels.length + '): ' +
    (existingLabels.length ? existingLabels.join(', ') : '(none)'));

  var aiThreads = GmailApp.search(READ_EMAILS_SEARCH_QUERY, 0, READ_AI_BATCH_SIZE);
  if (aiThreads.length === 0) {
    Logger.log('No read emails remaining for AI pass.');
    logToDashboard(aiLogEntries);
    return;
  }

  var rawText;
  try {
    rawText = callGemini(buildGeminiBatchUserPrompt_(aiThreads), existingLabels);
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    return;
  }

  Logger.log('Gemini raw response: ' + rawText);

  var suggestions;
  try {
    suggestions = parseLabelSuggestionsFromAi(rawText);
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    return;
  }

  var labeled = 0, deleted = 0, keptAi = 0, skipped = 0;

  suggestions.forEach(function(d) {
    var thread = aiThreads[d.index - 1];
    if (!thread) return;

    var msg = thread.getMessages()[0];
    var sender = msg.getFrom();
    var subject = msg.getSubject();
    var rawLabel = d.suggestedLabel !== undefined ? d.suggestedLabel : d.decision;
    var decision = d.decision !== undefined ? d.decision : rawLabel;
    var resolved = resolveSuggestedLabel(rawLabel, existingLabels);
    var reason = d.reason || '';

    if (resolved === SPECIAL_DELETE) {
      thread.addLabel(cleanupLabel);
      aiLogEntries.push(buildDashboardLogRow_('Gemini AI', sender, subject, LABEL_NAME, reason, thread, true));
      Logger.log('-> ' + LABEL_NAME + ' [read, inbox]: ' + subject + ' — ' + reason);
      deleted++;
      return;
    }

    if (resolved === SPECIAL_KEEP) {
      aiLogEntries.push(buildDashboardLogRow_('Gemini AI', sender, subject, 'Keep', reason, thread, false));
      Logger.log('Keep [read, AI]: ' + subject + ' — ' + reason);
      keptAi++;
      return;
    }

    var entry = labelLookup[resolved.toLowerCase()];
    if (entry) {
      thread.addLabel(entry.label);
      aiLogEntries.push(buildDashboardLogRow_('Gemini AI', sender, subject, entry.name, reason, thread, true));
      Logger.log('-> Label "' + entry.name + '" [read]: ' + subject + ' — ' + reason);
      labeled++;
      return;
    }

    aiLogEntries.push(buildDashboardLogRow_(
      'Gemini AI',
      sender,
      subject,
      'Skipped',
      'Unrecognized label: ' + rawLabel,
      thread,
      false
    ));
    Logger.log('? Unknown label "' + rawLabel + '", left in inbox: ' + subject);
    skipped++;
  });

  logToDashboard(aiLogEntries);

  Logger.log('--- Read Mail AI Pass ---');
  Logger.log('Labeled    : ' + labeled);
  Logger.log('Deleted    : ' + deleted);
  Logger.log('Kept       : ' + keptAi);
  Logger.log('Unrecognized: ' + skipped);
  Logger.log('Total      : ' + aiThreads.length);
  Logger.log('Model      : ' + getGeminiModel());
}

/** Log all custom labels available to the classifier. */
function listExistingLabels() {
  var labels = getExistingLabels();
  Logger.log('Custom labels (' + labels.length + ', excluding ' + LABEL_NAME + '):');
  labels.forEach(function(name) { Logger.log('  - ' + name); });
}

function runFullCleanup(includeRead) {
  markCleanupRunStart_();
  setupLabel();
  classifyInbox(includeRead);
  classifyUnclassified(includeRead);
}

/**
 * Returns thread IDs the script labeled (not Keep) since the last cleanup run.
 * @returns {string[]}
 */
function getScriptLabeledThreadIdsSinceLastRun_() {
  var lastRunAt = getLastCleanupRunAt_();
  if (!lastRunAt) {
    return [];
  }

  var ss;
  try {
    ss = ensureDashboardSpreadsheet();
  } catch (e) {
    Logger.log('Could not read dashboard log: ' + e.message);
    return [];
  }

  var sheet = ss.getSheetByName(DASHBOARD_LOG_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  var numRows = sheet.getLastRow() - 1;
  var numCols = Math.max(sheet.getLastColumn(), DASHBOARD_HEADERS.length);
  var rawRows = sheet.getRange(2, 1, numRows, numCols).getValues();
  var ids = {};
  var cutoff = lastRunAt;

  for (var i = rawRows.length - 1; i >= 0; i--) {
    var row = rawRows[i];
    var ts = row[0];
    if (!(ts instanceof Date)) {
      ts = new Date(ts);
    }
    if (isNaN(ts.getTime()) || ts < cutoff) {
      continue;
    }

    var action = String(row[LOG_COL_ACTION] || '');
    if (!isScriptLabelAppliedAction_(action, String(row[LOG_COL_SOURCE] || ''), row[LOG_COL_APPLIED])) {
      continue;
    }

    var threadId = getLogThreadId_(row);
    if (threadId) {
      ids[threadId] = true;
    }
  }

  return Object.keys(ids);
}

/**
 * Archives inbox threads the script labeled since the last run (by thread ID in the log).
 * Ignores current labels — includes mail even if you relabeled manually.
 * @returns {{archived: number, skipped: number, message: string}}
 */
function archiveReviewedCleanupItems() {
  var threadIds = getScriptLabeledThreadIdsSinceLastRun_();
  if (!threadIds.length) {
    return {
      archived: 0,
      skipped: 0,
      message: buildArchiveReviewedMessage_(0, 0),
    };
  }

  var archived = 0;
  var skipped = 0;

  threadIds.forEach(function(threadId) {
    try {
      var thread = GmailApp.getThreadById(threadId);
      if (!thread) {
        skipped++;
        return;
      }
      if (!thread.isInInbox()) {
        skipped++;
        return;
      }
      thread.moveToArchive();
      archived++;
    } catch (e) {
      Logger.log('Archive skip ' + threadId + ': ' + e.message);
      skipped++;
    }
  });

  Logger.log('Archived script-labeled threads: ' + archived + ', skipped: ' + skipped);
  return {
    archived: archived,
    skipped: skipped,
    message: buildArchiveReviewedMessage_(archived, skipped),
  };
}

/** @deprecated Use archiveReviewedCleanupItems */
function archiveReviewedReadCleanupItems() {
  return archiveReviewedCleanupItems();
}

function buildArchiveReviewedMessage_(archived, skipped) {
  if (!archived && !skipped) {
    return 'No script-labeled emails from the last run are waiting in your inbox to archive. Run Classify Read or Cleanup Unread first.';
  }
  if (!archived) {
    return 'Nothing to archive — script-labeled emails from the last run are already out of the inbox.';
  }
  var message = 'Archived ' + archived + ' email' + (archived === 1 ? '' : 's') +
    ' the script labeled during the last run';
  if (skipped) {
    message += ' (' + skipped + ' already archived or not found)';
  }
  return message + '.';
}

function markCleanupRunStart_() {
  PropertiesService.getScriptProperties().setProperty(
    LAST_CLEANUP_RUN_KEY,
    new Date().toISOString()
  );
}

function getLastCleanupRunAt_() {
  var raw = PropertiesService.getScriptProperties().getProperty(LAST_CLEANUP_RUN_KEY);
  if (!raw) return null;
  var dt = new Date(raw);
  return isNaN(dt.getTime()) ? null : dt;
}

function buildCleanupCompleteMessage_(digest) {
  if (!digest || digest.empty || digest.totalCount === 0) {
    return 'All done. No emails needed processing this time.';
  }
  var n = digest.totalCount;
  return 'All done. Processed ' + n + ' email' + (n === 1 ? '' : 's') + ' — see the table below for details.';
}

/** Run once to verify GEMINI_API_KEY works. */
function testGeminiKey() {
  if (!getGeminiApiKey()) {
    Logger.log('Set GEMINI_API_KEY in Script Properties first.');
    return;
  }
  var sampleLabels = getExistingLabels().slice(0, 3);
  try {
    var text = callGemini(
      'Classify this single test email:\n1. From: test@example.com | Subject: Test | Snippet: hello',
      sampleLabels
    );
    Logger.log('OK: ' + text);
  } catch (err) {
    Logger.log('Failed: ' + err.message);
  }
}

/** Sheet only — create/bind dashboard + write sample rows (no Gmail changes). */
function testLogToDashboard() {
  var sampleRows = [
    buildDashboardLogRow_('Test', 'alice@example.com', 'Sample keep row', 'Keep', 'Manual test', null, false),
    buildDashboardLogRow_('Test', 'newsletter@marketing.com', 'Sample delete row', 'CleanupQueue', 'Manual test', null, true),
  ];
  logToDashboard(sampleRows);
  Logger.log('Done. Check the Classification Log tab in your dashboard sheet.');
}

/** Sheet only — ensure dashboard exists; no Gmail, no logging. */
function runSetupDashboardOnly() {
  setupDashboard();
}

// --- Scheduled tasks ---

var SCHEDULE_CONFIG_KEY = 'SCHEDULE_CONFIG';
var SCHEDULE_TRIGGER_HANDLER = 'executeScheduledTask';
var SCHEDULE_WEEK_DAYS = [
  'SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY',
];
var SCHEDULE_TASK_TYPE_ORDER = ['unread', 'read', 'archive'];

function loadSavedScheduleConfig_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(SCHEDULE_CONFIG_KEY);
  if (raw) {
    try {
      return normalizeScheduleConfig_(JSON.parse(raw));
    } catch (e) {
      Logger.log('Invalid SCHEDULE_CONFIG JSON: ' + e.message);
    }
  }
  return migrateLegacySchedule_(props);
}

/**
 * Returns the active schedule for the dashboard and API.
 * @returns {Object}
 */
function getSavedSchedule() {
  var config = loadSavedScheduleConfig_();
  if (!config.configured) {
    return {
      active: false,
      configured: false,
      message: 'No automated schedule active.',
      summary: 'No automated schedule active.',
    };
  }

  return {
    active: true,
    configured: true,
    taskTypes: config.taskTypes,
    taskType: config.taskTypes[0] || '',
    frequency: config.frequency,
    targetDays: config.targetDays,
    targetHour: config.targetHour,
    message: buildScheduleSavedMessage_(config),
    summary: buildScheduleDisplaySummary_(config),
  };
}

function buildScheduleDisplaySummary_(config) {
  var tasks = formatScheduleTaskTypesLabel_(config.taskTypes);
  if (config.frequency === 'hourly') {
    return 'Active: ' + tasks + ' running hourly.';
  }
  if (config.frequency === 'every6hours') {
    return 'Active: ' + tasks + ' running every 6 hours.';
  }

  var days = formatScheduleDaysLabel_(config.targetDays);
  var time = formatScheduleHourLabel_(config.targetHour);
  if (config.frequency === 'weekly') {
    return 'Active: ' + tasks + ' running every ' + days + ' at ' + time + '.';
  }
  return 'Active: ' + tasks + ' running on ' + days + ' at ' + time + '.';
}

function migrateLegacySchedule_(props) {
  var taskType = props.getProperty('SCHEDULE_TASK_TYPE') || '';
  var frequency = props.getProperty('SCHEDULE_FREQUENCY') || '';
  if (!taskType || !frequency) {
    return normalizeScheduleConfig_({});
  }

  var config = {
    taskTypes: normalizeScheduleTaskTypes_({ taskType: taskType }),
    frequency: frequency,
    targetDays: [],
    targetHour: 0,
  };

  if (frequency === 'dailyMidnight') {
    config.frequency = 'custom';
    config.targetDays = SCHEDULE_WEEK_DAYS.slice();
    config.targetHour = 0;
  }

  return normalizeScheduleConfig_(config);
}

function normalizeScheduleConfig_(config) {
  config = config || {};
  var taskTypes = normalizeScheduleTaskTypes_(config);
  var frequency = normalizeScheduleFrequency_(config.frequency) || '';
  var targetDays = normalizeScheduleDays_(config.targetDays);
  var targetHour = normalizeScheduleHour_(config.targetHour);

  return {
    taskTypes: taskTypes,
    frequency: frequency,
    targetDays: targetDays,
    targetHour: targetHour,
    configured: !!(taskTypes.length && frequency),
  };
}

function normalizeScheduleTaskTypes_(config) {
  config = config || {};
  var types = [];
  var list = [];

  if (config.taskTypes) {
    list = Array.isArray(config.taskTypes) ? config.taskTypes : String(config.taskTypes).split(',');
  } else if (config.taskType) {
    list = [config.taskType];
  }

  list.forEach(function(value) {
    var normalized = normalizeScheduleTaskType_(value);
    if (normalized && types.indexOf(normalized) === -1) {
      types.push(normalized);
    }
  });

  types.sort(function(a, b) {
    return SCHEDULE_TASK_TYPE_ORDER.indexOf(a) - SCHEDULE_TASK_TYPE_ORDER.indexOf(b);
  });

  return types;
}

function normalizeScheduleTaskType_(value) {
  var key = String(value || '').toLowerCase().replace(/[\s_-]+/g, '');
  if (key === 'unread' || key === 'classifyunreadonly') return 'unread';
  if (key === 'read' || key === 'classifyreadonly') return 'read';
  if (key === 'archive' || key === 'archiverevieweditems') return 'archive';
  return null;
}

function normalizeScheduleFrequency_(value) {
  var key = String(value || '').toLowerCase().replace(/[\s_-]+/g, '');
  if (key === 'hourly') return 'hourly';
  if (key === 'every6hours') return 'every6hours';
  if (key === 'weekly') return 'weekly';
  if (key === 'custom' || key === 'customday') return 'custom';
  if (key === 'dailymidnight' || key === 'dailyatmidnight') return 'custom';
  return null;
}

function normalizeScheduleDay_(value) {
  var key = String(value || '').toUpperCase().replace(/[\s_-]+/g, '');
  if (key === 'SUN' || key === 'SUNDAY') return 'SUNDAY';
  if (key === 'MON' || key === 'MONDAY') return 'MONDAY';
  if (key === 'TUE' || key === 'TUESDAY') return 'TUESDAY';
  if (key === 'WED' || key === 'WEDNESDAY') return 'WEDNESDAY';
  if (key === 'THU' || key === 'THURSDAY') return 'THURSDAY';
  if (key === 'FRI' || key === 'FRIDAY') return 'FRIDAY';
  if (key === 'SAT' || key === 'SATURDAY') return 'SATURDAY';
  return null;
}

function normalizeScheduleDays_(days) {
  if (!days) return [];
  var list = Array.isArray(days) ? days : String(days).split(',');
  var seen = {};
  var normalized = [];
  list.forEach(function(day) {
    var value = normalizeScheduleDay_(day);
    if (value && !seen[value]) {
      seen[value] = true;
      normalized.push(value);
    }
  });
  return normalized;
}

function normalizeScheduleHour_(value) {
  var hour = parseInt(value, 10);
  if (isNaN(hour) || hour < 0 || hour > 23) {
    return 0;
  }
  return hour;
}

function getScheduleTaskLabel_(taskType) {
  if (taskType === 'unread') return 'Classify Unread Only';
  if (taskType === 'read') return 'Classify Read Only';
  if (taskType === 'archive') return 'Archive Reviewed Items';
  return taskType;
}

function formatScheduleTaskTypesLabel_(taskTypes) {
  if (!taskTypes || !taskTypes.length) return 'No tasks selected';
  return taskTypes.map(getScheduleTaskLabel_).join(' + ');
}

function runScheduledTaskType_(taskType) {
  if (taskType === 'unread') {
    runFullCleanup(false);
    return;
  }
  if (taskType === 'read') {
    classifyReadEmails();
    return;
  }
  if (taskType === 'archive') {
    archiveReviewedCleanupItems();
    return;
  }
  throw new Error('Unknown scheduled task type: ' + taskType);
}

function getScheduleFrequencyLabel_(frequency) {
  if (frequency === 'hourly') return 'Hourly';
  if (frequency === 'every6hours') return 'Every 6 Hours';
  if (frequency === 'weekly') return 'Weekly';
  if (frequency === 'custom') return 'Custom Day';
  return frequency;
}

function formatScheduleHourLabel_(hour) {
  if (hour === 0) return '12:00 AM (Midnight)';
  if (hour === 12) return '12:00 PM (Noon)';
  if (hour < 12) return hour + ':00 AM';
  return (hour - 12) + ':00 PM';
}

function formatScheduleDaysLabel_(targetDays) {
  if (!targetDays || !targetDays.length) return '';
  return targetDays.map(function(day) {
    return day.charAt(0) + day.slice(1, 3).toLowerCase();
  }).join(', ');
}

function buildScheduleSavedMessage_(config) {
  var parts = [
    'Schedule saved: ' + formatScheduleTaskTypesLabel_(config.taskTypes),
    getScheduleFrequencyLabel_(config.frequency),
  ];
  if (config.frequency === 'weekly' || config.frequency === 'custom') {
    parts.push(formatScheduleDaysLabel_(config.targetDays));
    parts.push('at ' + formatScheduleHourLabel_(config.targetHour));
  }
  return parts.filter(function(part) { return part; }).join(' — ') + '.';
}

function persistScheduleConfig_(config) {
  PropertiesService.getScriptProperties().setProperty(
    SCHEDULE_CONFIG_KEY,
    JSON.stringify(config)
  );
}

/**
 * Deletes all project triggers and recreates them from the saved schedule config.
 * @param {Object} config
 */
function rebuildSystemTriggers(config) {
  config = normalizeScheduleConfig_(config);

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  if (config.frequency === 'hourly') {
    ScriptApp.newTrigger(SCHEDULE_TRIGGER_HANDLER).timeBased().everyHours(1).create();
    return;
  }

  if (config.frequency === 'every6hours') {
    ScriptApp.newTrigger(SCHEDULE_TRIGGER_HANDLER).timeBased().everyHours(6).create();
    return;
  }

  if (config.frequency === 'weekly' || config.frequency === 'custom') {
    if (!config.targetDays.length) {
      throw new Error('Select at least one day of the week.');
    }

    config.targetDays.forEach(function(dayName) {
      var weekDay = ScriptApp.WeekDay[dayName];
      if (!weekDay) {
        throw new Error('Invalid weekday: ' + dayName);
      }
      ScriptApp.newTrigger(SCHEDULE_TRIGGER_HANDLER)
        .timeBased()
        .onWeekDay(weekDay)
        .atHour(config.targetHour)
        .create();
    });
    return;
  }

  throw new Error('Unsupported schedule frequency: ' + config.frequency);
}

function scheduleApiResponse_() {
  var saved = getSavedSchedule();
  return {
    active: saved.active,
    configured: saved.configured,
    taskTypes: saved.taskTypes || [],
    taskType: saved.taskType || (saved.taskTypes && saved.taskTypes[0]) || 'unread',
    frequency: saved.frequency || 'weekly',
    targetDays: saved.targetDays || [],
    targetHour: saved.targetHour || 0,
    summary: saved.summary,
    message: saved.message,
  };
}

/**
 * Saves schedule preferences and reinstalls triggers.
 * @param {Object|string} taskTypeOrConfig
 * @param {string=} frequency
 * @param {Array|string=} targetDays
 * @param {number=} targetHour
 */
function saveUserSchedule(taskTypeOrConfig, frequency, targetDays, targetHour) {
  var input = taskTypeOrConfig;
  if (input && typeof input === 'object' && !(input instanceof Date)) {
    input = input;
  } else {
    input = {
      taskType: taskTypeOrConfig,
      frequency: frequency,
      targetDays: targetDays,
      targetHour: targetHour,
    };
  }

  var config = normalizeScheduleConfig_(input);
  if (!config.taskTypes.length) {
    throw new Error('Select at least one task type.');
  }
  if (!config.frequency) {
    throw new Error('Invalid frequency. Use hourly, every6hours, weekly, or custom.');
  }
  if ((config.frequency === 'weekly' || config.frequency === 'custom') && !config.targetDays.length) {
    throw new Error('Select at least one day of the week.');
  }

  persistScheduleConfig_(config);
  rebuildSystemTriggers(config);

  Logger.log('Schedule installed: ' + JSON.stringify(config));
  logSchedulerEvent_('Configured', buildScheduleSavedMessage_(config));

  return {
    ok: true,
    taskTypes: config.taskTypes,
    taskType: config.taskTypes[0] || '',
    frequency: config.frequency,
    targetDays: config.targetDays,
    targetHour: config.targetHour,
    message: buildScheduleSavedMessage_(config),
  };
}

/**
 * Master trigger entry point — runs the task saved in Script Properties.
 */
function executeScheduledTask() {
  var startedAt = Date.now();
  var schedule = loadSavedScheduleConfig_();

  if (!schedule.configured) {
    Logger.log('executeScheduledTask: no schedule configured — skipping.');
    logSchedulerEvent_('Skipped', 'No schedule configured.');
    return;
  }

  Logger.log('executeScheduledTask starting: ' + JSON.stringify(schedule));

  try {
    var completed = [];
    schedule.taskTypes.forEach(function(taskType) {
      runScheduledTaskType_(taskType);
      completed.push(getScheduleTaskLabel_(taskType));
    });

    var elapsedMs = Date.now() - startedAt;
    var detail = completed.join(' → ') + ' finished in ' + elapsedMs + 'ms';
    try {
      var remainingMs = ScriptApp.getRemainingTime();
      if (remainingMs > 0 && remainingMs < 60000) {
        detail += '; low remaining execution time (' + remainingMs + 'ms)';
        logSchedulerEvent_('Warning', detail);
      }
    } catch (ignore) {
      // getRemainingTime unavailable outside execution context
    }
    logSchedulerEvent_('Success', detail);
  } catch (err) {
    var message = err && err.message ? err.message : String(err);
    if (/limit|quota|exceeded|timeout|too many|service invoked/i.test(message)) {
      message = 'Execution limit or trigger failure: ' + message;
    }
    Logger.log('executeScheduledTask failed: ' + message);
    logSchedulerEvent_('Error', message);
  }
}

// --- Web App API ---

function getWebAppApiSecret() {
  return PropertiesService.getScriptProperties().getProperty('WEBAPP_API_SECRET');
}

function parseWebAppPayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing JSON body. POST application/json with { "action": "...", "token": "..." }.');
  }
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error('Invalid JSON payload.');
  }
}

function isAuthorizedWebAppRequest_(payload, e) {
  var secret = getWebAppApiSecret();
  if (!secret) {
    throw new Error('WEBAPP_API_SECRET is not set in Script Properties.');
  }
  var provided = (payload && payload.token) || (e && e.parameter && e.parameter.token) || '';
  return String(provided) === String(secret);
}

function jsonResponse_(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Builds structured digest data from the dashboard log since the last cleanup run.
 * @returns {Object}
 */
function buildDashboardDigestData_() {
  var emptyResult = function(message, periodLabel) {
    return {
      empty: true,
      message: message,
      periodLabel: periodLabel || 'since last cleanup run',
      periodHours: null,
      since: null,
      totalCount: 0,
      rows: [],
      byAction: [],
      bySource: [],
    };
  };

  var lastRunAt = getLastCleanupRunAt_();
  if (!lastRunAt) {
    return emptyResult(
      'Run inbox cleanup first, then come back here to see what changed.',
      'since last cleanup run'
    );
  }

  var ss;
  try {
    ss = ensureDashboardSpreadsheet();
  } catch (e) {
    return emptyResult('Could not open the log spreadsheet. ' + e.message, 'since last cleanup run');
  }

  var sheet = ss.getSheetByName(DASHBOARD_LOG_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) {
    return emptyResult(
      'No activity was logged during the last cleanup run.',
      'since last cleanup run'
    );
  }

  var numRows = sheet.getLastRow() - 1;
  var rawRows = sheet.getRange(2, 1, numRows, DASHBOARD_HEADERS.length).getValues();
  var cutoff = lastRunAt;
  var actionCounts = {};
  var sourceCounts = {};
  var entries = [];

  for (var i = rawRows.length - 1; i >= 0; i--) {
    var row = rawRows[i];
    var ts = row[0];
    if (!(ts instanceof Date)) {
      ts = new Date(ts);
    }
    if (isNaN(ts.getTime()) || ts < cutoff) {
      continue;
    }

    var source = String(row[1] || 'Unknown');
    var sender = String(row[2] || '');
    var subject = String(row[3] || '(no subject)');
    var action = String(row[4] || 'Unknown');
    var reason = String(row[5] || '');

    sourceCounts[source] = (sourceCounts[source] || 0) + 1;
    actionCounts[action] = (actionCounts[action] || 0) + 1;

    entries.push({
      timestamp: ts.toISOString(),
      source: source,
      sender: sender,
      subject: subject,
      action: action,
      reason: reason,
    });
  }

  if (!entries.length) {
    return emptyResult(
      'Nothing was processed during the last cleanup run.',
      'since last cleanup run'
    );
  }

  function toSummary(counts) {
    return Object.keys(counts).sort().map(function(key) {
      return { label: key, count: counts[key] };
    });
  }

  return {
    empty: false,
    message: '',
    periodLabel: 'since last cleanup run',
    periodHours: null,
    since: lastRunAt.toISOString(),
    totalCount: entries.length,
    rows: entries,
    byAction: toSummary(actionCounts),
    bySource: toSummary(sourceCounts),
  };
}

function formatDigestAsText_(digest) {
  if (digest.empty) {
    return digest.message;
  }

  var lines = [
    'Gmail Cleanup digest (' + digest.periodLabel + ')',
    'Total actions: ' + digest.totalCount,
    '',
    'By action:',
  ];

  digest.byAction.forEach(function(item) {
    lines.push('  • ' + item.label + ': ' + item.count);
  });

  lines.push('', 'By source:');
  digest.bySource.forEach(function(item) {
    lines.push('  • ' + item.label + ': ' + item.count);
  });

  return lines.join('\n');
}

/**
 * Summarizes recent dashboard log activity (last 24 hours).
 * @returns {string}
 */
function generateDashboardActionDigest() {
  return formatDigestAsText_(buildDashboardDigestData_());
}

/**
 * Web App entry point (POST only).
 * Body: { "action": "cleanup"|"digest"|"archiveReviewed"|"classifyRead"|"saveSchedule"|"getSchedule", ... }
 *
 * Deploy: Deploy → New deployment → Web app
 *   Execute as: Me
 *   Who has access: Anyone (token still required)
 */
function doPost(e) {
  try {
    var payload = parseWebAppPayload_(e);

    if (!isAuthorizedWebAppRequest_(payload, e)) {
      return jsonResponse_({ ok: false, error: 'Unauthorized' });
    }

    var action = String(payload.action || '').toLowerCase();

    if (action === 'cleanup') {
      var includeRead = payload.includeRead === true ||
        String(payload.includeRead || '').toLowerCase() === 'true';
      runFullCleanup(includeRead);
      var cleanupDigest = buildDashboardDigestData_();
      return jsonResponse_({
        ok: true,
        action: 'cleanup',
        message: buildCleanupCompleteMessage_(cleanupDigest),
        digest: cleanupDigest,
      });
    }

    if (action === 'digest') {
      var digest = buildDashboardDigestData_();
      return jsonResponse_({
        ok: true,
        action: 'digest',
        content: formatDigestAsText_(digest),
        digest: digest,
      });
    }

    if (action === 'archivereviewed') {
      var archiveResult = archiveReviewedCleanupItems();
      return jsonResponse_({
        ok: true,
        action: 'archiveReviewed',
        message: archiveResult.message,
        archived: archiveResult.archived,
      });
    }

    if (action === 'classifyread') {
      classifyReadEmails();
      var readDigest = buildDashboardDigestData_();
      return jsonResponse_({
        ok: true,
        action: 'classifyRead',
        message: buildCleanupCompleteMessage_(readDigest),
        digest: readDigest,
      });
    }

    if (action === 'saveschedule') {
      var scheduleResult = saveUserSchedule({
        taskTypes: payload.taskTypes || payload.taskType,
        taskType: payload.taskType,
        frequency: payload.frequency,
        targetDays: payload.targetDays,
        targetHour: payload.targetHour,
      });
      var savedSchedule = getSavedSchedule();
      return jsonResponse_({
        ok: true,
        action: 'saveSchedule',
        taskTypes: scheduleResult.taskTypes,
        taskType: scheduleResult.taskType,
        frequency: scheduleResult.frequency,
        targetDays: scheduleResult.targetDays,
        targetHour: scheduleResult.targetHour,
        active: savedSchedule.active,
        configured: savedSchedule.configured,
        summary: savedSchedule.summary,
        message: scheduleResult.message,
      });
    }

    if (action === 'getschedule') {
      var schedulePayload = scheduleApiResponse_();
      schedulePayload.ok = true;
      schedulePayload.action = 'getSchedule';
      return jsonResponse_(schedulePayload);
    }

    return jsonResponse_({
      ok: false,
      error: 'Unknown action. Use "cleanup", "digest", "archiveReviewed", "classifyRead", "saveSchedule", or "getSchedule".',
    });
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonResponse_({ ok: false, error: err.message });
  }
}

/** Generate a random API secret — run once, copy output to WEBAPP_API_SECRET. */
function generateWebAppApiSecret() {
  var token = Utilities.getUuid() + Utilities.getUuid();
  Logger.log('Add to Script Properties as WEBAPP_API_SECRET:');
  Logger.log(token);
  return token;
}

// --- Dashboard UI (HtmlService) ---

/**
 * Serves the HTML dashboard. Deploy as Web app; open the deployment URL in a browser.
 * GET shows the UI; POST (doPost) handles API calls.
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile('dashboard');
  template.webAppUrl = ScriptApp.getService().getUrl();
  template.savedScheduleJson = JSON.stringify(getSavedSchedule());
  return template.evaluate()
    .setTitle('Gmail Cleanup')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Called from dashboard.html via google.script.run (token stays server-side). */
function dashboardGetSchedule() {
  var response = scheduleApiResponse_();
  response.ok = true;
  return response;
}

/** Called from dashboard.html via google.script.run (token stays server-side). */
function dashboardSaveSchedule(config) {
  var input = config;
  if (arguments.length > 1 || (input && input.taskTypes === undefined && input.taskType === undefined && typeof input !== 'object')) {
    input = {
      taskTypes: arguments[0],
      frequency: arguments[1],
      targetDays: arguments[2],
      targetHour: arguments[3],
    };
  }
  var result = saveUserSchedule(input);
  var saved = getSavedSchedule();
  return {
    ok: true,
    action: 'saveSchedule',
    taskTypes: result.taskTypes,
    taskType: result.taskType,
    frequency: result.frequency,
    targetDays: result.targetDays,
    targetHour: result.targetHour,
    active: saved.active,
    configured: saved.configured,
    summary: saved.summary,
    message: result.message,
  };
}

/** Called from dashboard.html via google.script.run (token stays server-side). */
function dashboardRunUnreadCleanup() {
  runFullCleanup(false);
  var digest = buildDashboardDigestData_();
  return {
    ok: true,
    action: 'cleanup',
    message: buildCleanupCompleteMessage_(digest),
    digest: digest,
  };
}

/** Called from dashboard.html via google.script.run (token stays server-side). */
function dashboardClassifyRead() {
  classifyReadEmails();
  var digest = buildDashboardDigestData_();
  return {
    ok: true,
    action: 'classifyRead',
    message: buildCleanupCompleteMessage_(digest),
    digest: digest,
  };
}

/** Called from dashboard.html via google.script.run (token stays server-side). */
function dashboardRunCleanup(includeRead) {
  var includeReadEmails = includeRead === true ||
    String(includeRead || '').toLowerCase() === 'true';
  runFullCleanup(includeReadEmails);
  var digest = buildDashboardDigestData_();
  return {
    ok: true,
    message: buildCleanupCompleteMessage_(digest),
    digest: digest,
  };
}

/** Called from dashboard.html via google.script.run (token stays server-side). */
function dashboardArchiveReviewed() {
  var result = archiveReviewedCleanupItems();
  return {
    ok: true,
    action: 'archiveReviewed',
    message: result.message,
    archived: result.archived,
  };
}

/** Called from dashboard.html via google.script.run (token stays server-side). */
function dashboardArchiveReviewedReadItems() {
  return dashboardArchiveReviewed();
}

/** Called from dashboard.html via google.script.run (token stays server-side). */
function dashboardFetchDigest() {
  var digest = buildDashboardDigestData_();
  return {
    ok: true,
    action: 'digest',
    content: formatDigestAsText_(digest),
    digest: digest,
  };
}
