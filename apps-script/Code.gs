// Gmail Cleanup Agent v4 (Gemini multi-label)
// Pass 1: rule-based (classifyInbox)
// Pass 2: AI assigns your existing Gmail labels (classifyUnclassified)
// Run order: resetLabel -> setupLabel -> classifyInbox -> classifyUnclassified
//
// Script properties:
//   GEMINI_API_KEY              — https://aistudio.google.com/apikey
//   GEMINI_MODEL                — optional, default gemini-2.5-flash
//   DASHBOARD_SPREADSHEET_ID    — optional; auto-created on first log if unset

const LABEL_NAME = 'CleanupQueue';
const SPECIAL_KEEP = 'Keep';
const SPECIAL_DELETE = 'Delete';
const BATCH_SIZE = 200;
const AI_BATCH_SIZE = 50;
const AI_SNIPPET_CHARS = 800;
const DEFAULT_GEMINI_MODEL = 'gemini-2.5-flash';
const MAX_SCHEMA_ENUM_LABELS = 128;

// --- Dashboard logging (must be in this file for classifyInbox / classifyUnclassified) ---

const DASHBOARD_LOG_SHEET = 'Classification Log';
const DASHBOARD_SPREADSHEET_TITLE = 'Gmail Cleanup Dashboard';
const DASHBOARD_HEADERS = ['Timestamp', 'Source', 'Sender', 'Subject', 'Action', 'Reason'];

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
  var sheet = ss.getSheetByName(DASHBOARD_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_LOG_SHEET);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(DASHBOARD_HEADERS);
    sheet.setFrozenRows(1);
  }
  Logger.log('Dashboard ready: ' + ss.getUrl());
  Logger.log('Log tab: "' + DASHBOARD_LOG_SHEET + '"');
  return ss.getUrl();
}

function getDashboardLogSheet(ss) {
  var sheet = ss.getSheetByName(DASHBOARD_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_LOG_SHEET);
    sheet.appendRow(DASHBOARD_HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(DASHBOARD_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Appends classification rows to the dashboard sheet.
 * Each entry: [Date, source, sender, subject, action, reason]
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
    ];
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, DASHBOARD_HEADERS.length).setValues(rows);
  Logger.log('Dashboard: appended ' + rows.length + ' row(s) → ' + ss.getUrl());
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
 * @returns {'keep'|'delete'|null}
 */
function classifyByRules(sender, subject) {
  var fromLower = sender.toLowerCase();
  var email = extractEmailAddress(sender);
  var domain = extractDomain(sender);
  var subjectLower = subject.toLowerCase();

  if (senderMatches(fromLower, email, KEEP_SENDERS_SET, KEEP_SENDERS_DESC)) {
    return 'keep';
  }
  if (domainMatchesIndex(domain, KEEP_DOMAIN_INDEX)) {
    return 'keep';
  }
  if (senderMatches(fromLower, email, DELETE_SENDERS_SET, DELETE_SENDERS_DESC)) {
    return 'delete';
  }
  if (domainMatchesIndex(domain, DELETE_DOMAIN_INDEX)) {
    return 'delete';
  }
  if (DELETE_SUBJECT_REGEX && DELETE_SUBJECT_REGEX.test(subjectLower)) {
    return 'delete';
  }
  return null;
}

function ruleShouldDelete(sender, subject) {
  return classifyByRules(sender, subject) === 'delete';
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
    'You are an email triage assistant. Classify mail for someone running a startup, doing software development, and managing personal life.\n' +
    'For each email, assign exactly one outcome: a user Gmail label from the list below, or a special value.\n\n' +
    'USER GMAIL LABELS (use the exact string when it fits):\n' +
    labelLines + '\n\n' +
    'SPECIAL VALUES:\n' +
    '- "' + SPECIAL_KEEP + '": Keep in inbox when no label fits — messages from real people; startup operations (customers, partners, hiring, finance, legal); software/dev work (code, CI/CD, cloud, APIs, vendors, security); personal life (family, health, home, travel, accounts); receipts and transactional mail you may need later.\n' +
    '- "' + SPECIAL_DELETE + '": Aggressively delete junk — marketing, newsletters, promotions, retail deals, product updates, drip campaigns, event invites from brands, LinkedIn/social digests, SaaS nurture mail, job-alert spam, and any automated notification that is not actionable.\n\n' +
    'Rules: Prefer ' + SPECIAL_DELETE + ' for clear marketing and bulk mail. When uncertain between a custom label and ' + SPECIAL_KEEP + ', prefer ' + SPECIAL_KEEP + ' only if the message looks personally or operationally important.\n' +
    'Never invent label names — only values from the list above or the two special values.\n\n' +
    'Respond with ONLY a JSON array — no markdown:\n' +
    '[{"index": 1, "suggestedLabel": "LabelName", "reason": "one short phrase"}, ...]'
  );
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

function classifyInbox() {
  var logEntries = [];
  var label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    Logger.log('Run setupLabel() first!');
    return;
  }

  var threads = GmailApp.search('is:unread in:inbox -label:CleanupQueue', 0, BATCH_SIZE);
  var tagged = 0, kept = 0, unclassified = 0;

  threads.forEach(function(thread) {
    var msg = thread.getMessages()[0];
    var sender = msg.getFrom();
    var subject = msg.getSubject();
    var verdict = classifyByRules(sender, subject);

    if (verdict === 'keep') {
      logEntries.push([new Date(), 'Rule-Based', sender, subject, 'Keep', 'Matched Rule']);
      kept++;
      return;
    }

    if (verdict === 'delete') {
      thread.addLabel(label);
      thread.moveToArchive();
      logEntries.push([new Date(), 'Rule-Based', sender, subject, LABEL_NAME, 'Matched Rule']);
      Logger.log('-> CleanupQueue: ' + subject);
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

function classifyUnclassified() {
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

  var threads = GmailApp.search('is:unread in:inbox -label:CleanupQueue', 0, AI_BATCH_SIZE);
  if (threads.length === 0) {
    Logger.log('No unclassified emails remaining!');
    logToDashboard(logEntries);
    return;
  }

  var items = threads.map(function(thread, i) {
    var msg = thread.getMessages()[0];
    return (i + 1) + '. From: ' + msg.getFrom() +
      ' | Subject: ' + msg.getSubject() +
      ' | Snippet: ' + (msg.getPlainBody() || '').slice(0, AI_SNIPPET_CHARS);
  });

  var userPrompt = 'Emails to classify:\n' + items.join('\n');

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

    var msg = thread.getMessages()[0];
    var sender = msg.getFrom();
    var subject = msg.getSubject();
    var rawLabel = d.suggestedLabel !== undefined ? d.suggestedLabel : d.decision;
    var decision = d.decision !== undefined ? d.decision : rawLabel;
    var resolved = resolveSuggestedLabel(rawLabel, existingLabels);
    var reason = d.reason || '';

    logEntries.push([new Date(), 'Gemini AI', sender, subject, decision, reason]);

    if (resolved === SPECIAL_DELETE) {
      thread.addLabel(cleanupLabel);
      thread.moveToArchive();
      Logger.log('-> ' + LABEL_NAME + ' [Delete]: ' + subject + ' — ' + reason);
      deleted++;
      return;
    }

    if (resolved === SPECIAL_KEEP) {
      Logger.log('Keep [AI]: ' + subject + ' — ' + reason);
      kept++;
      return;
    }

    var entry = labelLookup[resolved.toLowerCase()];
    if (entry) {
      thread.addLabel(entry.label);
      Logger.log('-> Label "' + entry.name + '": ' + subject + ' — ' + reason);
      labeled++;
      return;
    }

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

/** Log all custom labels available to the classifier. */
function listExistingLabels() {
  var labels = getExistingLabels();
  Logger.log('Custom labels (' + labels.length + ', excluding ' + LABEL_NAME + '):');
  labels.forEach(function(name) { Logger.log('  - ' + name); });
}

function runFullCleanup() {
  setupLabel();
  classifyInbox();
  classifyUnclassified();
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
    [new Date(), 'Test', 'alice@example.com', 'Sample keep row', 'Keep', 'Manual test'],
    [new Date(), 'Test', 'newsletter@marketing.com', 'Sample delete row', 'CleanupQueue', 'Manual test'],
  ];
  logToDashboard(sampleRows);
  Logger.log('Done. Check the Classification Log tab in your dashboard sheet.');
}

/** Sheet only — ensure dashboard exists; no Gmail, no logging. */
function runSetupDashboardOnly() {
  setupDashboard();
}
