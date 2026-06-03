const Anthropic = require('@anthropic-ai/sdk').default;

// ── Rule-based lists ────────────────────────────────────────────────────────

const DELETE_DOMAINS = [
  'substack.com',
  'beehiiv.com',
  'mailchimp.com',
  'sendgrid.net',
  'constantcontact.com',
  'klaviyo.com',
  'campaigns.amazon.com',
  'marketing.amazon.com',
  'news.linkedin.com',
  'e.linkedin.com',
  'em.linkedin.com',
  'bounce.linkedin.com',
  'notifications.google.com',
  'mailer.robinhood.com',
  'email.robinhood.com',
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
  'cheapsoftwarestocks@substack.com', 'info@e.tikvah.org',
  'no-reply@substack.com', 'support@beenverified-newsletter.com',
  'info.asif.group@send.vpcontact.com', 'books@greenbrothers.co.il',
  'info@delta.co.il', 'Info@alinasleep.co.il', 'info@kapra.co.il',
  'info@kitchenware.co.il', 'info@teimim.com', 'info@tevabari.co.il',
  'info@zap.co.il', 'do-not-reply@pandazzz.co.il', 'itay@ets-hasade.com',
  'info@ortokal.co.il', 'info@marathonisrael.co.il',
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
  'tommy.barav@magical.team',
  'hello@letterboxd.com',
  'yehonatan@ti-space.com',
  'hello@gamma.app',
  'email@mail.pmi.org',
  'newsletter@mail.bubble.io',
];

const KEEP_DOMAINS = [
  'gmail.com',
];

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

const DELETE_SUBJECT_KEYWORDS = [
  // Hebrew
  'פרסומת', 'פרסום', 'מבצע', 'הנחה',
  // English
  'sale', 'unsubscribe', 'newsletter', 'weekly digest',
  'weekly roundup', 'monthly digest', 'special offer',
  'limited time', 'exclusive deal', 'flash sale',
  '% off', 'promo code', 'coupon', 'save big',
  'black friday', 'cyber monday', "you're invited",
  "don't miss out", 'last chance', 'job alert', 'jobs at', 'is hiring',
];

const KEEP_SUBJECT_KEYWORDS = [];

const AI_SNIPPET_CHARS = 800;

// ── Rule lookup indexes (built once at load) ────────────────────────────────

function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function toLowerSet(items) {
  return new Set(items.map(item => item.toLowerCase()));
}

function sortByLengthDesc(items) {
  return items.slice().sort((a, b) => b.length - a.length);
}

function buildDomainIndex(domains) {
  const lowered = domains.map(d => d.toLowerCase());
  return {
    exact: new Set(lowered),
    suffixes: sortByLengthDesc(lowered),
  };
}

function buildSubjectKeywordRegex(keywords) {
  if (!keywords.length) return null;
  const pattern = keywords.map(kw => escapeRegExp(kw)).join('|');
  return new RegExp(pattern, 'i');
}

const DELETE_SENDERS_SET = toLowerSet(DELETE_SENDERS);
const KEEP_SENDERS_SET = toLowerSet(KEEP_SENDERS);
const DELETE_SENDERS_DESC = sortByLengthDesc(DELETE_SENDERS.map(s => s.toLowerCase()));
const KEEP_SENDERS_DESC = sortByLengthDesc(KEEP_SENDERS.map(s => s.toLowerCase()));
const DELETE_DOMAIN_INDEX = buildDomainIndex(DELETE_DOMAINS);
const KEEP_DOMAIN_INDEX = buildDomainIndex(KEEP_DOMAINS);
const DELETE_SUBJECT_REGEX = buildSubjectKeywordRegex(DELETE_SUBJECT_KEYWORDS);
const KEEP_SUBJECT_REGEX = buildSubjectKeywordRegex(KEEP_SUBJECT_KEYWORDS);

function extractEmailAddress(from) {
  const angle = from.match(/<([^>]+)>/);
  if (angle) return angle[1].trim().toLowerCase();
  const bare = from.match(/[\w.+-]+@[\w.-]+/);
  return bare ? bare[0].toLowerCase() : from.toLowerCase().trim();
}

function extractDomain(from) {
  const email = extractEmailAddress(from);
  const at = email.lastIndexOf('@');
  return at === -1 ? '' : email.slice(at + 1);
}

function domainMatchesIndex(domain, index) {
  if (!domain) return false;
  if (index.exact.has(domain)) return true;
  for (const pattern of index.suffixes) {
    const suffix = '.' + pattern;
    if (domain.length > suffix.length && domain.endsWith(suffix)) return true;
  }
  return false;
}

function senderMatches(fromLower, email, exactSet, listDesc) {
  if (email && exactSet.has(email)) return true;
  for (const s of listDesc) {
    if (fromLower.includes(s)) return true;
  }
  return false;
}

function getGmailCategoryNames(labelIds) {
  const names = [];
  for (const id of labelIds || []) {
    if (id === 'CATEGORY_PROMOTIONS') names.push('promotions');
    if (id === 'CATEGORY_SOCIAL') names.push('social');
  }
  return names;
}

function ruleClassify(thread) {
  const { from = '', subject = '', listUnsubscribe, labelIds = [] } = thread;
  const fromLower = from.toLowerCase();
  const email = extractEmailAddress(from);
  const domain = extractDomain(from);
  const subjectLower = subject.toLowerCase();

  if (senderMatches(fromLower, email, KEEP_SENDERS_SET, KEEP_SENDERS_DESC)) {
    return { decision: 'keep', reason: 'matched keep sender', method: 'rule' };
  }
  if (domainMatchesIndex(domain, KEEP_DOMAIN_INDEX)) {
    return { decision: 'keep', reason: `known keep domain: ${domain}`, method: 'rule' };
  }

  if (listUnsubscribe) {
    return { decision: 'delete', reason: 'List-Unsubscribe header', method: 'rule' };
  }

  const categories = getGmailCategoryNames(labelIds);
  if (categories.length) {
    return {
      decision: 'delete',
      reason: `Gmail category: ${categories.join(', ')}`,
      method: 'rule',
    };
  }

  if (senderMatches(fromLower, email, DELETE_SENDERS_SET, DELETE_SENDERS_DESC)) {
    return { decision: 'delete', reason: 'matched delete sender', method: 'rule' };
  }
  if (domainMatchesIndex(domain, DELETE_DOMAIN_INDEX)) {
    return { decision: 'delete', reason: `known delete domain: ${domain}`, method: 'rule' };
  }
  if (DELETE_SUBJECT_REGEX && DELETE_SUBJECT_REGEX.test(subjectLower)) {
    return { decision: 'delete', reason: 'delete keyword in subject', method: 'rule' };
  }
  if (KEEP_SUBJECT_REGEX && KEEP_SUBJECT_REGEX.test(subjectLower)) {
    return { decision: 'keep', reason: 'keep keyword in subject', method: 'rule' };
  }
  return null;
}

// ── AI classification (batched) ─────────────────────────────────────────────

const { GoogleGenAI } = require('@google/genai');

const SYSTEM_PROMPT = `You are an email triage assistant. You receive batches of messages with metadata and body snippets.

Analyze each batch on its own: infer who sends bulk mail, which items are newsletters or marketing, which are direct personal or operational mail, and which automated notices look actionable.

Do not assume any fixed job, industry, or inbox profile — reason only from batch evidence.

Keep emails that appear individually important: direct human correspondence, actionable requests, reservations or receipts you may need, security or account alerts, time-sensitive coordination.

Delete emails that are junk or low-value bulk: promotional blasts, newsletters, drip campaigns, marketing automation, social or network digests, retailer deals, product-update spam, and non-actionable bulk notifications. Be decisive when metadata or copy signals mass mail (List-Unsubscribe, List-Id, marketing footers, percent-off language, etc.).

When uncertain, prefer keep only if importance is clear from the content.

Respond with ONLY a JSON array — no markdown, no explanation outside the JSON:
[{"index": 1, "decision": "keep", "reason": "one short phrase"}, ...]`;

function formatEmailBatch(threads) {
  return threads.map((t, i) => {
    const parts = [
      `${i + 1}. From: ${t.from || '(unknown)'}`,
      `Subject: ${t.subject || '(no subject)'}`,
    ];
    if (t.listUnsubscribe) parts.push('List-Unsubscribe: yes');
    if (t.listId) parts.push('List-Id: present');
    if (t.precedence) parts.push(`Precedence: ${t.precedence}`);
    const categories = getGmailCategoryNames(t.labelIds);
    if (categories.length) parts.push(`Gmail categories: ${categories.join(', ')}`);
    parts.push(`Snippet: ${(t.snippet || '(empty)').slice(0, AI_SNIPPET_CHARS)}`);
    return parts.join(' | ');
  });
}

function parseAiDecisions(raw) {
  const match = raw.match(/\[[\s\S]*\]/);
  if (!match) throw new Error(`AI returned no JSON array. Raw: ${raw.slice(0, 200)}`);

  const decisions = JSON.parse(match[0]);
  return decisions.map(d => ({
    index: d.index - 1,
    decision: d.decision === 'delete' ? 'delete' : 'keep',
    reason: String(d.reason || 'AI decision'),
  }));
}

let anthropicClient;
function getAnthropic() {
  if (!anthropicClient) anthropicClient = new Anthropic();
  return anthropicClient;
}

let geminiClient;
function getGemini() {
  if (!geminiClient) {
    geminiClient = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
  }
  return geminiClient;
}

async function anthropicBatchClassify(threads) {
  const items = formatEmailBatch(threads);
  const message = await getAnthropic().messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 4096,
    system: [
      { type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } },
    ],
    messages: [
      { role: 'user', content: `Emails to classify:\n${items.join('\n')}` },
    ],
  });

  return parseAiDecisions(message.content[0].text.trim());
}

async function geminiBatchClassify(threads) {
  const items = formatEmailBatch(threads);
  const response = await getGemini().models.generateContent({
    model: process.env.GEMINI_MODEL || 'gemini-2.5-flash',
    contents: `Emails to classify:\n${items.join('\n')}`,
    config: {
      systemInstruction: SYSTEM_PROMPT,
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            index: { type: 'INTEGER' },
            decision: { type: 'STRING', enum: ['keep', 'delete'] },
            reason: { type: 'STRING' },
          },
          required: ['index', 'decision', 'reason'],
        },
      },
    },
  });

  return parseAiDecisions(response.text.trim());
}

function getAiProvider() {
  return (process.env.AI_PROVIDER || 'gemini').toLowerCase();
}

async function aiBatchClassify(threads) {
  if (!threads.length) return [];

  const provider = getAiProvider();
  if (provider === 'anthropic') return anthropicBatchClassify(threads);
  if (provider === 'gemini') return geminiBatchClassify(threads);
  throw new Error(`Unknown AI_PROVIDER "${provider}". Use "gemini" or "anthropic".`);
}

module.exports = {
  ruleClassify,
  aiBatchClassify,
  AI_SNIPPET_CHARS,
  DELETE_DOMAINS,
  KEEP_DOMAINS,
  DELETE_SENDERS,
  KEEP_SENDERS,
};
