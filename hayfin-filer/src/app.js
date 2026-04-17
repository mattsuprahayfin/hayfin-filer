// ─── HAYFIN INBOX FILER — OUTLOOK ADD-IN ────────────────────────────────────
// Reads the selected email or inbox, gets AI folder suggestions,
// and moves emails via the Office.js API.

const STORAGE_KEY_APIKEY   = 'hf_filer_apikey';
const STORAGE_KEY_LEARNING = 'hf_filer_learning';
const MAX_LEARNING_ENTRIES = 30;

// ─── STATE ───────────────────────────────────────────────────────────────────
let apiKey = null;
let pendingItems  = [];  // { id, subject, sender, conversationId, internetMessageId, emailIds }
let stats = { total: 0, moved: 0, skipped: 0 };

// ─── INIT ────────────────────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    apiKey = localStorage.getItem(STORAGE_KEY_APIKEY);
    if (!apiKey) {
      showSetup();
    } else {
      showReady();
    }
  } else {
    document.getElementById('init-status').textContent = 'Open in Outlook to use this add-in.';
  }
});

// ─── SETUP PANEL ─────────────────────────────────────────────────────────────
function showSetup() {
  const main = document.getElementById('main-content');
  main.innerHTML = `
    <div class="setup-panel">
      <h3>API Key Required</h3>
      <p>Enter your Anthropic API key to enable AI filing suggestions. The key is stored locally in this browser and never sent anywhere except directly to Anthropic.</p>
      <input type="password" id="api-key-input" placeholder="sk-ant-..." autocomplete="off" />
      <button class="btn-save" onclick="saveApiKey()">Save & Continue</button>
      <p class="note">Get a key at console.anthropic.com → API Keys. Claude claude-sonnet-4-20250514 is used for suggestions.</p>
    </div>`;
  document.getElementById('btn-load').disabled = true;
  document.getElementById('stats').textContent = '—';
}

function saveApiKey() {
  const input = document.getElementById('api-key-input');
  const key = (input?.value || '').trim();
  if (!key.startsWith('sk-ant-')) {
    input.style.outline = '1px solid var(--red)';
    return;
  }
  localStorage.setItem(STORAGE_KEY_APIKEY, key);
  apiKey = key;
  showReady();
}

// ─── READY STATE ─────────────────────────────────────────────────────────────
function showReady() {
  const main = document.getElementById('main-content');

  // Check if we have a selected item
  const item = Office.context.mailbox?.item;
  if (item) {
    main.innerHTML = `<div class="status">Selected email detected. Press <b>Load inbox</b> to load all unread, or <b>File this</b> for just the open email.</div>`;
  } else {
    main.innerHTML = `<div class="status">Press <b>Load inbox</b> to fetch your inbox and get filing suggestions.</div>`;
  }

  const btn = document.getElementById('btn-load');
  btn.disabled = false;
  btn.onclick = loadInbox;
  updateStats();
}

// ─── LOAD INBOX ──────────────────────────────────────────────────────────────
async function loadInbox() {
  const btn = document.getElementById('btn-load');
  btn.disabled = true;
  btn.textContent = 'Loading…';

  const main = document.getElementById('main-content');
  main.innerHTML = `<div class="status"><span class="pulse"></span><span>Fetching inbox…</span></div>`;

  try {
    const messages = await fetchInboxMessages();
    if (!messages.length) {
      main.innerHTML = `<div class="status">Inbox is empty or no messages found.</div>`;
      btn.disabled = false;
      btn.textContent = 'Load inbox';
      return;
    }

    // Group by conversation
    pendingItems = groupByConversation(messages);
    stats = { total: pendingItems.length, moved: 0, skipped: 0 };
    updateStats();

    // Render cards
    main.innerHTML = '';
    for (const item of pendingItems) {
      main.appendChild(createCard(item));
    }

    // Kick off AI suggestions (async, non-blocking)
    for (const item of pendingItems) {
      getSuggestion(item);
    }

    btn.textContent = 'Refresh';
    btn.disabled = false;
    btn.onclick = loadInbox;

  } catch (err) {
    main.innerHTML = `<div class="error">Error loading inbox: ${err.message}</div>`;
    btn.disabled = false;
    btn.textContent = 'Load inbox';
  }
}

// ─── FETCH INBOX VIA EWS / REST ───────────────────────────────────────────────
function fetchInboxMessages() {
  return new Promise((resolve, reject) => {
    // Use the Office.js mailbox to get inbox items
    // We use makeEwsRequestAsync for broad compatibility
    const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="message:From"/>
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
          <t:FieldURI FieldURI="message:ConversationId"/>
          <t:FieldURI FieldURI="message:InternetMessageId"/>
          <t:FieldURI FieldURI="message:IsRead"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="50" Offset="0" BasePoint="Beginning"/>
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }

      try {
        const parser = new DOMParser();
        const xml = parser.parseFromString(result.value, 'text/xml');
        const items = xml.querySelectorAll('Message');
        const messages = [];

        items.forEach(item => {
          const id = item.querySelector('ItemId')?.getAttribute('Id');
          const changeKey = item.querySelector('ItemId')?.getAttribute('ChangeKey');
          const subject = item.querySelector('Subject')?.textContent || '(no subject)';
          const fromName = item.querySelector('Name')?.textContent || '';
          const fromEmail = item.querySelector('EmailAddress')?.textContent || '';
          const received = item.querySelector('DateTimeReceived')?.textContent || '';
          const convId = item.querySelector('ConversationId')?.getAttribute('Id') || id;
          const msgId = item.querySelector('InternetMessageId')?.textContent || '';
          const isRead = item.querySelector('IsRead')?.textContent === 'true';

          if (id) {
            messages.push({ id, changeKey, subject, fromName, fromEmail, received, convId, msgId, isRead });
          }
        });

        resolve(messages);
      } catch (e) {
        reject(new Error('Failed to parse inbox response: ' + e.message));
      }
    });
  });
}

// ─── GROUP INTO CONVERSATIONS ────────────────────────────────────────────────
function groupByConversation(messages) {
  const map = new Map();
  for (const msg of messages) {
    const key = msg.convId || normaliseSubject(msg.subject);
    if (!map.has(key)) {
      map.set(key, {
        convId: key,
        subject: stripRePrefixes(msg.subject),
        senders: new Set(),
        latestDate: msg.received,
        emailIds: [],
        changeKeys: {},
        state: 'pending',   // pending | suggesting | done | skipped | error
        suggestion: null,
      });
    }
    const thread = map.get(key);
    thread.emailIds.push(msg.id);
    thread.changeKeys[msg.id] = msg.changeKey;
    if (msg.fromEmail) thread.senders.add(msg.fromEmail.toLowerCase());
    if (msg.fromName)  thread.senders.add(msg.fromName);
    if (msg.received > thread.latestDate) {
      thread.latestDate = msg.received;
      thread.subject = stripRePrefixes(msg.subject);
    }
  }
  return [...map.values()];
}

function stripRePrefixes(s) {
  return s.replace(/^(re|fw|fwd|aw|wg):\s*/i, '').trim();
}

function normaliseSubject(s) {
  return stripRePrefixes(s).toLowerCase().replace(/\s+/g, ' ');
}

// ─── CARD RENDERING ──────────────────────────────────────────────────────────
function createCard(item) {
  const card = document.createElement('div');
  card.className = 'card';
  card.id = `card-${item.convId}`;

  const senderStr = [...item.senders].slice(0, 2).join(', ');
  const dateStr = item.latestDate ? new Date(item.latestDate).toLocaleDateString('en-GB', { day:'numeric', month:'short' }) : '';

  card.innerHTML = `
    <div class="card-header">
      <div class="card-subject" title="${escHtml(item.subject)}">${escHtml(item.subject)}</div>
      <div class="card-meta">
        <span class="sender" title="${escHtml(senderStr)}">${escHtml(senderStr)}</span>
        <span>${dateStr}</span>
        <span>${item.emailIds.length} email${item.emailIds.length !== 1 ? 's' : ''}</span>
      </div>
    </div>
    <div class="suggestion loading" id="sug-${item.convId}">
      <span class="pulse"></span>&nbsp;Getting suggestion…
    </div>
    <div class="actions" id="act-${item.convId}" style="display:none"></div>`;

  return card;
}

function renderSuggestion(item) {
  const sugEl = document.getElementById(`sug-${item.convId}`);
  const actEl = document.getElementById(`act-${item.convId}`);
  if (!sugEl || !actEl) return;

  if (item.state === 'done') {
    sugEl.className = 'done-badge';
    sugEl.innerHTML = `✓ Moved to <b>${escHtml(item.movedTo)}</b>`;
    actEl.style.display = 'none';
    document.getElementById(`card-${item.convId}`)?.classList.add('done');
    return;
  }

  if (item.state === 'skipped') {
    sugEl.className = 'skipped-badge';
    sugEl.textContent = '— Skipped';
    actEl.style.display = 'none';
    document.getElementById(`card-${item.convId}`)?.classList.add('done');
    return;
  }

  if (item.state === 'error') {
    sugEl.className = 'suggestion';
    sugEl.innerHTML = `<span style="color:var(--red);font-size:11px">${escHtml(item.errorMsg || 'Error')}</span>`;
    actEl.style.display = 'none';
    return;
  }

  if (!item.suggestion) {
    sugEl.className = 'suggestion loading';
    sugEl.innerHTML = `<span class="pulse"></span>&nbsp;Getting suggestion…`;
    actEl.style.display = 'none';
    return;
  }

  // Has suggestion
  const { folder, reason } = item.suggestion;
  sugEl.className = 'suggestion';
  sugEl.innerHTML = `
    <span class="folder-badge" title="${escHtml(folder)}">${escHtml(folder)}</span>
    <span class="reason" title="${escHtml(reason)}">${escHtml(reason)}</span>`;

  actEl.style.display = 'flex';
  actEl.innerHTML = `
    <button class="btn btn-confirm" onclick="confirmMove('${escAttr(item.convId)}')">Move ✓</button>
    <button class="btn btn-skip" onclick="skipItem('${escAttr(item.convId)}')">Skip</button>
    <select class="folder-select" id="sel-${escAttr(item.convId)}" onchange="changeSuggestion('${escAttr(item.convId)}', this.value)">
      ${SUGGESTABLE_FOLDERS.map(f =>
        `<option value="${escAttr(f.name)}" ${f.name === folder ? 'selected' : ''}>${escHtml(f.name)}</option>`
      ).join('')}
    </select>`;
}

// ─── AI SUGGESTION ───────────────────────────────────────────────────────────
async function getSuggestion(item) {
  try {
    const learningRaw = localStorage.getItem(STORAGE_KEY_LEARNING) || '[]';
    const learning = JSON.parse(learningRaw).slice(-15);
    const learningContext = learning.length
      ? '\n\nPast corrections:\n' + learning.map(l => `- "${l.subject}" → ${l.folder}`).join('\n')
      : '';

    const folderList = SUGGESTABLE_FOLDERS.map(f => f.name).join(', ');
    const senderList = [...item.senders].join(', ');

    const systemPrompt = `You are an email filing assistant for Matthew Supranowicz at Hayfin Capital Management.

Filing rules:
- "1. Admin" → Internal HR/ops: compliance, benefits, tools, office, IR@hayfin.com, commsteam@hayfin.com
- "1. Admin / Expenses" → donotreply@notification.gb.webexpenses.com
- "1. Admin / GWI" → GWI internal emails
- "1. Admin / Training" → Training, e-learning, donotreply@traliant.com
- "1. Admin / Arctic Project" → Arctic Project emails
- "2. Personal" → Non-work personal from outside hayfin.com
- "3. PSG" → General PSG team, LP relations, capital markets, product strategy
- "3. PSG / APAC" → APAC investor relations, Japan/Australia roadshows
- "3. PSG / Co-investments" → Co-investment deal threads (DermCare, Stradbroke co-invest)
- "3. PSG / DealCloud" → DealCloud system notifications, pre-tracker
- "3. PSG / DLF" → DLF II/III/IV/V threads, DLF cashflow, rated note feeders, DLF liquidity
- "3. PSG / ESG" → ESG reporting, LP sustainability queries, ESG data collection
- "3. PSG / HOF" → Healthcare Opportunities Fund
- "3. PSG / HTS-SOF" → HTS/SOF threads, HTS/SOF liquidity meetings
- "3. PSG / HYSL" → High Yield and Syndicated Loans marketing
- "3. PSG / Interval Fund" → Interval Fund / DLF Evergreen
- "3. PSG / LGPS" → Local Government Pension Scheme
- "3. PSG / Maritime" → Hayfin Maritime Yield Fund, ASRS maritime, maritime risk
- "3. PSG / PES" → PES / PES III fund emails
- "4. Multi-strat SMAs" → General multi-strat SMA, PM dashboard, G2N reporting
- "4. Multi-strat SMAs / ART" → Australian Retirement Trust (ART/Stradbroke), CPS230, side letter, ART co-invest
- "4. Multi-strat SMAs / Big Cypress" → Big Cypress SMA
- "4. Multi-strat SMAs / Chief Illinois" → Chief Illinois SMA
- "4. Multi-strat SMAs / Future Fund" → Future Fund Australia, Future Growth Capital
- "4. Multi-strat SMAs / Hostplus" → HOSTPLUS super fund
- "4. Multi-strat SMAs / HSBC" → HSBC SMA
- "4. Multi-strat SMAs / Other OZ Accounts" → Other Australian accounts (LGT Crestone, IFM, TRSIL, TPT)
- "4. Multi-strat SMAs / OTPP" → Ontario Teachers' Pension Plan
- "4. Multi-strat SMAs / QIC" → QIC, JANA QIC emails
- "4. Multi-strat SMAs / REST" → REST Super, REST geopolitical analysis
- "5. Research team" → Internal research: alexandra.fischer, john.macgreevy, aubain.noel
- "6. External research" → External providers: Bloomberg, JPM, Barclays, BNPP, GS, DB, Citi, BofA, newsletters, market updates
- "6. External research / Octus" → Octus (formerly Reorg) research
- "6. External research / Reading" → Reading material saved for later${learningContext}

Respond ONLY with JSON: {"folder": "exact folder name", "reason": "max 8 words"}
Folder must exactly match one of: ${folderList}`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 120,
        system: systemPrompt,
        messages: [{
          role: 'user',
          content: `Subject: ${item.subject}\nSenders: ${senderList}\nEmails in thread: ${item.emailIds.length}`,
        }],
      }),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      if (response.status === 401) {
        // Bad API key — show setup
        localStorage.removeItem(STORAGE_KEY_APIKEY);
        apiKey = null;
        showSetup();
        return;
      }
      throw new Error(err?.error?.message || `API error ${response.status}`);
    }

    const data = await response.json();
    const text = data.content?.[0]?.text || '';
    const clean = text.replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(clean);

    // Validate folder name
    const matchedFolder = SUGGESTABLE_FOLDERS.find(f => f.name === parsed.folder);
    if (!matchedFolder) {
      // Try fuzzy match
      const fuzzy = SUGGESTABLE_FOLDERS.find(f =>
        f.name.toLowerCase().includes(parsed.folder.toLowerCase()) ||
        parsed.folder.toLowerCase().includes(f.name.toLowerCase())
      );
      parsed.folder = fuzzy ? fuzzy.name : SUGGESTABLE_FOLDERS[0].name;
    }

    item.suggestion = { folder: parsed.folder, reason: parsed.reason || '' };
    renderSuggestion(item);

  } catch (err) {
    item.state = 'error';
    item.errorMsg = err.message;
    renderSuggestion(item);
  }
}

// ─── CONFIRM MOVE ────────────────────────────────────────────────────────────
async function confirmMove(convId) {
  const item = pendingItems.find(i => i.convId === convId);
  if (!item || !item.suggestion) return;

  const btn = document.querySelector(`#act-${convId} .btn-confirm`);
  if (btn) { btn.disabled = true; btn.textContent = 'Moving…'; }

  const folderName = document.getElementById(`sel-${convId}`)?.value || item.suggestion.folder;
  const folder = KNOWN_FOLDERS.find(f => f.name === folderName);
  if (!folder) {
    item.state = 'error';
    item.errorMsg = 'Folder not found: ' + folderName;
    renderSuggestion(item);
    return;
  }

  try {
    await moveEmails(item.emailIds, item.changeKeys, folder.id);
    item.state = 'done';
    item.movedTo = folderName;
    stats.moved++;
    stats.total = Math.max(stats.total, stats.moved + stats.skipped);
    updateStats();
    logLearning(item.subject, folderName);
    renderSuggestion(item);
  } catch (err) {
    item.state = 'error';
    item.errorMsg = 'Move failed: ' + err.message;
    if (btn) { btn.disabled = false; btn.textContent = 'Move ✓'; }
    renderSuggestion(item);
  }
}

// ─── MOVE VIA EWS ────────────────────────────────────────────────────────────
function moveEmails(emailIds, changeKeys, targetFolderId) {
  return new Promise((resolve, reject) => {
    const itemIds = emailIds.map(id =>
      `<t:ItemId Id="${escXml(id)}" ChangeKey="${escXml(changeKeys[id] || '')}"/>`
    ).join('');

    const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:MoveItem>
      <m:ToFolderId>
        <t:FolderId Id="${escXml(targetFolderId)}"/>
      </m:ToFolderId>
      <m:ItemIds>${itemIds}</m:ItemIds>
    </m:MoveItem>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }

      const parser = new DOMParser();
      const xml = parser.parseFromString(result.value, 'text/xml');
      const responseClass = xml.querySelector('MoveItemResponseMessage')?.getAttribute('ResponseClass');

      if (responseClass === 'Success' || responseClass === 'Warning') {
        resolve();
      } else {
        const msg = xml.querySelector('MessageText')?.textContent || 'Unknown error';
        reject(new Error(msg));
      }
    });
  });
}

// ─── SKIP ─────────────────────────────────────────────────────────────────────
function skipItem(convId) {
  const item = pendingItems.find(i => i.convId === convId);
  if (!item) return;
  item.state = 'skipped';
  stats.skipped++;
  updateStats();
  renderSuggestion(item);
}

// ─── CHANGE SUGGESTION ───────────────────────────────────────────────────────
function changeSuggestion(convId, folderName) {
  const item = pendingItems.find(i => i.convId === convId);
  if (!item || !item.suggestion) return;
  item.suggestion.folder = folderName;
  item.suggestion.reason = 'manually selected';
  // Don't re-render fully — just update the badge text
  const sugEl = document.getElementById(`sug-${convId}`);
  if (sugEl) {
    const badge = sugEl.querySelector('.folder-badge');
    if (badge) { badge.textContent = folderName; badge.title = folderName; }
    const reason = sugEl.querySelector('.reason');
    if (reason) { reason.textContent = 'manually selected'; }
  }
}

// ─── LEARNING LOG ────────────────────────────────────────────────────────────
function logLearning(subject, folder) {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_LEARNING) || '[]';
    const log = JSON.parse(raw);
    log.push({ subject: stripRePrefixes(subject), folder, ts: Date.now() });
    while (log.length > MAX_LEARNING_ENTRIES) log.shift();
    localStorage.setItem(STORAGE_KEY_LEARNING, JSON.stringify(log));
  } catch (_) {}
}

// ─── STATS ───────────────────────────────────────────────────────────────────
function updateStats() {
  const el = document.getElementById('stats');
  if (!el) return;
  const pending = pendingItems.filter(i => i.state === 'pending' || i.state === 'suggesting').length;
  el.innerHTML = `<b>${stats.moved}</b> moved · <b>${stats.skipped}</b> skipped · <b>${pending}</b> pending`;
}

// ─── UTILS ───────────────────────────────────────────────────────────────────
function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function escAttr(s) {
  return String(s).replace(/'/g,"\\'").replace(/"/g,'&quot;');
}
function escXml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
