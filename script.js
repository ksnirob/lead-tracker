const DEFAULT_WORKBOOK_PATHS = [
  'excel/Batch_01_Leads_Tracker.xlsx',
  'excel/Batch_02_Leads_Tracker.xlsx',
  'excel/Batch_03_Leads_Tracker.xlsx',
  'excel/Batch_04_Leads_Tracker.xlsx',
  'excel/Batch_05_Leads_Tracker.xlsx'
];
const FALLBACK_WORKBOOK_NAME = 'leads-tracker.xlsx';
const HANDLE_DB = 'lead-tracker-file-handles';
const HANDLE_STORE = 'handles';
const HANDLE_KEY = 'primary-workbook';

const STATUS_OPTIONS = [
  'Draft',
  'Sent',
  'Replied',
  'Mail Not Found',
  'Bounce Back',
  'Closed'
];

const PIPELINE_STATUSES = [
  'Draft',
  'Sent',
  'Replied',
  'Win',
  'Lost'
];

const LEAD_STATUS_OPTIONS = [
  'Win',
  'Lost'
];

const REQUIRED_HEADERS = [
  'Name',
  'Website',
  'Email',
  'Phone',
  'Social Links',
  '1st Email Subject',
  '1st Email',
  'Status',
  'Dates',
  'Follow Up 1 Subject',
  'Follow Up 1',
  'Status',
  'Dates',
  'Follow Up 2 Subject',
  'Follow Up 2',
  'Status',
  'Dates',
  'Lead Status'
];

const HEADER_ALIASES = {
  companyName: ['Name', 'Company Name', 'Company', 'CompanyName'],
  website: ['Website', 'Site', 'URL'],
  email: ['Email', 'Email Address'],
  contactNumber: ['Phone', 'Contact Number', 'Contact', 'Phone Number'],
  socialMediaLinks: ['Social Links', 'Social / Associated Profile Links', 'Social Media Links', 'Social Media'],
  firstEmailSubject: ['1st Email Subject', 'First Email Subject'],
  firstEmailMessage: ['1st Email', '1st Email Message', 'First Email Message'],
  followUpEmail1Subject: ['Follow Up 1 Subject', 'Follow-up 1 Subject', 'Followup 1 Subject'],
  followUpEmail1Message: ['Follow Up 1', 'Follow Up Email 1 Message', 'Follow-up 1', 'Followup 1'],
  followUpEmail2Subject: ['Follow Up 2 Subject', 'Follow-up 2 Subject', 'Followup 2 Subject'],
  followUpEmail2Message: ['Follow Up 2', 'Follow Up Email 2 Message', 'Follow-up 2', 'Followup 2'],
  firstEmailStatus: ['1st Email Status', 'First Email Status'],
  followUp1Status: ['Follow Up 1 Status', 'Follow-up 1 Status', 'Followup 1 Status'],
  followUp2Status: ['Follow Up 2 Status', 'Follow-up 2 Status', 'Followup 2 Status'],
  firstEmailDate: ['1st Email Date', 'First Email Date'],
  followUp1Date: ['Follow Up 1 Date', 'Follow-up 1 Date', 'Followup 1 Date'],
  followUp2Date: ['Follow Up 2 Date', 'Follow-up 2 Date', 'Followup 2 Date'],
  leadStatus: ['Lead Status', 'Overall Status']
};

const PRICING_COPY = 'The redesign can be kept affordable, usually around $200 to $800 depending on what you need. You do not need to pay anything upfront. Payment is only after the website is delivered and you are happy with the result.';
const STATUS_VALIDATION_FORMULA = '"Draft,Sent,Replied,Mail Not Found,Bounce Back,Closed"';
const LEAD_STATUS_VALIDATION_FORMULA = '"Win,Lost"';

const dom = {
  leadsBody: document.getElementById('leadsBody'),
  leadDetail: document.getElementById('leadDetail'),
  drawerBackdrop: document.getElementById('drawerBackdrop'),
  searchInput: document.getElementById('searchInput'),
  clearSearchButton: document.getElementById('clearSearchButton'),
  statusFilter: document.getElementById('statusFilter'),
  openFileButton: document.getElementById('openFileButton'),
  saveButton: document.getElementById('saveButton'),
  downloadButton: document.getElementById('downloadButton'),
  fileInput: document.getElementById('fileInput'),
  autoSaveToggle: document.getElementById('autoSaveToggle'),
  saveState: document.getElementById('saveState'),
  workbookChip: document.getElementById('workbookChip'),
  pipelineSummary: document.getElementById('pipelineSummary'),
  pipelineTrack: document.getElementById('pipelineTrack'),
  resultCount: document.getElementById('resultCount'),
  tableFooter: document.getElementById('tableFooter'),
  totalCount: document.getElementById('totalCount'),
  draftCount: document.getElementById('draftCount'),
  firstCount: document.getElementById('firstCount'),
  followUp1Count: document.getElementById('followUp1Count'),
  followUp2Count: document.getElementById('followUp2Count'),
  repliedCount: document.getElementById('repliedCount'),
  mailNotFoundCount: document.getElementById('mailNotFoundCount'),
  bounceBackCount: document.getElementById('bounceBackCount'),
  winCount: document.getElementById('winCount'),
  lostCount: document.getElementById('lostCount'),
  modal: document.getElementById('messageModal'),
  modalCompany: document.getElementById('modalCompany'),
  modalTitle: document.getElementById('modalTitle'),
  modalSubjectRow: document.getElementById('modalSubjectRow'),
  modalSubject: document.getElementById('modalSubject'),
  modalMessage: document.getElementById('modalMessage'),
  copySubjectButton: document.getElementById('copySubjectButton'),
  copyMessageButton: document.getElementById('copyMessageButton')
};

let leads = [];
let workbookSources = [];
let dirty = false;
const dirtySourceIds = new Set();
let autosaveTimer = null;
let selectedLeadId = null;
let editingLeadId = null;

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function normalizeText(value) {
  return String(value || '').trim();
}

function normalizeStatus(value) {
  const raw = normalizeText(value);
  if (!raw) return 'Draft';
  const lower = raw.toLowerCase().replace(/[-_]+/g, ' ').replace(/\s+/g, ' ');
  if (lower === 'sent' || lower === 'contacted' || lower === 'sent first email' || lower === 'sent 1st email' || lower === 'follow up sent' || lower === 'followup sent' || lower === 'follow up email sent') return 'Sent';
  if (lower === 'replied' || lower === 'reply') return 'Replied';
  if (lower === 'mail not found' || lower === 'email not found' || lower === 'not found') return 'Mail Not Found';
  if (lower === 'bounce back' || lower === 'bounced' || lower === 'bounce') return 'Bounce Back';
  if (lower === 'closed' || lower === 'won') return 'Closed';
  if (lower === 'draft' || lower === 'not recorded' || lower === 'not contacted' || lower === 'not interested' || lower === 'lost') return 'Draft';
  return STATUS_OPTIONS.includes(raw) ? raw : 'Draft';
}

function normalizeLeadStatus(value) {
  const raw = normalizeText(value);
  const lower = raw.toLowerCase().replace(/[-_]+/g, ' ').replace(/\s+/g, ' ');
  if (lower === 'win' || lower === 'won') return 'Win';
  if (lower === 'closed' || lower === 'lost' || lower === 'not interested') return 'Lost';
  return LEAD_STATUS_OPTIONS.includes(raw) ? raw : '';
}

function splitSubjectFromMessage(value) {
  const message = normalizeText(value);
  const match = message.match(/^Subject:\s*([^\r\n]+)(?:\r?\n)+([\s\S]*)$/i);
  return match
    ? { subject: normalizeText(match[1]), message: normalizeText(match[2]) }
    : { subject: '', message };
}

function normalizeDate(value) {
  const text = normalizeText(value);
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    const first = Number(slashMatch[1]);
    const second = Number(slashMatch[2]);
    const year = Number(slashMatch[3].length === 2 ? `20${slashMatch[3]}` : slashMatch[3]);
    const month = first > 12 ? second : first;
    const day = first > 12 ? first : second;
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return '';
}

function statusClass(status) {
  return `status-${status.replace(/\s+/g, '-')}`;
}

function getValue(row, key) {
  const aliases = HEADER_ALIASES[key];
  const match = aliases.find(header => Object.prototype.hasOwnProperty.call(row, header));
  return match ? row[match] : '';
}

function compactIndexes(headers) {
  const indexes = {};
  let statusCount = 0;
  let datesCount = 0;

  headers.forEach((header, index) => {
    const name = normalizeText(header);
    if (name === 'Status') {
      statusCount += 1;
      if (statusCount === 1) indexes.firstEmailStatus = index;
      if (statusCount === 2) indexes.followUp1Status = index;
      if (statusCount === 3) indexes.followUp2Status = index;
      return;
    }
    if (name === 'Dates') {
      datesCount += 1;
      if (datesCount === 1) indexes.firstEmailDate = index;
      if (datesCount === 2) indexes.followUp1Date = index;
      if (datesCount === 3) indexes.followUp2Date = index;
      return;
    }

    Object.entries(HEADER_ALIASES).forEach(([key, aliases]) => {
      if (aliases.includes(name) && indexes[key] === undefined) indexes[key] = index;
    });
  });

  return indexes;
}

function getByIndex(row, indexes, key) {
  const index = indexes[key];
  return index === undefined ? '' : row[index];
}

function escapeHtml(value) {
  return normalizeText(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function asUrl(value) {
  const url = normalizeText(value);
  if (!url) return '';
  return /^https?:\/\//i.test(url) ? url : `https://${url}`;
}

function asMailto(value) {
  const email = normalizeText(value);
  return email ? `mailto:${email}` : '';
}

function splitLinks(value) {
  return parseSocialLinks(value)
    .map(link => link.url)
    .filter(Boolean);
}

function parseSocialLinks(value) {
  return normalizeText(value)
    .split(/\n+/)
    .map(item => item.trim())
    .filter(Boolean)
    .flatMap(parseSocialLine);
}

function sanitizeSocialLinks(value) {
  return parseSocialLinks(value)
    .map(item => item.url ? `${item.label}: ${item.url}` : item.label)
    .join('\n');
}

function parseSocialLine(item) {
  if (/^email\s*:/i.test(item)) return [];
  if (/[^\s@]+@[^\s@]+\.[^\s@]+/.test(item)) return [];
  if (/not found/i.test(item) && !/https?:\/\//i.test(item)) return [];

  const urls = [...item.matchAll(/https?:\/\/[^\s;]+/gi)].map(match => match[0].replace(/[).,]+$/, ''));
  if (urls.length) {
    return urls.map((url, index) => ({
      label: socialLabelForUrl(item, url, index, urls.length),
      url
    }));
  }

  return socialTextLabels(item).map(label => ({ label, url: '' }));
}

function socialLabelForUrl(line, url, index, total) {
  const lower = `${line} ${url}`.toLowerCase();
  const known = socialTextLabels(lower)[0];
  if (known) return known;

  const prefix = normalizeText(line.split(':')[0]);
  if (/source pages?/i.test(prefix)) return total > 1 ? `Source page ${index + 1}` : 'Source page';
  if (/website contact|contact page/i.test(prefix)) return 'Contact page';
  if (/website profile|profile page/i.test(prefix)) return 'Profile page';
  if (prefix && !/associated profiles?/i.test(prefix)) return prefix;

  try {
    return new URL(url).hostname.replace(/^www\./i, '');
  } catch {
    return `Link ${index + 1}`;
  }
}

function socialTextLabels(value) {
  const labels = [
    ['WhatsApp', /whatsapp|wa\.me/i],
    ['Facebook', /facebook/i],
    ['Instagram', /instagram/i],
    ['LinkedIn', /linkedin/i],
    ['X', /twitter|x\.com/i],
    ['TikTok', /tiktok/i],
    ['YouTube', /youtube/i],
    ['Pinterest', /pinterest/i],
    ['Threads', /threads/i],
    ['Checkatrade', /checkatrade/i],
    ['Houzz', /houzz/i],
    ['FMB', /\bfmb\b/i],
    ['MyBuilder', /mybuilder/i],
    ['Companies House', /companies house/i]
  ];

  return labels
    .filter(([, pattern]) => pattern.test(value))
    .map(([label]) => label);
}

function currentStatus(lead) {
  if (lead.leadStatus === 'Win') return 'Win';
  if (lead.leadStatus === 'Lost') return 'Lost';
  if (lead.firstEmailStatus === 'Closed' || lead.followUp1Status === 'Closed' || lead.followUp2Status === 'Closed') return 'Lost';
  if (lead.firstEmailStatus === 'Replied' || lead.followUp1Status === 'Replied' || lead.followUp2Status === 'Replied') return 'Replied';
  if (lead.firstEmailStatus === 'Sent' || lead.followUp1Status === 'Sent' || lead.followUp2Status === 'Sent' || lead.firstEmailDate || lead.followUp1Date || lead.followUp2Date) return 'Sent';
  return 'Draft';
}

function hasCompletedStage(lead, statusField, dateField) {
  return ['Sent', 'Replied', 'Closed'].includes(lead[statusField]) || Boolean(lead[dateField]);
}

function matchesStatusFilter(lead, filter) {
  if (filter === 'all') return true;
  if (filter === 'firstEmail') return hasCompletedStage(lead, 'firstEmailStatus', 'firstEmailDate');
  if (filter === 'followUp1') return hasCompletedStage(lead, 'followUp1Status', 'followUp1Date');
  if (filter === 'followUp2') return hasCompletedStage(lead, 'followUp2Status', 'followUp2Date');
  if (['Replied', 'Mail Not Found', 'Bounce Back'].includes(filter)) {
    return [lead.firstEmailStatus, lead.followUp1Status, lead.followUp2Status].includes(filter);
  }
  return currentStatus(lead) === filter;
}

function countEmailStatus(status) {
  return leads.filter(lead => [lead.firstEmailStatus, lead.followUp1Status, lead.followUp2Status].includes(status)).length;
}

function mapRows(rows, sourceId) {
  const headerRowIndex = rows.findIndex(row => row.some(value => normalizeText(value) === 'Name'));
  const headers = rows[headerRowIndex] || [];
  const indexes = compactIndexes(headers);

  return rows
    .slice(headerRowIndex + 1)
    .map((row, index) => {
      const firstEmail = splitSubjectFromMessage(getByIndex(row, indexes, 'firstEmailMessage'));
      const followUp1 = splitSubjectFromMessage(getByIndex(row, indexes, 'followUpEmail1Message'));
      const followUp2 = splitSubjectFromMessage(getByIndex(row, indexes, 'followUpEmail2Message'));
      return {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${index}`,
      sourceId,
      companyName: normalizeText(getByIndex(row, indexes, 'companyName')),
      website: normalizeText(getByIndex(row, indexes, 'website')),
      email: normalizeText(getByIndex(row, indexes, 'email')),
      contactNumber: normalizeText(getByIndex(row, indexes, 'contactNumber')),
      socialMediaLinks: sanitizeSocialLinks(getByIndex(row, indexes, 'socialMediaLinks')),
      firstEmailSubject: normalizeText(getByIndex(row, indexes, 'firstEmailSubject')) || firstEmail.subject,
      firstEmailMessage: firstEmail.message,
      followUpEmail1Subject: normalizeText(getByIndex(row, indexes, 'followUpEmail1Subject')) || followUp1.subject,
      followUpEmail1Message: followUp1.message,
      followUpEmail2Subject: normalizeText(getByIndex(row, indexes, 'followUpEmail2Subject')) || followUp2.subject,
      followUpEmail2Message: followUp2.message,
      firstEmailStatus: normalizeStatus(getByIndex(row, indexes, 'firstEmailStatus')),
      followUp1Status: normalizeStatus(getByIndex(row, indexes, 'followUp1Status')),
      followUp2Status: normalizeStatus(getByIndex(row, indexes, 'followUp2Status')),
      firstEmailDate: normalizeDate(getByIndex(row, indexes, 'firstEmailDate')),
      followUp1Date: normalizeDate(getByIndex(row, indexes, 'followUp1Date')),
      followUp2Date: normalizeDate(getByIndex(row, indexes, 'followUp2Date')),
      leadStatus: normalizeLeadStatus(getByIndex(row, indexes, 'leadStatus'))
      };
    })
    .filter(lead => lead.companyName || lead.website || lead.email);
}

function parseWorkbook(arrayBuffer, name, handle = null) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
  const sheetName = workbook.Sheets.Leads ? 'Leads' : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  const sourceId = crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${name}`;
  return {
    source: { id: sourceId, name: name || FALLBACK_WORKBOOK_NAME, handle },
    leads: mapRows(rows, sourceId)
  };
}

function readWorkbooks(entries) {
  const parsed = entries.map(entry => parseWorkbook(entry.arrayBuffer, entry.name, entry.handle));
  workbookSources = parsed.map(item => item.source);
  leads = parsed.flatMap(item => item.leads);
  selectedLeadId = leads[0]?.id || null;
  dirtySourceIds.clear();
  dirty = false;
  dom.saveButton.disabled = !leads.length;
  dom.downloadButton.disabled = !leads.length;
  const connected = workbookSources.every(source => source.handle);
  setStatus(`${leads.length} leads from ${workbookSources.length} workbook${workbookSources.length === 1 ? '' : 's'}`);
  dom.workbookChip.textContent = `${connected ? 'Connected to' : 'Loaded'} ${workbookSources.length} workbook${workbookSources.length === 1 ? '' : 's'}`;
  render();
}

async function loadDefaultWorkbook() {
  if (window.location.protocol === 'file:') {
    setStatus('Open with localhost to auto-load Excel');
    dom.workbookChip.textContent = 'Browser security blocks automatic folder reads from file://. Use http://127.0.0.1:8000.';
    dom.leadsBody.innerHTML = '<tr><td colspan="8" class="empty-state">Open this tracker from the local server URL to load the Excel folder automatically, or click Open Excel files.</td></tr>';
    dom.leadDetail.innerHTML = '<div class="empty-state">Choose the Excel file once to edit lead details here.</div>';
    return;
  }

  try {
    const entries = await Promise.all(DEFAULT_WORKBOOK_PATHS.map(async path => {
      const response = await fetch(`${path}?v=${Date.now()}`, { cache: 'no-store' });
      if (!response.ok) throw new Error(`${path} returned ${response.status}.`);
      return { arrayBuffer: await response.arrayBuffer(), name: path.split('/').pop(), handle: null };
    }));
    readWorkbooks(entries);
  } catch (error) {
    console.error('Default workbook load failed:', error);
    setStatus('Excel folder file did not load');
    dom.workbookChip.textContent = 'Excel folder';
    dom.leadsBody.innerHTML = '<tr><td colspan="8" class="empty-state">Could not load the Excel folder automatically. Start the local server from this folder, then open http://127.0.0.1:8000.</td></tr>';
    dom.leadDetail.innerHTML = '<div class="empty-state">Workbook details will appear here after loading.</div>';
  }
}

function setStatus(message) {
  dom.saveState.textContent = message;
}

function filteredLeads() {
  const query = dom.searchInput.value.trim().toLowerCase();
  const filter = dom.statusFilter.value;
  return leads.filter(lead => {
    const haystack = `${lead.companyName} ${lead.website} ${lead.email} ${lead.contactNumber}`.toLowerCase();
    const matchesSearch = !query || haystack.includes(query);
    const matchesFilter = matchesStatusFilter(lead, filter);
    return matchesSearch && matchesFilter;
  });
}

function render() {
  const visible = filteredLeads();
  dom.resultCount.textContent = visible.length;
  dom.tableFooter.textContent = `Showing ${visible.length} of ${leads.length} leads`;
  renderStats();

  if (!visible.length) {
    dom.leadsBody.innerHTML = `<tr><td colspan="8" class="empty-state">${leads.length ? 'No leads match the current search or filter.' : 'Open your Excel workbook to load leads.'}</td></tr>`;
    dom.leadDetail.innerHTML = '<div class="empty-state">No selected lead.</div>';
    return;
  }

  if (!visible.some(lead => lead.id === selectedLeadId)) {
    selectedLeadId = visible[0].id;
  }

  dom.leadsBody.innerHTML = visible.map(lead => renderLeadRow(lead)).join('');
  renderLeadDetail(findLead(selectedLeadId));
}

function renderStats() {
  const totals = PIPELINE_STATUSES.reduce((memo, status) => ({ ...memo, [status]: 0 }), {});
  const stages = { firstEmail: 0, followUp1: 0, followUp2: 0 };
  leads.forEach(lead => {
    totals[currentStatus(lead)] += 1;
    if (hasCompletedStage(lead, 'firstEmailStatus', 'firstEmailDate')) stages.firstEmail += 1;
    if (hasCompletedStage(lead, 'followUp1Status', 'followUp1Date')) stages.followUp1 += 1;
    if (hasCompletedStage(lead, 'followUp2Status', 'followUp2Date')) stages.followUp2 += 1;
  });
  dom.totalCount.textContent = leads.length;
  dom.draftCount.textContent = totals.Draft;
  dom.firstCount.textContent = stages.firstEmail;
  dom.followUp1Count.textContent = stages.followUp1;
  dom.followUp2Count.textContent = stages.followUp2;
  dom.repliedCount.textContent = countEmailStatus('Replied');
  dom.mailNotFoundCount.textContent = countEmailStatus('Mail Not Found');
  dom.bounceBackCount.textContent = countEmailStatus('Bounce Back');
  dom.winCount.textContent = totals.Win;
  dom.lostCount.textContent = totals.Lost;
  renderPipeline(totals);
  document.querySelectorAll('.stat-card[data-filter]').forEach(card => {
    card.classList.toggle('active', card.dataset.filter === dom.statusFilter.value);
  });
}

function renderPipeline(totals) {
  if (!leads.length) {
    dom.pipelineSummary.textContent = 'Load leads to see your outreach flow';
    dom.pipelineTrack.innerHTML = '<span></span>';
    return;
  }

  const active = leads.length - totals.Draft;
  const activePercent = Math.round((active / leads.length) * 100);
  dom.pipelineSummary.textContent = `${active} active · ${totals.Replied} replied · ${totals.Win} won · ${totals.Lost} lost`;
  dom.pipelineTrack.innerHTML = PIPELINE_STATUSES.map(status => {
    const width = (totals[status] / leads.length) * 100;
    return `<span class="${statusClass(status)}" style="width:${width}%"></span>`;
  }).join('');
  dom.pipelineTrack.setAttribute('title', `${activePercent}% of leads have moved beyond Draft`);
}

function renderLeadRow(lead) {
  const website = asUrl(lead.website);
  const email = asMailto(lead.email);

  return `
    <tr class="${lead.id === selectedLeadId ? 'selected' : ''}" data-select="${lead.id}">
      <td>
        <strong class="sheet-name">${escapeHtml(lead.companyName || 'Unnamed lead')}</strong>
      </td>
      <td>${website ? `<a href="${escapeHtml(website)}" target="_blank" rel="noreferrer">${escapeHtml(shortUrl(lead.website))}</a>` : '<span class="empty-value">Not recorded</span>'}</td>
      <td>${email ? `<a href="${escapeHtml(email)}">${escapeHtml(lead.email)}</a>` : '<span class="empty-value">Not recorded</span>'}</td>
      <td>${escapeHtml(lead.contactNumber) || '<span class="empty-value">Not recorded</span>'}</td>
      <td>${sheetStatus(lead.firstEmailStatus, lead.firstEmailDate)}</td>
      <td>${sheetStatus(lead.followUp1Status, lead.followUp1Date)}</td>
      <td>${sheetStatus(lead.followUp2Status, lead.followUp2Date)}</td>
      <td><span class="status-pill ${statusClass(lead.leadStatus || 'Not Set')}">${escapeHtml(lead.leadStatus || 'Not set')}</span></td>
    </tr>
  `;
}

function shortUrl(value) {
  return normalizeText(value)
    .replace(/^https?:\/\//i, '')
    .replace(/^www\./i, '')
    .replace(/\/$/, '');
}

function sheetStatus(status, date) {
  return `
    <div class="sheet-status">
      <span class="mini-status ${statusClass(status)}">${escapeHtml(status)}</span>
      ${date ? `<small>${escapeHtml(date)}</small>` : ''}
    </div>
  `;
}

function renderLeadDetail(lead) {
  if (!lead) {
    dom.leadDetail.innerHTML = '<div class="empty-state">Select a lead to view outreach details.</div>';
    return;
  }

  const website = asUrl(lead.website);
  const email = asMailto(lead.email);
  const socialLinks = parseSocialLinks(lead.socialMediaLinks);
  const socialHtml = socialLinks.length
    ? `<div class="social-link-list">${socialLinks.map((link, index) => socialLinkItem(link, index)).join('')}</div>`
    : 'Not recorded';

  dom.leadDetail.innerHTML = `
    <button class="drawer-close" type="button" data-close-drawer aria-label="Close details">x</button>
    <div class="detail-hero">
      <div class="lead-avatar large">${escapeHtml(initials(lead.companyName || 'Lead'))}</div>
      <div>
        <p class="eyebrow">SELECTED LEAD</p>
        <h3>${escapeHtml(lead.companyName || 'Unnamed lead')}</h3>
        <span class="status-pill ${statusClass(lead.leadStatus || 'Not Set')}">${escapeHtml(lead.leadStatus || 'Not set')}</span>
      </div>
      <button class="secondary-button detail-edit-button" type="button" data-edit-lead="${lead.id}">${editingLeadId === lead.id ? 'Editing lead' : 'Edit lead'}</button>
    </div>

    ${editingLeadId === lead.id ? editLeadFields(lead) : `<div class="detail-contact-grid">
      ${contactItem('Website', website ? `<a href="${escapeHtml(website)}" target="_blank" rel="noreferrer">${escapeHtml(lead.website)}</a>` : 'Not recorded', '', lead.website)}
      ${contactItem('Email', email ? `<a href="${escapeHtml(email)}">${escapeHtml(lead.email)}</a>` : 'Not recorded', '', lead.email)}
      ${contactItem('Phone', escapeHtml(lead.contactNumber) || 'Not recorded', '', lead.contactNumber)}
      ${contactItem('Social', socialHtml, 'social-contact')}
    </div>`}

    <label class="lead-status-control detail-status">
      <span>Lead Status</span>
      ${statusSelect(lead, 'leadStatus')}
    </label>

    <div class="stage-grid detail-stages">
      ${stageCard(lead, '1st Email', 'firstEmailMessage', 'firstEmailStatus', 'firstEmailDate')}
      ${stageCard(lead, 'Follow Up 1', 'followUpEmail1Message', 'followUp1Status', 'followUp1Date')}
      ${stageCard(lead, 'Follow Up 2', 'followUpEmail2Message', 'followUp2Status', 'followUp2Date')}
    </div>
`;
}

function socialLinkItem(link, index) {
  const label = escapeHtml(link.label || `Link ${index + 1}`);
  if (!link.url) {
    return `<span class="social-note"><span>${label}</span><small>Mentioned</small></span>`;
  }

  const url = asUrl(link.url);
  return `
    <a class="social-link-item" href="${escapeHtml(url)}" target="_blank" rel="noreferrer">
      <span>${label}</span>
      <small>${escapeHtml(shortUrl(link.url))}</small>
    </a>
  `;
}

function miniStep(label, status) {
  return `<span class="${statusClass(status)}">${label}</span>`;
}

function initials(name) {
  return normalizeText(name)
    .split(/\s+/)
    .filter(Boolean)
    .map(word => word[0])
    .join('')
    .slice(0, 2)
    .toUpperCase() || 'LT';
}

function contactItem(label, value, className = '', copyValue = '') {
  return `
    <div class="contact-item ${className}">
      <span>${label}</span>
      <strong>${value}</strong>
      ${copyValue ? `<button class="copy-contact-button" type="button" data-copy="${escapeHtml(copyValue)}" aria-label="Copy ${label}">Copy</button>` : ''}
    </div>
  `;
}

function editLeadFields(lead) {
  const fields = [
    ['Company name', 'companyName', 'text'],
    ['Website', 'website', 'text'],
    ['Email', 'email', 'email'],
    ['Phone', 'contactNumber', 'text']
  ];
  return `<div class="lead-edit-form">
    ${fields.map(([label, field, type]) => `<label><span>${label}</span><input type="${type}" data-edit-field="${field}" value="${escapeHtml(lead[field])}"></label>`).join('')}
    <label class="edit-wide"><span>Social links</span><textarea data-edit-field="socialMediaLinks" rows="3">${escapeHtml(lead.socialMediaLinks)}</textarea></label>
    ${emailEditField(lead, '1st Email', 'firstEmailSubject', 'firstEmailMessage')}
    ${emailEditField(lead, 'Follow Up 1', 'followUpEmail1Subject', 'followUpEmail1Message')}
    ${emailEditField(lead, 'Follow Up 2', 'followUpEmail2Subject', 'followUpEmail2Message')}
    <div class="edit-actions"><button class="primary-button" type="button" data-save-lead="${lead.id}">Save edits</button><button class="secondary-button" type="button" data-cancel-edit>Cancel</button></div>
  </div>`;
}

function emailEditField(lead, label, subjectField, messageField) {
  return `<fieldset class="edit-wide email-edit-field"><legend>${label}</legend><label><span>Subject</span><input type="text" data-edit-field="${subjectField}" value="${escapeHtml(lead[subjectField])}"></label><label><span>Message</span><textarea data-edit-field="${messageField}" rows="5">${escapeHtml(lead[messageField])}</textarea></label></fieldset>`;
}

function stageCard(lead, title, messageField, statusField, dateFieldName) {
  return `
    <section class="stage-card">
      <div class="stage-heading">
        <span>${title}</span>
        ${messageButton(lead, messageField, 'Email')}
      </div>
      <div class="stage-controls">
        <label>
          <span>Status</span>
          ${statusSelect(lead, statusField)}
        </label>
        <label>
          <span>Date</span>
          ${dateField(lead, dateFieldName)}
        </label>
      </div>
    </section>
  `;
}

function statusSelect(lead, field) {
  if (field === 'leadStatus') return leadStatusSelect(lead);

  return `
    <select class="status-select ${statusClass(lead[field])}" data-id="${lead.id}" data-field="${field}" aria-label="${field}">
      ${STATUS_OPTIONS.map(status => `<option value="${status}" ${lead[field] === status ? 'selected' : ''}>${status}</option>`).join('')}
    </select>
  `;
}

function leadStatusSelect(lead) {
  return `
    <select class="status-select ${statusClass(lead.leadStatus || 'Not Set')}" data-id="${lead.id}" data-field="leadStatus" aria-label="Lead Status">
      <option value="" ${lead.leadStatus ? '' : 'selected'} disabled>Select</option>
      ${LEAD_STATUS_OPTIONS.map(status => `<option value="${status}" ${lead.leadStatus === status ? 'selected' : ''}>${status}</option>`).join('')}
    </select>
  `;
}

function dateField(lead, field) {
  return `
    <input class="date-input single-date" type="date" data-id="${lead.id}" data-field="${field}" value="${escapeHtml(lead[field])}">
  `;
}

function messageButton(lead, field, label) {
  return `
    <button class="message-button" type="button" data-message="${field}" data-id="${lead.id}" ${lead[field] ? '' : 'disabled'}>${label}</button>
  `;
}

function findLead(id) {
  return leads.find(lead => lead.id === id);
}

function saveLeadEdits(id) {
  const lead = findLead(id);
  if (!lead) return;
  dom.leadDetail.querySelectorAll('[data-edit-field]').forEach(input => {
    const field = input.dataset.editField;
    const value = input.value;
    lead[field] = field === 'socialMediaLinks' ? sanitizeSocialLinks(value) : normalizeText(value);
  });
  editingLeadId = null;
  markDirty(lead.sourceId);
  render();
}

function updateLead(id, field, value) {
  const lead = findLead(id);
  if (!lead) return;
  const oldValue = lead[field];
  lead[field] = field === 'leadStatus' ? normalizeLeadStatus(value) : field.includes('Status') ? normalizeStatus(value) : normalizeText(value);

  if (field.includes('Status')) {
    applyAutomaticDate(lead, field, lead[field]);
    updateLeadStatus(lead, field);
  }

  if (oldValue !== lead[field]) {
    markDirty(lead.sourceId);
    render();
  }
}

function applyAutomaticDate(lead, field, status) {
  const date = todayIso();
  if ((status === 'Sent' || status === 'Replied' || status === 'Closed') && field === 'firstEmailStatus' && !lead.firstEmailDate) lead.firstEmailDate = date;
  if ((status === 'Sent' || status === 'Replied' || status === 'Closed') && field === 'followUp1Status' && !lead.followUp1Date) lead.followUp1Date = date;
  if ((status === 'Sent' || status === 'Replied' || status === 'Closed') && field === 'followUp2Status' && !lead.followUp2Date) lead.followUp2Date = date;
}

function updateLeadStatus(lead, field) {
  if (field === 'leadStatus') return;
  if (lead.firstEmailStatus === 'Closed' || lead.followUp1Status === 'Closed' || lead.followUp2Status === 'Closed') {
    lead.leadStatus = 'Lost';
  }
}

function markDirty(sourceId) {
  dirty = true;
  if (sourceId) dirtySourceIds.add(sourceId);
  setStatus('Unsaved changes');
  if (dom.autoSaveToggle.checked) {
    window.clearTimeout(autosaveTimer);
    autosaveTimer = window.setTimeout(() => saveWorkbook({ quiet: true }), 700);
  }
}

function toWorkbookRows(sourceId = null) {
  return leads
    .filter(lead => !sourceId || lead.sourceId === sourceId)
    .map(lead => [
    lead.companyName,
    lead.website,
    lead.email,
    lead.contactNumber,
    sanitizeSocialLinks(lead.socialMediaLinks),
    lead.firstEmailSubject,
    lead.firstEmailMessage,
    lead.firstEmailStatus,
    lead.firstEmailDate,
    lead.followUpEmail1Subject,
    addPricingCopyIfUseful(lead.followUpEmail1Message),
    lead.followUp1Status,
    lead.followUp1Date,
    lead.followUpEmail2Subject,
    addPricingCopyIfUseful(lead.followUpEmail2Message),
    lead.followUp2Status,
    lead.followUp2Date,
    lead.leadStatus
  ]);
}

function addPricingCopyIfUseful(message) {
  const text = normalizeText(message);
  if (!text) return '';
  if (/200\s*to\s*800|pay anything upfront|payment is only after/i.test(text)) return text;
  return `${text}\n\n${PRICING_COPY}`;
}

async function buildWorkbookArray(sourceId = null) {
  const worksheet = XLSX.utils.aoa_to_sheet([REQUIRED_HEADERS, ...toWorkbookRows(sourceId)]);
  worksheet['!cols'] = [
    18, 30, 30, 18, 24, 36, 52, 18, 15, 36, 52, 18, 15, 36, 52, 18, 15, 20
  ].map(width => ({ wch: width }));
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Leads');
  const output = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  return addExcelDropdownValidation(output);
}

async function addExcelDropdownValidation(output) {
  if (!globalThis.JSZip) return output;

  const zip = await globalThis.JSZip.loadAsync(output);
  const sheetPath = 'xl/worksheets/sheet1.xml';
  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) return output;

  let xml = await sheetFile.async('string');
  xml = xml.replace(/<dataValidations[\s\S]*?<\/dataValidations>/, '');

  const validationXml = [
    '<dataValidations count="2">',
    `<dataValidation type="list" allowBlank="0" showErrorMessage="1" sqref="H2:H1000 L2:L1000 P2:P1000"><formula1>${escapeXml(STATUS_VALIDATION_FORMULA)}</formula1></dataValidation>`,
    `<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="R2:R1000"><formula1>${escapeXml(LEAD_STATUS_VALIDATION_FORMULA)}</formula1></dataValidation>`,
    '</dataValidations>'
  ].join('');

  if (xml.includes('<pageMargins')) {
    xml = xml.replace('<pageMargins', `${validationXml}<pageMargins`);
  } else if (xml.includes('</worksheet>')) {
    xml = xml.replace('</worksheet>', `${validationXml}</worksheet>`);
  }

  zip.file(sheetPath, xml);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

async function verifyFilePermission(handle, mode = 'readwrite') {
  if (!handle || !handle.queryPermission) return false;
  const options = { mode };
  if ((await handle.queryPermission(options)) === 'granted') return true;
  return (await handle.requestPermission(options)) === 'granted';
}

async function writeWorkbookToHandle(handle, output) {
  const writable = await handle.createWritable();
  await writable.write(new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  await writable.close();
}

async function chooseSaveHandle() {
  if (!('showSaveFilePicker' in window)) return null;

  const handle = await window.showSaveFilePicker({
    suggestedName: workbookSources[0]?.name || FALLBACK_WORKBOOK_NAME,
    types: [{
      description: 'Excel workbooks',
      accept: {
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
      }
    }]
  });
  await saveHandles([handle]);
  return handle;
}

async function saveWorkbook({ quiet = false } = {}) {
  if (!leads.length) return;
  let output = null;

  try {
    const dirtySources = workbookSources.filter(source => dirtySourceIds.has(source.id));
    if (!dirtySources.length) {
      setStatus('No unsaved changes');
      return;
    }

    const connectedSources = dirtySources.filter(source => source.handle);
    if (connectedSources.length === dirtySources.length && dirtySources.every(source => source.handle)) {
      for (const source of dirtySources) {
        if (!await verifyFilePermission(source.handle, 'readwrite')) throw new Error(`Write permission denied for ${source.name}.`);
        await writeWorkbookToHandle(source.handle, await buildWorkbookArray(source.id));
      }
      dirtySources.forEach(source => dirtySourceIds.delete(source.id));
      dirty = false;
      setStatus(quiet ? `Auto-saved ${dirtySources.length} workbook${dirtySources.length === 1 ? '' : 's'}` : `Saved ${dirtySources.length} workbook${dirtySources.length === 1 ? '' : 's'}`);
      return;
    }

    if (!quiet) {
      if (dirtySources.length === 1 && 'showSaveFilePicker' in window) {
        output = await buildWorkbookArray(dirtySources[0].id);
        const pickedHandle = await chooseSaveHandle();
        await writeWorkbookToHandle(pickedHandle, output);
        dirtySources[0].handle = pickedHandle;
        dirtySources[0].name = pickedHandle.name || dirtySources[0].name;
        dirtySourceIds.delete(dirtySources[0].id);
        dirty = false;
        dom.workbookChip.textContent = `Connected to ${workbookSources.length} workbooks`;
        setStatus('Saved to Excel');
        return;
      }

      const sourceId = dirtySources.length === 1 ? dirtySources[0].id : null;
      output = await buildWorkbookArray(sourceId);
      await downloadWorkbook(output, sourceId);
      dirtySources.forEach(source => dirtySourceIds.delete(source.id));
      dirty = false;
      setStatus(sourceId ? 'Downloaded updated workbook copy' : 'Downloaded combined updated copy');
    } else {
      setStatus('Auto-save needs Open Excel permission');
    }
  } catch (error) {
    if (error?.name === 'AbortError') {
      setStatus('Save cancelled');
      return;
    }
    console.error('Workbook save failed:', error);
    dom.autoSaveToggle.checked = false;
    if (!quiet && output) {
      await downloadWorkbook(output);
      setStatus('Excel was locked. Downloaded an updated copy.');
      return;
    }
    setStatus('Save failed. Close Excel, then click Save.');
  }
}

async function downloadWorkbook(output, sourceId = null) {
  const workbookOutput = output || await buildWorkbookArray();
  const link = document.createElement('a');
  const source = sourceId ? workbookSources.find(item => item.id === sourceId) : null;
  const baseName = source ? source.name : workbookSources.length === 1 ? workbookSources[0].name : 'all-leads';
  const safeName = baseName.replace(/\.(xlsx|xls)$/i, '') || 'leads-tracker';
  link.href = URL.createObjectURL(new Blob([workbookOutput], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  link.download = `${safeName}-updated.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

async function openFile() {
  if ('showOpenFilePicker' in window) {
    const handles = await window.showOpenFilePicker({
      multiple: true,
      types: [{
        description: 'Excel workbooks',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
          'application/vnd.ms-excel': ['.xls']
        }
      }]
    });
    const entries = await Promise.all(handles.map(async handle => {
      const file = await handle.getFile();
      return { arrayBuffer: await file.arrayBuffer(), name: file.name, handle };
    }));
    await saveHandles(handles);
    readWorkbooks(entries);
    return;
  }
  dom.fileInput.click();
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(HANDLE_DB, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(HANDLE_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveHandles(handles) {
  if (!('indexedDB' in window)) return;
  const db = await openDb();
  const tx = db.transaction(HANDLE_STORE, 'readwrite');
  tx.objectStore(HANDLE_STORE).put(handles, HANDLE_KEY);
}

async function getSavedHandles() {
  if (!('indexedDB' in window)) return null;
  const db = await openDb();
  return new Promise(resolve => {
    const tx = db.transaction(HANDLE_STORE, 'readonly');
    const request = tx.objectStore(HANDLE_STORE).get(HANDLE_KEY);
    request.onsuccess = () => {
      const stored = request.result;
      resolve(stored ? (Array.isArray(stored) ? stored : [stored]) : []);
    };
    request.onerror = () => resolve([]);
  });
}

async function loadSavedHandles() {
  if (!('showOpenFilePicker' in window)) return false;
  const handles = await getSavedHandles();
  if (!handles.length) return false;
  const entries = [];
  for (const handle of handles) {
    if (!await verifyFilePermission(handle, 'read')) return false;
    const file = await handle.getFile();
    entries.push({ arrayBuffer: await file.arrayBuffer(), name: file.name, handle });
  }
  readWorkbooks(entries);
  return true;
}

function showMessage(id, field) {
  const lead = findLead(id);
  if (!lead) return;
  const labels = {
    firstEmailMessage: '1st Email',
    followUpEmail1Message: 'Follow Up 1',
    followUpEmail2Message: 'Follow Up 2'
  };
  const subjectFields = {
    firstEmailMessage: 'firstEmailSubject',
    followUpEmail1Message: 'followUpEmail1Subject',
    followUpEmail2Message: 'followUpEmail2Subject'
  };
  const subject = lead[subjectFields[field]] || '';
  dom.modalCompany.textContent = lead.companyName || 'Lead message';
  dom.modalTitle.textContent = labels[field];
  dom.modalSubject.textContent = subject;
  dom.modalSubjectRow.hidden = !subject;
  dom.modalMessage.textContent = lead[field] || '';
  dom.modal.showModal();
}

dom.openFileButton.addEventListener('click', async () => {
  try {
    await openFile();
  } catch (error) {
    if (error.name !== 'AbortError') {
      alert(`Could not open workbook: ${error.message}`);
    }
  }
});

dom.fileInput.addEventListener('change', async event => {
  const files = [...event.target.files];
  if (!files.length) return;
  const entries = await Promise.all(files.map(async file => ({
    arrayBuffer: await file.arrayBuffer(),
    name: file.name,
    handle: null
  })));
  readWorkbooks(entries);
  event.target.value = '';
});

dom.saveButton.addEventListener('click', () => saveWorkbook());
dom.downloadButton.addEventListener('click', () => downloadWorkbook());
dom.searchInput.addEventListener('input', () => {
  dom.clearSearchButton.hidden = !dom.searchInput.value;
  render();
});
dom.clearSearchButton.addEventListener('click', () => {
  dom.searchInput.value = '';
  dom.clearSearchButton.hidden = true;
  dom.searchInput.focus();
  render();
});
dom.statusFilter.addEventListener('change', render);
dom.autoSaveToggle.addEventListener('change', () => {
  if (dom.autoSaveToggle.checked && dirty) saveWorkbook({ quiet: true });
});

document.querySelectorAll('.stat-card[data-filter]').forEach(card => {
  const applyFilter = () => {
    dom.statusFilter.value = card.dataset.filter;
    render();
    document.querySelector('.table-panel').scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  card.addEventListener('click', applyFilter);
  card.addEventListener('keydown', event => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      applyFilter();
    }
  });
});

dom.leadsBody.addEventListener('click', event => {
  const row = event.target.closest('[data-select]');
  if (row) {
    selectedLeadId = row.dataset.select;
    render();
    openDrawer();
    return;
  }

  const button = event.target.closest('button[data-message]');
  if (button) showMessage(button.dataset.id, button.dataset.message);
});

dom.leadDetail.addEventListener('change', event => {
  const target = event.target;
  if (!target.dataset.id || !target.dataset.field) return;
  updateLead(target.dataset.id, target.dataset.field, target.value);
});

dom.leadDetail.addEventListener('click', event => {
  if (event.target.closest('[data-close-drawer]')) {
    closeDrawer();
    return;
  }

  const button = event.target.closest('button[data-message]');
  if (button) showMessage(button.dataset.id, button.dataset.message);

  const editButton = event.target.closest('[data-edit-lead]');
  if (editButton) {
    editingLeadId = editButton.dataset.editLead;
    renderLeadDetail(findLead(editingLeadId));
    return;
  }

  if (event.target.closest('[data-save-lead]')) {
    saveLeadEdits(event.target.closest('[data-save-lead]').dataset.saveLead);
    return;
  }

  if (event.target.closest('[data-cancel-edit]')) {
    editingLeadId = null;
    renderLeadDetail(findLead(selectedLeadId));
    return;
  }

  const copyButton = event.target.closest('[data-copy]');
  if (copyButton) {
    navigator.clipboard.writeText(copyButton.dataset.copy);
    copyButton.textContent = 'Copied';
    window.setTimeout(() => { copyButton.textContent = 'Copy'; }, 1000);
  }
});

dom.drawerBackdrop.addEventListener('click', closeDrawer);
document.addEventListener('keydown', event => {
  if (event.key === 'Escape') closeDrawer();
});

function openDrawer() {
  dom.leadDetail.classList.add('open');
  dom.drawerBackdrop.hidden = false;
}

function closeDrawer() {
  dom.leadDetail.classList.remove('open');
  dom.drawerBackdrop.hidden = true;
}

dom.copyMessageButton.addEventListener('click', async () => {
  await navigator.clipboard.writeText(dom.modalMessage.textContent);
  dom.copyMessageButton.textContent = 'Copied';
  window.setTimeout(() => {
    dom.copyMessageButton.textContent = 'Copy email';
  }, 1000);
});

dom.copySubjectButton.addEventListener('click', async () => {
  await navigator.clipboard.writeText(dom.modalSubject.textContent);
  dom.copySubjectButton.textContent = 'Copied';
  window.setTimeout(() => {
    dom.copySubjectButton.textContent = 'Copy subject';
  }, 1000);
});

(async function init() {
  dom.autoSaveToggle.checked = true;
  try {
    if (await loadSavedHandles()) return;
  } catch (error) {
    console.warn('Saved file handle could not be restored.', error);
  }
  await loadDefaultWorkbook();
})();
