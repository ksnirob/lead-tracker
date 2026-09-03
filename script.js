const DEFAULT_WORKBOOK_PATH = 'excel/Leads_Completed_Tracker.xlsx';
const FALLBACK_WORKBOOK_NAME = 'leads-tracker.xlsx';
const HANDLE_DB = 'lead-tracker-file-handles';
const HANDLE_STORE = 'handles';
const HANDLE_KEY = 'primary-workbook';

const STATUS_OPTIONS = [
  'Draft',
  'Sent',
  'Replied',
  'Closed'
];

const LEAD_STATUS_OPTIONS = [
  'Win',
  'Closed'
];

const REQUIRED_HEADERS = [
  'Name',
  'Website',
  'Email',
  'Phone',
  'Social Links',
  '1st Email',
  'Status',
  'Dates',
  'Follow Up 1',
  'Status',
  'Dates',
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
  firstEmailMessage: ['1st Email', '1st Email Message', 'First Email Message'],
  followUpEmail1Message: ['Follow Up 1', 'Follow Up Email 1 Message', 'Follow-up 1', 'Followup 1'],
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
const STATUS_VALIDATION_FORMULA = '"Draft,Sent,Replied,Closed"';
const LEAD_STATUS_VALIDATION_FORMULA = '"Win,Closed"';

const dom = {
  leadsBody: document.getElementById('leadsBody'),
  leadDetail: document.getElementById('leadDetail'),
  drawerBackdrop: document.getElementById('drawerBackdrop'),
  searchInput: document.getElementById('searchInput'),
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
  followCount: document.getElementById('followCount'),
  closedCount: document.getElementById('closedCount'),
  modal: document.getElementById('messageModal'),
  modalCompany: document.getElementById('modalCompany'),
  modalTitle: document.getElementById('modalTitle'),
  modalMessage: document.getElementById('modalMessage'),
  copyMessageButton: document.getElementById('copyMessageButton')
};

let leads = [];
let workbookName = FALLBACK_WORKBOOK_NAME;
let fileHandle = null;
let dirty = false;
let autosaveTimer = null;
let selectedLeadId = null;

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
  if (lower === 'closed' || lower === 'won') return 'Closed';
  if (lower === 'draft' || lower === 'not recorded' || lower === 'not contacted' || lower === 'not interested' || lower === 'lost') return 'Draft';
  return STATUS_OPTIONS.includes(raw) ? raw : 'Draft';
}

function normalizeLeadStatus(value) {
  const raw = normalizeText(value);
  const lower = raw.toLowerCase().replace(/[-_]+/g, ' ').replace(/\s+/g, ' ');
  if (lower === 'win' || lower === 'won') return 'Win';
  if (lower === 'closed' || lower === 'lost' || lower === 'not interested') return 'Closed';
  return LEAD_STATUS_OPTIONS.includes(raw) ? raw : '';
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
  if (lead.leadStatus === 'Win') return 'Closed';
  if (lead.leadStatus === 'Closed') return 'Closed';
  if (lead.firstEmailStatus === 'Closed' || lead.followUp1Status === 'Closed' || lead.followUp2Status === 'Closed') return 'Closed';
  if (lead.firstEmailStatus === 'Replied' || lead.followUp1Status === 'Replied' || lead.followUp2Status === 'Replied') return 'Replied';
  if (lead.firstEmailStatus === 'Sent' || lead.followUp1Status === 'Sent' || lead.followUp2Status === 'Sent' || lead.firstEmailDate || lead.followUp1Date || lead.followUp2Date) return 'Sent';
  return 'Draft';
}

function mapRows(rows) {
  const headers = rows[0] || [];
  const indexes = compactIndexes(headers);

  return rows
    .slice(1)
    .map((row, index) => ({
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${index}`,
      companyName: normalizeText(getByIndex(row, indexes, 'companyName')),
      website: normalizeText(getByIndex(row, indexes, 'website')),
      email: normalizeText(getByIndex(row, indexes, 'email')),
      contactNumber: normalizeText(getByIndex(row, indexes, 'contactNumber')),
      socialMediaLinks: sanitizeSocialLinks(getByIndex(row, indexes, 'socialMediaLinks')),
      firstEmailMessage: normalizeText(getByIndex(row, indexes, 'firstEmailMessage')),
      followUpEmail1Message: normalizeText(getByIndex(row, indexes, 'followUpEmail1Message')),
      followUpEmail2Message: normalizeText(getByIndex(row, indexes, 'followUpEmail2Message')),
      firstEmailStatus: normalizeStatus(getByIndex(row, indexes, 'firstEmailStatus')),
      followUp1Status: normalizeStatus(getByIndex(row, indexes, 'followUp1Status')),
      followUp2Status: normalizeStatus(getByIndex(row, indexes, 'followUp2Status')),
      firstEmailDate: normalizeDate(getByIndex(row, indexes, 'firstEmailDate')),
      followUp1Date: normalizeDate(getByIndex(row, indexes, 'followUp1Date')),
      followUp2Date: normalizeDate(getByIndex(row, indexes, 'followUp2Date')),
      leadStatus: normalizeLeadStatus(getByIndex(row, indexes, 'leadStatus'))
    }))
    .filter(lead => lead.companyName || lead.website || lead.email);
}

function readWorkbook(arrayBuffer, name, handle = null) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
  const sheetName = workbook.Sheets.Leads ? 'Leads' : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  leads = mapRows(rows);
  selectedLeadId = leads[0]?.id || null;
  workbookName = name || FALLBACK_WORKBOOK_NAME;
  fileHandle = handle;
  dirty = false;
  dom.saveButton.disabled = !leads.length;
  dom.downloadButton.disabled = !leads.length;
  setStatus(`${leads.length} leads loaded`);
  dom.workbookChip.textContent = handle ? `Connected to ${workbookName}` : `Loaded ${workbookName}`;
  render();
}

async function loadDefaultWorkbook() {
  if (window.location.protocol === 'file:') {
    setStatus('Open with localhost to auto-load Excel');
    dom.workbookChip.textContent = 'Browser security blocks automatic folder reads from file://. Use http://127.0.0.1:8000.';
    dom.leadsBody.innerHTML = '<tr><td colspan="8" class="empty-state">Open this tracker from the local server URL to load excel/London_25_Leads_and_Followups.xlsx automatically, or click Open Excel once.</td></tr>';
    dom.leadDetail.innerHTML = '<div class="empty-state">Choose the Excel file once to edit lead details here.</div>';
    return;
  }

  try {
    const response = await fetch(`${DEFAULT_WORKBOOK_PATH}?v=${Date.now()}`, { cache: 'no-store' });
    if (!response.ok) throw new Error(`Workbook request returned ${response.status}.`);
    const arrayBuffer = await response.arrayBuffer();
    readWorkbook(arrayBuffer, DEFAULT_WORKBOOK_PATH.split('/').pop());
  } catch (error) {
    console.error('Default workbook load failed:', error);
    setStatus('Excel folder file did not load');
    dom.workbookChip.textContent = DEFAULT_WORKBOOK_PATH;
    dom.leadsBody.innerHTML = '<tr><td colspan="8" class="empty-state">Could not load excel/London_25_Leads_and_Followups.xlsx automatically. Start the local server from this folder, then open http://127.0.0.1:8000.</td></tr>';
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
    const matchesFilter = filter === 'all' || currentStatus(lead) === filter;
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
  const totals = STATUS_OPTIONS.reduce((memo, status) => ({ ...memo, [status]: 0 }), {});
  leads.forEach(lead => {
    totals[currentStatus(lead)] += 1;
  });
  dom.totalCount.textContent = leads.length;
  dom.draftCount.textContent = totals.Draft;
  dom.firstCount.textContent = totals.Sent;
  dom.followCount.textContent = totals.Replied;
  dom.closedCount.textContent = totals.Closed;
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
  dom.pipelineSummary.textContent = `${active} active, ${totals.Replied} replied, ${totals.Closed} closed`;
  dom.pipelineTrack.innerHTML = STATUS_OPTIONS.map(status => {
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
    </div>

    <div class="detail-contact-grid">
      ${contactItem('Website', website ? `<a href="${escapeHtml(website)}" target="_blank" rel="noreferrer">${escapeHtml(lead.website)}</a>` : 'Not recorded')}
      ${contactItem('Email', email ? `<a href="${escapeHtml(email)}">${escapeHtml(lead.email)}</a>` : 'Not recorded')}
      ${contactItem('Phone', escapeHtml(lead.contactNumber) || 'Not recorded')}
      ${contactItem('Social', socialHtml, 'social-contact')}
    </div>

    <label class="lead-status-control detail-status">
      <span>Lead Status</span>
      ${statusSelect(lead, 'leadStatus')}
    </label>

    <div class="stage-grid detail-stages">
      ${stageCard(lead, '1st Email', 'firstEmailMessage', 'firstEmailStatus', 'firstEmailDate', 'View 1st Email')}
      ${stageCard(lead, 'Follow Up 1', 'followUpEmail1Message', 'followUp1Status', 'followUp1Date', 'View Follow Up 1')}
      ${stageCard(lead, 'Follow Up 2', 'followUpEmail2Message', 'followUp2Status', 'followUp2Date', 'View Follow Up 2')}
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

function contactItem(label, value, className = '') {
  return `
    <div class="contact-item ${className}">
      <span>${label}</span>
      <strong>${value}</strong>
    </div>
  `;
}

function stageCard(lead, title, messageField, statusField, dateFieldName, buttonLabel) {
  return `
    <section class="stage-card">
      <div class="stage-heading">
        <span>${title}</span>
        ${messageButton(lead, messageField, buttonLabel)}
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
    markDirty();
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
    lead.leadStatus = 'Closed';
  }
}

function markDirty() {
  dirty = true;
  setStatus('Unsaved changes');
  if (dom.autoSaveToggle.checked) {
    window.clearTimeout(autosaveTimer);
    autosaveTimer = window.setTimeout(() => saveWorkbook({ quiet: true }), 700);
  }
}

function toWorkbookRows() {
  return leads.map(lead => [
    lead.companyName,
    lead.website,
    lead.email,
    lead.contactNumber,
    sanitizeSocialLinks(lead.socialMediaLinks),
    lead.firstEmailMessage,
    lead.firstEmailStatus,
    lead.firstEmailDate,
    addPricingCopyIfUseful(lead.followUpEmail1Message),
    lead.followUp1Status,
    lead.followUp1Date,
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

async function buildWorkbookArray() {
  const worksheet = XLSX.utils.aoa_to_sheet([REQUIRED_HEADERS, ...toWorkbookRows()]);
  worksheet['!cols'] = [
    18, 30, 30, 18, 24, 52, 18, 15, 52, 18, 15, 52, 18, 15, 20
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
    `<dataValidation type="list" allowBlank="0" showErrorMessage="1" sqref="G2:G1000 J2:J1000 M2:M1000"><formula1>${escapeXml(STATUS_VALIDATION_FORMULA)}</formula1></dataValidation>`,
    `<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="O2:O1000"><formula1>${escapeXml(LEAD_STATUS_VALIDATION_FORMULA)}</formula1></dataValidation>`,
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
    suggestedName: workbookName || FALLBACK_WORKBOOK_NAME,
    types: [{
      description: 'Excel workbooks',
      accept: {
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
      }
    }]
  });
  await saveHandle(handle);
  return handle;
}

async function saveWorkbook({ quiet = false } = {}) {
  if (!leads.length) return;
  let output = null;

  try {
    output = await buildWorkbookArray();

    if (fileHandle && await verifyFilePermission(fileHandle, 'readwrite')) {
      await writeWorkbookToHandle(fileHandle, output);
      dirty = false;
      setStatus(quiet ? 'Auto-saved' : 'Saved to Excel');
      return;
    }

    if (!quiet) {
      const pickedHandle = await chooseSaveHandle();
      if (pickedHandle) {
        await writeWorkbookToHandle(pickedHandle, output);
        fileHandle = pickedHandle;
        workbookName = pickedHandle.name || workbookName;
        dirty = false;
        dom.workbookChip.textContent = `Connected to ${workbookName}`;
        setStatus('Saved to Excel');
        return;
      }

      await downloadWorkbook(output);
      dirty = false;
      setStatus('Downloaded updated copy');
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

async function downloadWorkbook(output) {
  const workbookOutput = output || await buildWorkbookArray();
  const link = document.createElement('a');
  const safeName = workbookName.replace(/\.(xlsx|xls)$/i, '') || 'leads-tracker';
  link.href = URL.createObjectURL(new Blob([workbookOutput], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  link.download = `${safeName}-updated.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

async function openFile() {
  if ('showOpenFilePicker' in window) {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [{
        description: 'Excel workbooks',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
          'application/vnd.ms-excel': ['.xls']
        }
      }]
    });
    const file = await handle.getFile();
    await saveHandle(handle);
    readWorkbook(await file.arrayBuffer(), file.name, handle);
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

async function saveHandle(handle) {
  if (!('indexedDB' in window)) return;
  const db = await openDb();
  const tx = db.transaction(HANDLE_STORE, 'readwrite');
  tx.objectStore(HANDLE_STORE).put(handle, HANDLE_KEY);
}

async function getSavedHandle() {
  if (!('indexedDB' in window)) return null;
  const db = await openDb();
  return new Promise(resolve => {
    const tx = db.transaction(HANDLE_STORE, 'readonly');
    const request = tx.objectStore(HANDLE_STORE).get(HANDLE_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => resolve(null);
  });
}

async function loadSavedHandle() {
  if (!('showOpenFilePicker' in window)) return false;
  const handle = await getSavedHandle();
  if (!handle) return false;
  if (!await verifyFilePermission(handle, 'read')) return false;
  const file = await handle.getFile();
  readWorkbook(await file.arrayBuffer(), file.name, handle);
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
  dom.modalCompany.textContent = lead.companyName || 'Lead message';
  dom.modalTitle.textContent = labels[field];
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
  const file = event.target.files[0];
  if (!file) return;
  readWorkbook(await file.arrayBuffer(), file.name);
  event.target.value = '';
});

dom.saveButton.addEventListener('click', () => saveWorkbook());
dom.downloadButton.addEventListener('click', () => downloadWorkbook());
dom.searchInput.addEventListener('input', render);
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
    dom.copyMessageButton.textContent = 'Copy';
  }, 1000);
});

(async function init() {
  dom.autoSaveToggle.checked = true;
  try {
    if (await loadSavedHandle()) return;
  } catch (error) {
    console.warn('Saved file handle could not be restored.', error);
  }
  await loadDefaultWorkbook();
})();
