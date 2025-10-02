const exportButton = document.getElementById('export-btn');
const exportFormatSelect = /** @type {HTMLSelectElement} */ (document.getElementById('export-format'));
exportButton.disabled = true;
exportFormatSelect.disabled = true;
const statusEl = document.getElementById('status');
const previewSection = document.getElementById('preview');
const summaryEl = document.getElementById('summary');
const previewTableContainer = document.getElementById('preview-table');
const configGroupEl = document.getElementById('config-group');
const configStructureEl = document.getElementById('config-structure');

// Optional fallback serials if auto-detection cannot resolve values from the active tab.
const DEFAULT_SERIALS = {
  groupSerial: '',
  structureSerial: ''
};

let activeSerials = {
  groupSerial: DEFAULT_SERIALS.groupSerial,
  structureSerial: DEFAULT_SERIALS.structureSerial
};

updateConfigDisplay({
  groupSerial: activeSerials.groupSerial,
  structureSerial: activeSerials.structureSerial,
  groupDetected: false,
  structureDetected: false
});

const state = {
  userGroups: [],
  userSubmissions: [],
  studentSubmissions: []
};

let isFetching = false;

const API_HEADERS = {
  'content-type': 'application/json',
  tenantname: 'uob',
  country: 'id',
  platform: 'Web',
  'with-auth': 'true'
};

function updateConfigDisplay({ groupSerial, structureSerial, groupDetected, structureDetected }) {
  configGroupEl.textContent = groupSerial
    ? `${groupSerial}${groupDetected ? '' : ' (default)'}`
    : 'Not detected';

  configStructureEl.textContent = structureSerial
    ? `${structureSerial}${structureDetected ? '' : ' (default)'}`
    : 'Not detected';
}

async function fetchAndRender() {
  if (isFetching) {
    return;
  }

  const { groupSerial, structureSerial } = activeSerials;

  if (!groupSerial || !structureSerial) {
    setStatus('Unable to resolve serials. Keep the assignment page active or configure defaults in popup.js.', 'error');
    return;
  }

  isFetching = true;

  exportButton.disabled = true;
  exportFormatSelect.disabled = true;
  previewSection.hidden = true;
  setStatus('Fetching user group data…');

  try {
    const userGroups = await fetchAllUserGroups(groupSerial);

    if (!userGroups.length) {
      throw new Error('No users found in the specified group.');
    }

    setStatus(`Fetched ${userGroups.length} users. Fetching submissions…`);

    const userSerials = userGroups
      .map(group => group.userSerial)
      .filter(Boolean);

    const userSubmissions = userSerials.length
      ? await fetchUserSubmissions(structureSerial, userSerials)
      : [];

    const studentSubmissions = groupSubmissionsByUser(userGroups, userSubmissions);

    state.userGroups = userGroups;
    state.userSubmissions = userSubmissions;
    state.studentSubmissions = studentSubmissions;

    renderPreview(studentSubmissions);
    exportButton.disabled = false;
    exportFormatSelect.disabled = false;
    setStatus('Data ready. Review the preview and export when ready.');
  } catch (error) {
    console.error(error);
    setStatus(error.message || 'Unable to fetch data. Please retry.', 'error');
    if (!state.studentSubmissions.length) {
      exportButton.disabled = true;
      exportFormatSelect.disabled = true;
      previewTableContainer.replaceChildren();
      previewSection.hidden = true;
    }
  } finally {
    const hasData = state.studentSubmissions.length > 0;
    exportButton.disabled = !hasData;
    exportFormatSelect.disabled = !hasData;
    isFetching = false;
  }
}

async function detectSerials() {
  try {
    const tabs = await new Promise((resolve, reject) => {
      chrome.tabs.query({ active: true, currentWindow: true }, queriedTabs => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
          return;
        }
        resolve(queriedTabs);
      });
    });

    const [tab] = tabs;
    if (!tab?.id) {
      return {};
    }

    if (tab.url && !tab.url.startsWith('https://cms.uobmydigitalspace.com')) {
      return {};
    }

    const [injectionResult] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => {
        const url = new URL(window.location.href);
        let structureSerial = url.searchParams.get('serial');

        const groupParamCandidates = ['groupSerial', 'groupSerials', 'group'];
        let groupSerial = null;
        for (const key of groupParamCandidates) {
          const value = url.searchParams.get(key);
          if (value) {
            groupSerial = value;
            break;
          }
        }

        if (!groupSerial) {
          const resources = performance.getEntriesByType('resource') || [];
          for (let index = resources.length - 1; index >= 0; index -= 1) {
            const entry = resources[index];
            const name = typeof entry.name === 'string' ? entry.name : '';
            if (!name) {
              continue;
            }

            try {
              const resourceUrl = new URL(name);
              for (const key of groupParamCandidates) {
                const value = resourceUrl.searchParams.get(key);
                if (value) {
                  groupSerial = value;
                  break;
                }
              }
            } catch (error) {
              const match = name.match(/groupSerials=([A-Z0-9-]+)/);
              if (match) {
                groupSerial = match[1];
              }
            }

            if (groupSerial) {
              break;
            }
          }
        }

        if (!groupSerial || !structureSerial) {
          const bodyText = document.body ? document.body.innerHTML : '';
          if (!structureSerial) {
            const structureMatch = bodyText.match(/Node-[A-Z0-9]+/);
            if (structureMatch) {
              structureSerial = structureMatch[0];
            }
          }

          if (!groupSerial) {
            const groupMatch = bodyText.match(/GRP-[A-Z0-9]+/);
            if (groupMatch) {
              groupSerial = groupMatch[0];
            }
          }
        }

        return {
          structureSerial: structureSerial || null,
          groupSerial: groupSerial || null
        };
      }
    });

    const result = injectionResult?.result || {};
    return {
      structureSerial: result.structureSerial || null,
      groupSerial: result.groupSerial || null
    };
  } catch (error) {
    console.error('detectSerials', error);
    return {};
  }
}

exportButton.addEventListener('click', () => {
  try {
    const format = exportFormatSelect.value;
    if (format === 'pdf') {
      exportPdfReport(state);
    } else {
      exportHtmlReport(state);
    }
  } catch (error) {
    console.error(error);
    setStatus('Failed to export report. Check console for details.', 'error');
  }
});

function setStatus(message, mode = 'info') {
  statusEl.textContent = message;
  statusEl.dataset.mode = mode;
}

async function fetchAllUserGroups(groupSerial) {
  const results = [];
  let currentPage = 1;
  let totalPage = 1;

  while (currentPage <= totalPage) {
    const url = new URL('https://cms.uobmydigitalspace.com/api/v3/user-group/group/undefined/users');
    url.searchParams.set('page', String(currentPage));
    url.searchParams.set('pageSize', '50');
    url.searchParams.set('groupSerials', groupSerial);
    url.searchParams.set('withUserData', 'true');
    url.searchParams.set('status', 'USER_GROUP_ACTIVE_BY_PERIOD_FILTER');
    url.searchParams.set('userId', '');
    url.searchParams.set('userSerial', '');

    const response = await fetch(url.toString(), {
      method: 'GET',
      headers: API_HEADERS,
      credentials: 'include'
    });

    if (!response.ok) {
      throw new Error(`User group request failed with status ${response.status}`);
    }

    const payload = await response.json();
    const pageGroups = payload?.data?.userGroups ?? [];
    const pagination = payload?.data?.pagination ?? {};

    totalPage = Number(pagination.totalPage) || 1;
    results.push(...pageGroups);
    currentPage += 1;
  }

  return results;
}

async function fetchUserSubmissions(structureSerial, userSerials) {
  const results = [];
  const chunkSize = 18;

  for (let index = 0; index < userSerials.length; index += chunkSize) {
    const chunk = userSerials.slice(index, index + chunkSize);
    const url = new URL('https://cms.uobmydigitalspace.com/api/v3/exam/user-submission/list');
    url.searchParams.set('structureSerial', structureSerial);
    url.searchParams.set('tenant', 'uob');
    chunk.forEach(serial => url.searchParams.append('userSerials', serial));

    const response = await fetch(url.toString(), {
      method: 'GET',
      headers: API_HEADERS,
      credentials: 'include'
    });

    if (!response.ok) {
      throw new Error(`Submission request failed with status ${response.status}`);
    }

    const payload = await response.json();
    const submissions = payload?.data?.userSubmissions ?? [];
    results.push(...submissions);
  }

  return results;
}

function groupSubmissionsByUser(userGroups, userSubmissions) {
  const groupMap = new Map(userGroups.map(group => [group.userSerial, group]));
  const byUser = new Map();

  userSubmissions.forEach(submission => {
    const group = groupMap.get(submission.userSerial);
    const studentRecord = byUser.get(submission.userSerial) || {
      userSerial: submission.userSerial,
      groupSerial: group?.groupSerial || '',
      name: group?.name || group?.userName || 'Unknown Student',
      email: group?.email || '',
      submissions: []
    };

    studentRecord.submissions.push({
      description: submission.description || '',
      score: submission.score,
      submittedAt: submission.submittedAt || '',
      attachments: submission.submissions || [],
      feedback: submission.feedback || ''
    });

    byUser.set(submission.userSerial, studentRecord);
  });

  const result = Array.from(byUser.values());

  result.forEach(student => {
    student.submissions.sort((a, b) => {
      const aTime = a.submittedAt ? new Date(a.submittedAt).getTime() : 0;
      const bTime = b.submittedAt ? new Date(b.submittedAt).getTime() : 0;
      return bTime - aTime;
    });
  });

  result.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

  return result;
}

function renderPreview(studentSubmissions) {
  if (!studentSubmissions.length) {
    setStatus('No submissions found for this selection.', 'error');
    previewTableContainer.replaceChildren();
    previewSection.hidden = true;
    return;
  }

  const studentCount = studentSubmissions.length;
  const submissionCount = studentSubmissions.reduce((total, student) => total + student.submissions.length, 0);
  const fileCount = studentSubmissions.reduce((total, student) => {
    return total + student.submissions.reduce((subtotal, submission) => subtotal + submission.attachments.length, 0);
  }, 0);

  summaryEl.textContent = `${studentCount} students · ${submissionCount} submission entries · ${fileCount} files.`;

  const table = document.createElement('table');
  const header = document.createElement('tr');

  ['Student', 'Submissions', 'Score', 'Feedback'].forEach(label => {
    const th = document.createElement('th');
    th.textContent = label;
    header.appendChild(th);
  });

  table.appendChild(header);

  studentSubmissions.slice(0, 5).forEach(student => {
    const tr = document.createElement('tr');

    const latestSubmission = student.submissions[0];

    const cells = [
      student.name || '—',
      String(student.submissions.length),
      latestSubmission?.score !== null && latestSubmission?.score !== undefined
        ? String(latestSubmission.score)
        : '—',
      latestSubmission?.feedback
        ? truncateText(latestSubmission.feedback, 60)
        : '—'
    ];

    cells.forEach(value => {
      const td = document.createElement('td');
      td.textContent = value;
      tr.appendChild(td);
    });

    table.appendChild(tr);
  });

  previewTableContainer.replaceChildren(table);
  previewSection.hidden = false;
}

function exportHtmlReport({ studentSubmissions }) {
  if (!studentSubmissions.length) {
    throw new Error('No submission data available. Fetch data before exporting.');
  }

  const html = buildHtmlDocument(studentSubmissions);
  const fileName = `uob-submissions-${new Date().toISOString().slice(0, 10)}.html`;
  downloadFile(fileName, html);
  setStatus('HTML report exported successfully.');
}

function exportPdfReport({ studentSubmissions }) {
  if (!studentSubmissions.length) {
    throw new Error('No submission data available. Fetch data before exporting.');
  }

  const html = buildHtmlDocument(studentSubmissions, {
    includePrintStyles: true,
    autoPrint: true
  });

  const blob = new Blob([html], { type: 'text/html' });
  const blobUrl = URL.createObjectURL(blob);
  const reportWindow = window.open(blobUrl, '_blank');

  if (!reportWindow) {
    URL.revokeObjectURL(blobUrl);
    throw new Error('Unable to open PDF preview window. Allow pop-ups and try again.');
  }

  const revokeUrl = () => {
    URL.revokeObjectURL(blobUrl);
  };

  try {
    reportWindow.addEventListener('beforeunload', revokeUrl, { once: true });
  } catch (error) {
    console.warn('Unable to attach unload handler for PDF window.', error);
    setTimeout(revokeUrl, 60_000);
  }

  setStatus('PDF preview opened in a new tab. Use the browser print dialog to save as PDF.');
}

function buildHtmlDocument(studentSubmissions, options = {}) {
  const { includePrintStyles = false, autoPrint = false } = options;
  const submissionCount = studentSubmissions.reduce((total, student) => total + student.submissions.length, 0);
  const fileCount = studentSubmissions.reduce((total, student) => {
    return total + student.submissions.reduce((subtotal, submission) => subtotal + submission.attachments.length, 0);
  }, 0);
  const generatedAt = new Date().toLocaleString();

  const rows = studentSubmissions
    .map(student => renderStudentRow(student))
    .join('\n');

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>UOB Assignment Submissions Report</title>
    <style>
      :root {
        color-scheme: light;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      body {
        margin: 0;
        padding: 24px;
        background: #f5f7fa;
      }
      header {
        margin-bottom: 24px;
      }
      h1 {
        margin: 0 0 8px;
        font-size: 22px;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        background: white;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
      }
      thead {
        background: #005ea8;
        color: white;
      }
      th, td {
        padding: 12px;
        text-align: left;
        vertical-align: top;
        border-bottom: 1px solid rgba(0, 0, 0, 0.08);
      }
      tr:last-child td {
        border-bottom: none;
      }
      .muted {
        color: rgba(0, 0, 0, 0.6);
        font-style: italic;
      }
      .student-meta {
        display: flex;
        flex-direction: column;
        gap: 4px;
      }
      .student-meta small {
        color: rgba(0, 0, 0, 0.6);
      }
      .submission-list {
        display: flex;
        flex-direction: column;
        gap: 12px;
      }
      .submission-card {
        border: 1px solid rgba(0, 0, 0, 0.1);
        border-radius: 6px;
        padding: 12px;
        background: rgba(0, 0, 0, 0.02);
      }
      .submission-card header {
        margin: 0 0 8px;
        font-weight: 600;
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
      }
      .submission-card header span {
        display: flex;
        gap: 6px;
      }
      .submission-card header strong {
        font-weight: 700;
      }
      .attachments {
        margin-top: 8px;
        display: grid;
        gap: 12px;
      }
      .attachments img {
        max-width: 260px;
        width: 100%;
        border: 1px solid rgba(0, 0, 0, 0.1);
        border-radius: 4px;
      }
      summary {
        cursor: pointer;
      }
      summary::marker {
        color: #005ea8;
      }
      a {
        color: #005ea8;
      }
      ${includePrintStyles ? '@page { margin: 12mm; } @media print { body { background: white; padding: 12mm; } table { box-shadow: none; } }' : ''}
    </style>
  </head>
  <body>
    <header>
      <h1>UOB Assignment Submissions Report</h1>
      <p>Generated at ${escapeHtml(generatedAt)} · ${studentSubmissions.length} students · ${submissionCount} submission entries · ${fileCount} files</p>
    </header>
    <table>
      <thead>
        <tr>
          <th style="width: 25%">Student</th>
          <th style="width: 10%">Submissions</th>
          <th style="width: 15%">Score</th>
          <th style="width: 30%">Feedback</th>
        </tr>
      </thead>
      <tbody>
${rows}
      </tbody>
    </table>
    ${autoPrint ? '<script>window.addEventListener("load",()=>{window.focus();window.print();});</script>' : ''}
  </body>
</html>`;
}

function renderStudentRow(student) {
  const submissionCount = student.submissions.length;
  if (!submissionCount) {
    return '';
  }

  const latestSubmission = student.submissions[0];
  const submissionsHtml = renderSubmissionList(student);
  const scoreMarkup = latestSubmission?.score !== null && latestSubmission?.score !== undefined
    ? escapeHtml(String(latestSubmission.score))
    : '<span class="muted">No score</span>';
  const feedbackMarkup = latestSubmission?.feedback
    ? formatMultiline(latestSubmission.feedback)
    : '<span class="muted">No feedback</span>';

  return `        <tr>
          <td>
            <div class="student-meta">
              <strong>${escapeHtml(student.name)}</strong>
              ${student.email ? `<small>${escapeHtml(student.email)}</small>` : ''}
              <small class="muted">User Serial: ${escapeHtml(student.userSerial)}</small>
              ${student.groupSerial ? `<small class="muted">Group Serial: ${escapeHtml(student.groupSerial)}</small>` : ''}
            </div>
          </td>
          <td>${submissionsHtml}</td>
          <td>${scoreMarkup}</td>
          <td>${feedbackMarkup}</td>
        </tr>`;
}

function renderSubmissionList(student) {
  const submissionCount = student.submissions.length;

  if (submissionCount === 1) {
    return renderSubmissionCard(student.submissions[0], student.name, 1);
  }

  const cards = student.submissions
    .map((submission, index) => renderSubmissionCard(submission, student.name, index + 1))
    .join('');

  return `<details open>
            <summary>${submissionCount} submissions</summary>
            <div class="submission-list">${cards}</div>
          </details>`;
}

function renderSubmissionCard(submission, studentName, index) {
  const submittedAt = submission.submittedAt
    ? escapeHtml(new Date(submission.submittedAt).toLocaleString())
    : '<span class="muted">Not submitted</span>';
  const description = submission.description
    ? formatMultiline(submission.description)
    : '<span class="muted">No description provided.</span>';
  const attachments = renderAttachments(submission.attachments, studentName, index);

  return `<article class="submission-card">
            <header>
              <span><strong>Submission ${index}</strong></span>
              <span>Submitted: ${submittedAt}</span>
            </header>
            <div>${description}</div>
            ${attachments}
          </article>`;
}

function renderAttachments(attachments, studentName, submissionIndex) {
  if (!attachments || !attachments.length) {
    return '<div class="attachments"><span class="muted">No files</span></div>';
  }

  const items = attachments
    .map((attachment, index) => renderAttachmentContent(attachment, studentName, submissionIndex, index + 1))
    .join('');

  return `<div class="attachments">${items}</div>`;
}

function renderAttachmentContent(attachment, studentName, submissionIndex, fileIndex) {
  const url = attachment.URL || '';
  const safeUrl = escapeAttribute(url);
  const label = `${attachment.type || 'File'} ${submissionIndex}.${fileIndex}`;

  if (isImageType(attachment.type)) {
    return `<div>
              <div><a href="${safeUrl}" target="_blank" rel="noopener">${escapeHtml(label)}</a></div>
              <img src="${safeUrl}" alt="${escapeAttribute(`${studentName} submission image ${submissionIndex}.${fileIndex}`)}" loading="lazy" />
            </div>`;
  }

  if (isLinkType(attachment.type)) {
    return `<div>${escapeHtml(label)}: <a href="${safeUrl}" target="_blank" rel="noopener">Open link</a></div>`;
  }

  return `<div>${escapeHtml(label)}: <a href="${safeUrl}" target="_blank" rel="noopener">Download file</a></div>`;
}

function isImageType(type) {
  return typeof type === 'string' && /PNG|JPE?G|GIF|WEBP/i.test(type);
}

function isLinkType(type) {
  return String(type || '').toUpperCase() === 'SUBMISSION_TYPE_LINK';
}

function formatMultiline(text) {
  return escapeHtml(text).replace(/\n/g, '<br />');
}

function truncateText(value, maxLength) {
  if (!value) {
    return '';
  }

  const trimmed = value.trim();
  if (trimmed.length <= maxLength) {
    return trimmed;
  }

  return `${trimmed.slice(0, maxLength - 1)}…`;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, char => {
    switch (char) {
      case '&':
        return '&amp;';
      case '<':
        return '&lt;';
      case '>':
        return '&gt;';
      case '"':
        return '&quot;';
      case "'":
        return '&#39;';
      default:
        return char;
    }
  });
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/`/g, '&#96;');
}

function downloadFile(fileName, html) {
  const blob = new Blob([html], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function initialize() {
  setStatus('Detecting assignment context…');
  const detected = await detectSerials();

  const groupSerial = detected.groupSerial || DEFAULT_SERIALS.groupSerial;
  const structureSerial = detected.structureSerial || DEFAULT_SERIALS.structureSerial;

  activeSerials = { groupSerial, structureSerial };

  updateConfigDisplay({
    groupSerial,
    structureSerial,
    groupDetected: Boolean(detected.groupSerial),
    structureDetected: Boolean(detected.structureSerial)
  });

  if (!groupSerial || !structureSerial) {
    setStatus('Unable to detect serials from the current tab. Keep the assignment page active or set DEFAULT_SERIALS in popup.js.', 'error');
    return;
  }

  fetchAndRender().catch(error => {
    console.error(error);
    setStatus(error.message || 'Unable to fetch data automatically.');
  });
}

initialize();
