/* ============================================
   VA/CA Report Automation - App Logic
   ============================================ */

// ---- Auth ----
function getToken() {
    return localStorage.getItem('token');
}

function isLoggedIn() {
    return !!getToken();
}

function checkAuth() {
    if (!isLoggedIn()) {
        window.location.href = '/login';
    }
}

function logout() {
    localStorage.removeItem('token');
    localStorage.removeItem('username');
    window.location.href = '/login';
}

// ---- Login ----
async function handleLogin(e) {
    e.preventDefault();
    const btn = e.target.querySelector('button[type="submit"]');
    const btnText = btn.querySelector('.btn-text');
    const btnSpinner = btn.querySelector('.btn-spinner');
    const errorDiv = document.getElementById('errorMessage');

    btnText.style.display = 'none';
    btnSpinner.style.display = 'inline-block';
    btn.disabled = true;
    errorDiv.style.display = 'none';

    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;

    try {
        const response = await fetch('/api/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username, password })
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.detail || 'Login failed');
        }

        localStorage.setItem('token', data.access_token);
        localStorage.setItem('username', username);
        window.location.href = '/';
    } catch (err) {
        errorDiv.textContent = err.message;
        errorDiv.style.display = 'block';
    } finally {
        btnText.style.display = 'inline';
        btnSpinner.style.display = 'none';
        btn.disabled = false;
    }
}

// ============================================
// STEP 1: Merge Files
// ============================================

let mergeFilesList = [];

function initMergeDropZone() {
    const dropZone = document.getElementById('mergeDropZone');
    const fileInput = document.getElementById('mergeFileInput');

    if (!dropZone) return;

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        handleMergeFiles(e.dataTransfer.files);
    });

    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        handleMergeFiles(e.target.files);
    });
}

function handleMergeFiles(files) {
    for (const file of files) {
        if (file.name.endsWith('.xlsx') || file.name.endsWith('.csv')) {
            if (!mergeFilesList.find(f => f.name === file.name)) {
                mergeFilesList.push(file);
            }
        }
    }
    updateMergeFileList();
    updateMergeButton();
}

function removeMergeFile(index) {
    mergeFilesList.splice(index, 1);
    updateMergeFileList();
    updateMergeButton();
}

function updateMergeFileList() {
    const fileList = document.getElementById('mergeFileList');
    if (!fileList) return;

    fileList.innerHTML = mergeFilesList.map((file, i) => `
        <div class="file-item">
            <span class="file-item-name">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/>
                </svg>
                ${file.name}
            </span>
            <button onclick="removeMergeFile(${i})" class="file-item-remove">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        </div>
    `).join('');
}

function updateMergeButton() {
    const mergeBtn = document.getElementById('mergeBtn');
    if (mergeBtn) {
        mergeBtn.disabled = mergeFilesList.length < 2;
    }
}

async function mergeFiles() {
    if (mergeFilesList.length < 2) return;

    const btn = document.getElementById('mergeBtn');
    setLoading(btn, true);

    try {
        const formData = new FormData();
        mergeFilesList.forEach(file => formData.append('files', file));

        const response = await fetch('/api/merge-csv', {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${getToken()}` },
            body: formData
        });

        if (!response.ok) {
            const data = await response.json();
            throw new Error(data.detail || 'Merge failed');
        }

        const blob = await response.blob();
        downloadBlob(blob, 'merged_raw.xlsx');
        showNotification('Files merged successfully! Download the file and use it in Step 2 or 3.', 'success');
    } catch (err) {
        showNotification(err.message, 'error');
    } finally {
        setLoading(btn, false);
    }
}

// ============================================
// STEP 2: Generate Excel Reports
// ============================================

function initExcelFileInput() {
    const fileInput = document.getElementById('excelFileInput');
    if (!fileInput) return;

    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            document.getElementById('excelFileName').textContent = file.name;
        }
    });
}

function buildFormData() {
    const form = document.getElementById('reportForm');
    const formData = new FormData(form);
    return formData;
}

async function generateExcelReports() {
    const fileInput = document.getElementById('excelFileInput');
    if (!fileInput.files[0]) {
        showNotification('Please select a merged file first', 'error');
        return;
    }

    const btn = document.getElementById('generateExcelBtn');
    setLoading(btn, true);

    try {
        const formData = buildFormData();
        formData.append('file', fileInput.files[0]);

        const response = await fetch('/api/report', {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${getToken()}` },
            body: formData
        });

        if (!response.ok) {
            const data = await response.json();
            throw new Error(data.detail || 'Report generation failed');
        }

        const data = await response.json();
        showExcelDownloadButtons(data.files);
        showNotification('Excel reports generated successfully!', 'success');
    } catch (err) {
        showNotification(err.message, 'error');
    } finally {
        setLoading(btn, false);
    }
}

function showExcelDownloadButtons(files) {
    const section = document.getElementById('downloadSection');
    const list = document.getElementById('downloadList');

    if (!section || !list) return;

    section.style.display = 'block';

    const labels = {
        'va_normal': 'VA Normal Report',
        'va_textjoin': 'VA TextJoin Report',
        'ca_normal': 'CA Normal Report',
        'ca_textjoin': 'CA TextJoin Report'
    };

    let items = '';
    for (const [key, url] of Object.entries(files)) {
        const label = labels[key] || key;
        items += `
            <div class="download-item">
                <div class="download-item-info">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                    <span class="download-item-name">${label}</span>
                </div>
                <button onclick="downloadFile('${url}')" class="btn btn-cta btn-sm">
                    Download .xlsx
                </button>
            </div>
        `;
    }

    list.innerHTML = items;
}

// ============================================
// STEP 3: Generate Word Report
// ============================================

let wordFilesList = [];

function initWordFileInput() {
    const fileInput = document.getElementById('wordFileInput');
    if (!fileInput) return;

    fileInput.addEventListener('change', (e) => {
        handleWordFiles(e.target.files);
    });
}

function handleWordFiles(files) {
    for (const file of files) {
        if (file.name.endsWith('.xlsx')) {
            if (!wordFilesList.find(f => f.name === file.name)) {
                wordFilesList.push(file);
            }
        }
    }
    updateWordFileList();
}

function removeWordFile(index) {
    wordFilesList.splice(index, 1);
    updateWordFileList();
}

function updateWordFileList() {
    const fileList = document.getElementById('wordFileList');
    if (!fileList) return;

    fileList.innerHTML = wordFilesList.map((file, i) => `
        <div class="file-item">
            <span class="file-item-name">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/>
                </svg>
                ${file.name}
            </span>
            <button onclick="removeWordFile(${i})" class="file-item-remove">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        </div>
    `).join('');
}

async function generateWordReport() {
    if (wordFilesList.length === 0) {
        showNotification('Please select at least one scan file', 'error');
        return;
    }

    const btn = document.getElementById('generateWordBtn');
    setLoading(btn, true);

    try {
        const formData = buildFormData();
        for (const file of wordFilesList) {
            formData.append('files', file);
        }

        const response = await fetch('/api/word', {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${getToken()}` },
            body: formData
        });

        if (!response.ok) {
            const data = await response.json();
            const detail = data.detail;
            const msg = Array.isArray(detail)
                ? detail.map(e => e.msg || e.loc?.join(' ') || JSON.stringify(e)).join('; ')
                : (detail || 'Word report generation failed');
            throw new Error(msg);
        }

        const data = await response.json();
        showWordDownloadButton(data.files);
        showNotification('Word report generated successfully!', 'success');
    } catch (err) {
        showNotification(err.message, 'error');
    } finally {
        setLoading(btn, false);
    }
}

function showWordDownloadButton(files) {
    const section = document.getElementById('downloadSection');
    const list = document.getElementById('downloadList');

    if (!section || !list) return;

    section.style.display = 'block';

    let items = list.innerHTML;
    for (const [key, url] of Object.entries(files)) {
        items += `
            <div class="download-item">
                <div class="download-item-info">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14 2 14 8 20 8"/>
                    </svg>
                    <span class="download-item-name">Word Report</span>
                </div>
                <button onclick="downloadFile('${url}')" class="btn btn-cta btn-sm">
                    Download .docx
                </button>
            </div>
        `;
    }

    list.innerHTML = items;
}

// ============================================
// Downloads
// ============================================

async function downloadFile(url) {
    try {
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${getToken()}` }
        });

        if (!response.ok) {
            throw new Error('Download failed');
        }

        const blob = await response.blob();
        const filename = url.split('/').pop() + (url.includes('word') ? '.docx' : '.xlsx');
        downloadBlob(blob, filename);
    } catch (err) {
        showNotification(err.message, 'error');
    }
}

function downloadBlob(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
}

// ============================================
// UI Helpers
// ============================================

function handleReportTypeChange() {
    const reportType = document.getElementById('reportType').value;
    const retestDates = document.getElementById('retestDates');
    const docTitle = document.getElementById('docTitle');

    if (retestDates) {
        retestDates.style.display = reportType === 'Final' ? 'grid' : 'none';
    }
    if (docTitle) {
        docTitle.value = reportType === 'Final' ? 'Final Audit Report' : 'First Audit Report';
    }
}

function setLoading(btn, loading) {
    if (!btn) return;
    const text = btn.querySelector('.btn-text');
    const spinner = btn.querySelector('.btn-spinner');

    if (loading) {
        btn.disabled = true;
        if (text) text.style.display = 'none';
        if (spinner) spinner.style.display = 'inline-block';
    } else {
        btn.disabled = false;
        if (text) text.style.display = 'inline';
        if (spinner) spinner.style.display = 'none';
    }
}

function showNotification(message, type) {
    const notification = document.getElementById('notification');
    if (!notification) return;

    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';

    setTimeout(() => {
        notification.style.display = 'none';
    }, 8000);
}

// ============================================
// Init
// ============================================
document.addEventListener('DOMContentLoaded', () => {
    initMergeDropZone();
    initExcelFileInput();
    initWordFileInput();
    handleReportTypeChange();
});
