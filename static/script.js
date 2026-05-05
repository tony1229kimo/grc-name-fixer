// ============================================================
// GRC 團名修正工具 - 主頁面互動
// ============================================================

const state = {
    calendar: null,
    reference: null,
};

// ----- 初始化 -----
document.addEventListener('DOMContentLoaded', () => {
    setupDropzone('dropzone-calendar', 'calendar');
    setupDropzone('dropzone-reference', 'reference');

    document.getElementById('btn-process').addEventListener('click', doProcess);
    document.getElementById('btn-reset').addEventListener('click', reset);
});

// ----- Dropzone -----
function setupDropzone(id, field) {
    const zone = document.getElementById(id);
    const input = zone.querySelector('input[type="file"]');

    zone.addEventListener('click', (e) => {
        if (e.target.classList.contains('file-remove')) return;
        if (!zone.classList.contains('has-file')) input.click();
    });

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('drag-over');
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        if (zone.classList.contains('has-file')) return;
        const file = e.dataTransfer.files[0];
        if (file) handleFile(file, field);
    });

    input.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleFile(file, field);
    });

    // 移除檔案
    zone.querySelector('.file-remove').addEventListener('click', (e) => {
        e.stopPropagation();
        removeFile(field);
    });
}

function handleFile(file, field) {
    // 驗證副檔名
    if (!/\.xlsx?$/i.test(file.name)) {
        showError('請選擇 .xlsx 檔案');
        return;
    }
    if (file.size > 20 * 1024 * 1024) {
        showError('檔案過大（最大 20 MB）');
        return;
    }

    state[field] = file;
    const zone = document.getElementById(`dropzone-${field}`);
    zone.classList.add('has-file');
    zone.querySelector('.dropzone-file').hidden = false;
    zone.querySelector('.file-name').textContent = file.name;

    updateProcessButton();
}

function removeFile(field) {
    state[field] = null;
    const zone = document.getElementById(`dropzone-${field}`);
    zone.classList.remove('has-file');
    zone.querySelector('.dropzone-file').hidden = true;
    zone.querySelector('input[type="file"]').value = '';
    updateProcessButton();
}

function updateProcessButton() {
    const btn = document.getElementById('btn-process');
    btn.disabled = !state.calendar;
}

// ----- 處理 -----
async function doProcess() {
    if (!state.calendar) return;

    showOverlay(true, '處理中，請稍候...');

    // 1 分鐘逾時：避免永遠卡在轉圈圈
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 60000);

    // 顯示等待提示（較長時間用）
    const tipTimer = setTimeout(() => {
        setOverlayText('仍在處理中，大檔案可能需要較久時間...');
    }, 8000);

    const form = new FormData();
    form.append('calendar', state.calendar);
    if (state.reference) form.append('reference', state.reference);

    try {
        const res = await fetch('/api/process', {
            method: 'POST',
            body: form,
            signal: controller.signal,
        });

        // 先檢查 HTTP 狀態
        if (!res.ok) {
            let errMsg = `伺服器回應異常 (HTTP ${res.status})`;
            try {
                const errData = await res.json();
                if (errData.error) errMsg = errData.error;
            } catch (_) {
                // response 不是 JSON，忽略
            }
            showError(errMsg);
            return;
        }

        let data;
        try {
            data = await res.json();
        } catch (jsonErr) {
            showError('伺服器回應格式錯誤，請查看黑色視窗的錯誤訊息');
            return;
        }

        if (!data.ok) {
            showError(data.error || '處理失敗，請重試');
            return;
        }

        showResult(data);
    } catch (err) {
        if (err.name === 'AbortError') {
            showError('處理逾時（超過 60 秒）。請檢查檔案是否過大，或黑色視窗是否有錯誤訊息。');
        } else {
            showError('連線錯誤：' + err.message + '（請確認黑色視窗還開著）');
        }
    } finally {
        clearTimeout(timeoutId);
        clearTimeout(tipTimer);
        showOverlay(false);
    }
}

function showResult(data) {
    const s = data.summary;
    document.getElementById('stat-total').textContent = s.total;
    document.getElementById('stat-dict').textContent = s.from_dictionary;
    document.getElementById('stat-ref').textContent = s.from_reference;
    document.getElementById('stat-learned').textContent = s.newly_learned;
    document.getElementById('stat-unmatched').textContent = s.unmatched;

    const box = document.getElementById('result-unmatched-box');
    const list = document.getElementById('unmatched-list');
    list.innerHTML = '';

    if (s.unmatched_rows && s.unmatched_rows.length > 0) {
        box.hidden = false;
        for (const u of s.unmatched_rows) {
            const div = document.createElement('div');
            div.className = 'unmatched-item';
            div.innerHTML = `
                <span class="name">${escapeHtml(u.raw_name)}</span>
                <span class="meta">
                    ${escapeHtml(u.sheet || '').replace(/^InterContinental\s+/, '')} ·
                    R${u.row} ·
                    到達 ${u.arrival || '-'} ·
                    營收 ${formatMoney(u.revenue)}
                </span>
            `;
            list.appendChild(div);
        }
    } else {
        box.hidden = true;
    }

    // 下載連結
    const dl = document.getElementById('btn-download');
    dl.href = `/api/download/${encodeURIComponent(data.job_id)}`;
    dl.setAttribute('download', data.download_name || 'fixed.xlsx');

    // 顯示 panel
    document.getElementById('panel-result').hidden = false;
    document.getElementById('panel-result').scrollIntoView(
        { behavior: 'smooth', block: 'start' }
    );
}

function reset() {
    state.calendar = null;
    state.reference = null;
    for (const f of ['calendar', 'reference']) {
        const zone = document.getElementById(`dropzone-${f}`);
        zone.classList.remove('has-file');
        zone.querySelector('.dropzone-file').hidden = true;
        zone.querySelector('input[type="file"]').value = '';
    }
    document.getElementById('panel-result').hidden = true;
    updateProcessButton();
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ----- UI Helpers -----
function showOverlay(show, text) {
    document.getElementById('overlay').hidden = !show;
    if (show && text) setOverlayText(text);
}

function setOverlayText(text) {
    const el = document.querySelector('#overlay .overlay-text');
    if (el) el.textContent = text;
}

function showError(msg) {
    const toast = document.getElementById('toast-error');
    toast.querySelector('.toast-msg').textContent = msg;
    toast.hidden = false;
    setTimeout(() => { toast.hidden = true; }, 5000);
}

function closeToast() {
    document.querySelectorAll('.toast').forEach((t) => t.hidden = true);
}

function escapeHtml(s) {
    if (s == null) return '';
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function formatMoney(n) {
    if (n == null) return '-';
    return 'TWD ' + Number(n).toLocaleString(undefined, {
        maximumFractionDigits: 0,
    });
}
