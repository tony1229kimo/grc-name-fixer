// ============================================================
// GRC 團名修正工具 - 字典管理頁
// ============================================================

let allItems = [];
let editingAbbr = null;   // 編輯中的原始縮寫（null 表示新增）

document.addEventListener('DOMContentLoaded', () => {
    loadDict();

    document.getElementById('btn-add-open').addEventListener('click', openAddModal);
    document.getElementById('btn-save').addEventListener('click', saveDict);

    // 搜尋 (debounce)
    let timer;
    document.getElementById('dict-search').addEventListener('input', (e) => {
        clearTimeout(timer);
        timer = setTimeout(() => loadDict(e.target.value), 200);
    });

    // ESC 關閉 modal
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeAddModal();
    });
});

async function loadDict(q = '') {
    const url = q ? `/api/dictionary?q=${encodeURIComponent(q)}` : '/api/dictionary';
    try {
        const res = await fetch(url);
        const data = await res.json();
        allItems = data.items || [];
        document.getElementById('dict-total').textContent = data.count;
        renderTable(allItems);
    } catch (err) {
        showError('載入字典失敗：' + err.message);
    }
}

function renderTable(items) {
    const tbody = document.getElementById('dict-tbody');
    if (!items || items.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-row">（沒有資料）</td></tr>';
        return;
    }

    tbody.innerHTML = items.map((it) => `
        <tr>
            <td><span class="abbr-cell">${escapeHtml(it.abbreviation)}</span></td>
            <td>${escapeHtml(it.full_name)}</td>
            <td>
                <span class="source-badge source-${it.source}">
                    ${it.source === 'manual' ? '手動' : '自動'}
                </span>
            </td>
            <td>${it.usage_count || 0}</td>
            <td>${formatDateShort(it.last_used)}</td>
            <td>
                <button class="btn-delete" onclick="confirmDelete('${escapeJs(it.abbreviation)}')">刪除</button>
            </td>
        </tr>
    `).join('');
}

function openAddModal(editItem = null) {
    editingAbbr = null;
    document.getElementById('modal-title').textContent = '新增字典對應';
    document.getElementById('input-abbr').value = '';
    document.getElementById('input-full').value = '';
    document.getElementById('input-note').value = '';
    document.getElementById('input-abbr').removeAttribute('readonly');

    if (editItem && editItem.abbreviation) {
        editingAbbr = editItem.abbreviation;
        document.getElementById('modal-title').textContent = '編輯字典對應';
        document.getElementById('input-abbr').value = editItem.abbreviation;
        document.getElementById('input-abbr').setAttribute('readonly', '');
        document.getElementById('input-full').value = editItem.full_name || '';
        document.getElementById('input-note').value = editItem.note || '';
    }

    document.getElementById('modal-add').hidden = false;
    setTimeout(() => document.getElementById('input-abbr').focus(), 50);
}

function closeAddModal() {
    document.getElementById('modal-add').hidden = true;
}

async function saveDict() {
    const abbr = document.getElementById('input-abbr').value.trim();
    const full = document.getElementById('input-full').value.trim();
    const note = document.getElementById('input-note').value.trim();

    if (!abbr) {
        showError('請填入縮寫');
        return;
    }
    if (!full) {
        showError('請填入完整名稱');
        return;
    }

    try {
        const res = await fetch('/api/dictionary', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ abbreviation: abbr, full_name: full, note }),
        });
        const data = await res.json();
        if (!data.ok) {
            showError(data.error || '儲存失敗');
            return;
        }
        showSuccess(
            data.status === 'added' ? '已新增對應' :
            data.status === 'updated' ? '已更新對應' : '無變更'
        );
        closeAddModal();
        loadDict(document.getElementById('dict-search').value);
    } catch (err) {
        showError('儲存失敗：' + err.message);
    }
}

async function confirmDelete(abbr) {
    if (!confirm(`確定要刪除這個對應？\n\n縮寫：${abbr}`)) return;

    try {
        const res = await fetch(
            `/api/dictionary/${encodeURIComponent(abbr)}`,
            { method: 'DELETE' }
        );
        const data = await res.json();
        if (data.ok) {
            showSuccess('已刪除');
            loadDict(document.getElementById('dict-search').value);
        } else {
            showError('刪除失敗');
        }
    } catch (err) {
        showError('刪除失敗：' + err.message);
    }
}

// ----- Helpers -----
function showError(msg) {
    const t = document.getElementById('toast-error');
    t.querySelector('.toast-msg').textContent = msg;
    t.hidden = false;
    setTimeout(() => { t.hidden = true; }, 5000);
}

function showSuccess(msg) {
    const t = document.getElementById('toast-success');
    t.querySelector('.toast-msg').textContent = msg;
    t.hidden = false;
    setTimeout(() => { t.hidden = true; }, 3000);
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

function escapeJs(s) {
    if (s == null) return '';
    return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function formatDateShort(iso) {
    if (!iso) return '-';
    // iso 從 SQLite 格式為 "YYYY-MM-DD HH:MM:SS"
    const parts = iso.split(' ');
    if (parts.length >= 1) return parts[0];
    return iso;
}
