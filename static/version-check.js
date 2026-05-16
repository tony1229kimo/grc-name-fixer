// ============================================================
// 新版本檢查 - 兩個頁面都會載入
// - 抓 /api/version (後端 4 小時快取一次,打 GitHub Releases)
// - 有新版 → 注入頂部 banner,提供「立即下載」按鈕
// - 在 footer 顯示目前版本號(無論有沒有更新)
// - 使用者按 ✕ 可以關掉,該版本通知存 localStorage 不再跳
// ============================================================
(function () {
    function escapeHtml(s) {
        return String(s).replace(/[&<>"']/g, function (c) {
            return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c];
        });
    }

    async function init() {
        let data;
        try {
            const res = await fetch('/api/version');
            if (!res.ok) return;
            data = await res.json();
        } catch (_) {
            return;  // 網路失敗就靜默,不打擾使用者
        }

        // footer 顯示版本號
        const footerSmall = document.querySelector('.app-footer small');
        if (footerSmall && data.current) {
            footerSmall.textContent += ' · v' + data.current;
        }

        if (!data.has_update || !data.latest) return;

        // 使用者已經關過這個版本的通知就不再跳
        try {
            if (localStorage.getItem('grc-update-dismissed') === data.latest) return;
        } catch (_) { /* 隱私模式可能沒 localStorage,忽略 */ }

        const banner = document.createElement('div');
        banner.className = 'update-banner';
        banner.innerHTML =
            '<div class="update-banner-inner">' +
                '<span class="update-banner-icon">🔔</span>' +
                '<span class="update-banner-text">' +
                    '<strong>新版本 ' + escapeHtml(data.latest) + ' 可下載</strong> ' +
                    '(目前 v' + escapeHtml(data.current) + ')' +
                '</span>' +
                '<a href="' + escapeHtml(data.release_url) + '" target="_blank" rel="noopener" class="update-banner-btn">' +
                    '立即前往下載 →' +
                '</a>' +
                '<button class="update-banner-close" aria-label="關閉這個通知" title="關閉(下次仍會檢查)">✕</button>' +
            '</div>';
        document.body.insertBefore(banner, document.body.firstChild);

        banner.querySelector('.update-banner-close').addEventListener('click', function () {
            banner.remove();
            try {
                localStorage.setItem('grc-update-dismissed', data.latest);
            } catch (_) { /* 隱私模式 */ }
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
