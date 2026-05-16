# GRC 團名修正工具

> 台中勤美洲際酒店 · 業務部 GRC 報表自動修正工具
> Delphi 跑出來的 GRC 報表,團名永遠是縮寫;這個工具一鍵自動補完整名稱。

[![Latest Release](https://img.shields.io/github/v/release/tony1229kimo/grc-name-fixer?label=latest&color=green)](https://github.com/tony1229kimo/grc-name-fixer/releases/latest)
[![License](https://img.shields.io/badge/license-internal-blue)]()

---

## 🎯 我是業務同事 — 怎麼用

> ⏱ **3 分鐘安裝、之後一輩子按一鍵搞定**

### 1. 下載

👉 **[點此前往最新版下載頁面](https://github.com/tony1229kimo/grc-name-fixer/releases/latest)**

在頁面下方 **Assets** 區塊,點 `grc-name-fixer-v1.x.x.zip` 下載(約 36 MB)。

### 2. 安裝

1. 解壓縮 zip,**整個資料夾**放到下面任一處:
   - ✅ **桌面**
   - ✅ `C:\GRC工具\`
   - ✅ D 槽根目錄
   - ❌ **不要放** OneDrive / Dropbox / Program Files / 公司網路硬碟
2. 雙擊資料夾裡的 **`GRC-團名修正工具.exe`**
3. 第一次 Windows 跳藍色警告 → 點「**其他資訊**」→「**仍要執行**」
4. 黑色視窗會自動跳出 + 瀏覽器自動開啟工具畫面 = 成功

### 3. 使用

1. 拖 Delphi 報表到「① 縮寫檔」
2. (選填) 拖含完整團名的扁平表到「② 參考檔」
3. 按「**🚀 開始自動替換團名**」
4. 看統計結果、按「**⬇️ 下載修正版檔案**」

完整教學見 **[`使用教學.docx`](使用教學.docx)**(下載 zip 裡也有一份)。

### 🔔 之後會自動提醒更新

工具頁面頂部會跳黃色 banner 通知新版可下載。點「立即前往下載」→ 重新下載 zip → 覆蓋舊資料夾 → 重啟即可。**你累積的字典 `dictionary.db` 不會被覆蓋**。

---

## 🛠️ 我是開發者 — 怎麼 build / 改 code

### 開發環境啟動

```powershell
# 1. 裝 Python deps
pip install -r requirements.txt

# 2. 啟動 dev server
python launcher.py
# 或直接 python app.py(沒有 port fallback / single-instance)
```

預設跑在 http://127.0.0.1:5000

### 打包 + 發 release

```powershell
# 改 _version.py 的 VERSION
# git commit + git push

# 一鍵打包 + zip + GitHub Release
.\build.ps1 -Release
```

腳本會做:
1. `pyinstaller grc-tool.spec --clean` → `dist\GRC-團名修正工具\`
2. 壓 `dist\grc-name-fixer-v{VERSION}.zip`
3. `gh release create v{VERSION}` 上傳 zip

⚠ 需要先 `gh auth login` 一次性登入。

### 檔案結構

```
grc-name-fixer/
├── launcher.py            ← 應用程式入口(port fallback / single-instance / banner)
├── app.py                 ← Flask 主程式(API endpoints)
├── parser.py              ← Excel 解析(openpyxl)
├── matcher.py             ← 匹配 & 替換邏輯
├── database.py            ← SQLite 字典 CRUD
├── seed.py                ← 預灌字典工具
├── _version.py            ← VERSION 常量 + GitHub repo 設定
├── grc-tool.spec          ← PyInstaller 設定
├── build.ps1              ← 一鍵 build + release 腳本
├── requirements.txt       ← Flask + openpyxl
├── dictionary.seed.db     ← 預灌字典種子(bundle 進 zip)
├── templates/             ← Jinja2 templates
├── static/                ← CSS/JS(含 version-check.js)
└── 使用教學.docx           ← 給同事的完整教學
```

### 預灌字典(累積歷史資料用)

手邊有歷史的 `0.xlsx` + 縮寫檔時,一次灌入種子:

```bash
python seed.py <0.xlsx> <縮寫檔.xlsx>
# 跑完會更新 dictionary.db,可手動 cp 成 dictionary.seed.db 進下一版
```

### 匹配邏輯

用「到達日期 + 總營收」當匹配 key:
1. 先查本地字典(最快)
2. 字典沒有 → 查上傳的參考檔
3. 匹配成功 → 學入字典(下次同縮寫免查)
4. 匹配失敗 → A 欄保留原文 + 加黃底紅字標示

容忍 ±2 日的到達日差異(因 Delphi 有時顯示實際 pickup 日而非 contract arrival)。

### Architecture 注意點

- **dictionary.seed.db vs dictionary.db** — seed 是 bundle 進 zip 的種子;launcher 第一次跑會 copy 成 dictionary.db。這樣升級新版本(覆蓋整個資料夾)不會沖掉同事累積的學習
- **output 檔案** — 存在硬碟用純 ASCII UUID 命名(`outputs/{uuid}.xlsx`),中文友善檔名透過 sidecar `{uuid}.name` 傳給瀏覽器(RFC 5987 編碼)。早期版本用中文檔名 + werkzeug `secure_filename()` 會把中文砍光導致 404
- **/api/version** — 4hr cache 打 GitHub Releases API 比對版號,前端 `static/version-check.js` 顯示通知 banner
- **Port fallback** — launcher 從 5000 試到 5019,找不到就吐錯誤
- **Single-instance** — 用 `%TEMP%\grc-name-fixer.lock` 記目前 port,雙開時偵測到就提示已執行
