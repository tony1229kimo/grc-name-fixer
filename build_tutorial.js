// 產生 GRC 團名修正工具使用教學 Word 文件
const fs = require('fs');
const path = require('path');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, Footer, AlignmentType, PageOrientation, LevelFormat,
    TabStopType, TabStopPosition, BorderStyle, WidthType, ShadingType,
    VerticalAlign, PageNumber, HeadingLevel, PageBreak,
    TableOfContents, ExternalHyperlink, Bookmark, InternalHyperlink
} = require('docx');

// ============ 樣式常數 ============
const FONT = "Microsoft JhengHei";
const DARK = "1E293B";
const BLUE = "2563EB";
const GRAY = "64748B";
const LIGHT_BG = "F1F5F9";
const YELLOW = "FEF3C7";
const GREEN = "D1FAE5";
const RED_BG = "FEE2E2";
const RED = "C00000";

const border = (color = "CBD5E1") => ({
    style: BorderStyle.SINGLE, size: 4, color
});
const fullBorder = (color = "CBD5E1") => ({
    top: border(color), bottom: border(color),
    left: border(color), right: border(color),
});

// ============ Helper ============
function H1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 180 },
        children: [new TextRun({ text, bold: true, size: 36, color: BLUE, font: FONT })],
    });
}

function H2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 280, after: 140 },
        children: [new TextRun({ text, bold: true, size: 28, color: DARK, font: FONT })],
    });
}

function H3(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 100 },
        children: [new TextRun({ text, bold: true, size: 24, color: DARK, font: FONT })],
    });
}

function P(text, opts = {}) {
    return new Paragraph({
        spacing: { after: 100, line: 340 },
        children: [new TextRun({
            text, size: 22, font: FONT, color: DARK,
            bold: opts.bold, italics: opts.italics, color: opts.color || DARK,
        })],
    });
}

// 支援混合樣式（把陣列傳入 runs）
function Pmix(runs, opts = {}) {
    return new Paragraph({
        spacing: { after: opts.after || 100, line: 340 },
        alignment: opts.align,
        children: runs.map(r => new TextRun({
            text: r.text,
            size: r.size || 22,
            bold: r.bold,
            italics: r.italics,
            color: r.color || DARK,
            font: FONT,
            highlight: r.highlight,
        })),
    });
}

function BulletList(items, level = 0) {
    return items.map(item => new Paragraph({
        numbering: { reference: "bullets", level },
        spacing: { after: 80, line: 320 },
        children: typeof item === 'string'
            ? [new TextRun({ text: item, size: 22, font: FONT, color: DARK })]
            : item.map(r => new TextRun({
                text: r.text, size: 22, bold: r.bold,
                color: r.color || DARK, font: FONT, highlight: r.highlight,
            })),
    }));
}

function NumberedList(items) {
    return items.map(item => new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 80, line: 320 },
        children: typeof item === 'string'
            ? [new TextRun({ text: item, size: 22, font: FONT, color: DARK })]
            : item.map(r => new TextRun({
                text: r.text, size: r.size || 22, bold: r.bold,
                color: r.color || DARK, font: FONT, highlight: r.highlight,
            })),
    }));
}

// 彩色提示盒子（用單格表格模擬）
function Callout(titleText, bodyText, variant = 'info') {
    const bgColors = { info: "DBEAFE", warn: YELLOW, danger: RED_BG, success: GREEN };
    const borderColors = { info: BLUE, warn: "D97706", danger: RED, success: "047857" };
    const fg = borderColors[variant];
    const bg = bgColors[variant];

    return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({
            children: [new TableCell({
                width: { size: 9360, type: WidthType.DXA },
                shading: { fill: bg, type: ShadingType.CLEAR },
                borders: {
                    top: { style: BorderStyle.SINGLE, size: 4, color: fg },
                    bottom: { style: BorderStyle.SINGLE, size: 4, color: fg },
                    left: { style: BorderStyle.SINGLE, size: 16, color: fg },
                    right: { style: BorderStyle.SINGLE, size: 4, color: fg },
                },
                margins: { top: 160, bottom: 160, left: 200, right: 160 },
                children: [
                    new Paragraph({
                        spacing: { after: 80 },
                        children: [new TextRun({
                            text: titleText, bold: true, size: 22, color: fg, font: FONT,
                        })],
                    }),
                    ...(Array.isArray(bodyText) ? bodyText : [bodyText]).map(b =>
                        new Paragraph({
                            spacing: { after: 40, line: 320 },
                            children: [new TextRun({
                                text: b, size: 22, color: DARK, font: FONT,
                            })],
                        })
                    ),
                ],
            })],
        })],
    });
}

function Code(text) {
    return new Paragraph({
        spacing: { before: 80, after: 120 },
        shading: { fill: "F8FAFC", type: ShadingType.CLEAR },
        border: {
            top: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
            left: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
            right: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
        },
        children: [new TextRun({
            text, font: "Consolas", size: 20, color: DARK,
        })],
    });
}

function Spacer() {
    return new Paragraph({ spacing: { after: 120 }, children: [] });
}

function HorizontalRule() {
    return new Paragraph({
        spacing: { before: 160, after: 160 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
        children: [],
    });
}

// 簡易 3 欄表格
function simpleTable(headers, rows, widths) {
    const makeCell = (content, opts = {}) => new TableCell({
        width: { size: opts.width, type: WidthType.DXA },
        shading: opts.headerRow
            ? { fill: BLUE, type: ShadingType.CLEAR }
            : (opts.altBg ? { fill: LIGHT_BG, type: ShadingType.CLEAR } : undefined),
        borders: fullBorder("CBD5E1"),
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
            alignment: opts.align || AlignmentType.LEFT,
            children: [new TextRun({
                text: content, size: 21, font: FONT,
                bold: opts.headerRow,
                color: opts.headerRow ? "FFFFFF" : DARK,
            })],
        })],
    });

    return new Table({
        width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
        columnWidths: widths,
        rows: [
            new TableRow({
                tableHeader: true,
                children: headers.map((h, i) =>
                    makeCell(h, { width: widths[i], headerRow: true, align: AlignmentType.CENTER })
                ),
            }),
            ...rows.map((row, rowIdx) => new TableRow({
                children: row.map((c, i) =>
                    makeCell(c, { width: widths[i], altBg: rowIdx % 2 === 1 })
                ),
            })),
        ],
    });
}

// ============ 文件內容 ============

const sections = [];

// ----- 封面 -----
sections.push(
    new Paragraph({
        spacing: { before: 2400, after: 160 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({
            text: "GRC 團名修正工具", bold: true, size: 56, color: BLUE, font: FONT,
        })],
    }),
    new Paragraph({
        spacing: { after: 400 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({
            text: "使用教學手冊", size: 32, color: GRAY, font: FONT,
        })],
    }),
    new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 800, after: 100 },
        children: [new TextRun({
            text: "台中勤美洲際酒店", size: 28, bold: true, color: DARK, font: FONT,
        })],
    }),
    new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 2000 },
        children: [new TextRun({
            text: "業務部 GRC 報表處理專用", size: 24, color: GRAY, font: FONT,
        })],
    }),
    new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "版本 1.0", size: 20, color: GRAY, font: FONT })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
);

// ----- 目錄 -----
sections.push(
    H1("目錄"),
    Spacer(),
    Pmix([
        { text: "一、工具簡介 ......................................................", color: GRAY },
        { text: "1", bold: true },
    ]),
    Pmix([
        { text: "二、第一次使用（安裝） ............................................", color: GRAY },
        { text: "2", bold: true },
    ]),
    Pmix([
        { text: "三、日常使用流程 ..................................................", color: GRAY },
        { text: "4", bold: true },
    ]),
    Pmix([
        { text: "四、看懂處理結果 ..................................................", color: GRAY },
        { text: "6", bold: true },
    ]),
    Pmix([
        { text: "五、字典管理 ......................................................", color: GRAY },
        { text: "8", bold: true },
    ]),
    Pmix([
        { text: "六、常見狀況說明 ..................................................", color: GRAY },
        { text: "10", bold: true },
    ]),
    Pmix([
        { text: "七、疑難排解 Q&A ...................................................", color: GRAY },
        { text: "11", bold: true },
    ]),
    new Paragraph({ children: [new PageBreak()] }),
);

// =============================================================
// 一、工具簡介
// =============================================================
sections.push(
    H1("一、工具簡介"),
    H2("這個工具是做什麼的？"),
    P("Delphi 系統跑出來的 GRC 報表，團名欄（A 欄）常常是「縮寫」或被「截斷」的版本，例如："),
);

// 對照表
sections.push(
    Spacer(),
    simpleTable(
        ["Delphi 原始文字（縮寫）", "應該顯示的完整名稱"],
        [
            ["Yung Shin Phar...", "Yung Shin Pharmaceutical Industrial Meeting Group"],
            ["Inter Asia_HK ... ^", "Inter Asia_HK Inbound Group"],
            ["TUA_Bon Voyage TA ^", "Taiwan Urological Association_Bon Voyage TA"],
            ["Grundfos Pumps... ^", "Grundfos Pumps Taiwan Ltd_Meeting"],
        ],
        [3600, 5760],
    ),
    Spacer(),
    P("以前要手動一個一個從其他檔案複製貼上，花很多時間。這個工具會："),
    ...BulletList([
        [
            { text: "自動匹配 ", bold: false },
            { text: "縮寫 ↔ 完整名稱", bold: true, color: BLUE },
            { text: "，一鍵替換 A 欄團名" },
        ],
        [
            { text: "只改 A 欄文字，", bold: false },
            { text: "其他數字、日期、公式、格式完全不動", bold: true, color: "047857" },
        ],
        [
            { text: "會「" },
            { text: "學習", bold: true, color: BLUE },
            { text: "」看過的縮寫，越用越聰明，下次免上傳參考檔" },
        ],
        [
            { text: "找不到的團會用 " },
            { text: "黃底紅字標示", bold: true, color: RED, highlight: "yellow" },
            { text: "，方便之後手動查補" },
        ],
    ]),
    Spacer(),
    Callout("💡 一句話總結",
        "把 Delphi 原始檔拖進來 → 按一個按鈕 → 團名就全部變正確了。",
        'info'),
    new Paragraph({ children: [new PageBreak()] }),
);

// =============================================================
// 二、第一次使用
// =============================================================
sections.push(
    H1("二、第一次使用（安裝）"),
    H2("收到的東西長這樣"),
    P("主管會給你一個資料夾，裡面類似這樣："),
    Code("GRC-團名修正工具/\n  ├── GRC-團名修正工具.exe       ← 雙擊這個啟動！\n  ├── _internal/                  ← 程式需要的東西（不要刪！）\n  └── dictionary.db               ← 字典資料庫（不要刪！）"),
    Spacer(),
    Callout("⚠️ 注意",
        [
            "請把整個資料夾放到一個固定位置（例如桌面或 D 槽），不要只複製 exe 出來。",
            "資料夾裡的 _internal 和 dictionary.db 都不要刪掉或移動。",
        ],
        'warn'),
    Spacer(),

    H2("第一次啟動的 3 個步驟"),
    ...NumberedList([
        [
            { text: "打開資料夾，找到 " },
            { text: "「GRC-團名修正工具.exe」", bold: true, color: BLUE },
            { text: "，雙擊開啟。" },
        ],
        [
            { text: "Windows 可能會跳出" },
            { text: "「Windows 已保護您的電腦」", bold: true },
            { text: "藍色警告畫面，這是因為程式沒有花錢買簽章。請點左側的" },
            { text: "「其他資訊」", bold: true, color: BLUE },
            { text: "，再按下方的" },
            { text: "「仍要執行」", bold: true, color: BLUE },
            { text: "。" },
        ],
        [
            { text: "會自動跳出一個黑色視窗（不要關），然後瀏覽器會自動開啟" },
            { text: "「http://127.0.0.1:5000」", bold: true, color: BLUE },
            { text: "，看到工具主畫面就成功了。" },
        ],
    ]),
    Spacer(),
    Callout("✅ 啟動成功的樣子",
        [
            "• 有一個黑色視窗寫著「GRC 團名修正工具」、「伺服器位置: http://127.0.0.1:5000」",
            "• 瀏覽器自動打開並顯示工具畫面",
            "• 黑色視窗不能關（關了工具就停了）",
        ],
        'success'),
    Spacer(),

    H2("要關閉工具"),
    P("用完之後，直接把那個黑色視窗關掉就可以了（按右上角 X 或 Ctrl+C）。"),
    Callout("💡 下次要用怎麼辦？",
        "再雙擊一次 GRC-團名修正工具.exe 就好，不用重新安裝。",
        'info'),
    new Paragraph({ children: [new PageBreak()] }),
);

// =============================================================
// 三、日常使用流程
// =============================================================
sections.push(
    H1("三、日常使用流程"),
    P("整個流程大約 10 秒就能跑完一份報表。共 3 步："),
    Spacer(),
    simpleTable(
        ["步驟", "動作", "備註"],
        [
            ["1", "上傳縮寫檔（必填）", "Delphi 直接跑出來的原始檔"],
            ["2", "上傳參考檔（選填）", "字典有學過的團就免上傳"],
            ["3", "按「開始自動替換團名」", "等幾秒跑完"],
            ["4", "下載修正版", "打開檢查、存檔"],
        ],
        [1000, 4000, 4360],
    ),
    Spacer(),

    H2("步驟 1：上傳縮寫檔（必填）"),
    P("這是 Delphi 系統剛跑出來的原始 GRC 檔，通常檔名類似："),
    Code("GRC Report_MMDDYYYY_R1.xlsx\n或\nGrc (xx).xlsx"),
    Spacer(),
    P("操作方式："),
    ...BulletList([
        "把檔案「拖曳」到頁面上 ① 縮寫檔 的黃色框框裡",
        "或者點一下框框，選檔案",
        "上傳成功會變綠色並顯示檔名",
    ]),
    Spacer(),
    Callout("🎯 判斷是不是縮寫檔的方法",
        "打開 Excel 看一下 A 欄，如果團名有「...」或結尾有「^」符號，就是縮寫檔。",
        'info'),
    Spacer(),

    H2("步驟 2：上傳參考檔（選填）"),
    P("這是「扁平格式」的完整名單，Delphi 用另一個報表跑出來的，A 欄是完整團名。"),
    Spacer(),
    Callout("⚡ 什麼時候要填參考檔？",
        [
            "第一次用這個工具 → 建議填，讓字典學",
            "這週有新的團從來沒出現過 → 要填",
            "只要是字典已經認識的團 → 不用填",
        ],
        'info'),
    Spacer(),

    H2("步驟 3：按按鈕 → 等結果"),
    ...NumberedList([
        "確認檔案都上傳好了",
        [
            { text: "按下藍色大按鈕 " },
            { text: "「🚀 開始自動替換團名」", bold: true, color: BLUE },
        ],
        "畫面會出現「處理中...」的轉圈圈",
        [
            { text: "幾秒後會跳出" },
            { text: "「處理結果」", bold: true, color: BLUE },
            { text: "區塊，顯示各項統計" },
        ],
    ]),
    new Paragraph({ children: [new PageBreak()] }),
);

// =============================================================
// 四、看懂處理結果
// =============================================================
sections.push(
    H1("四、看懂處理結果"),
    P("處理完會看到 5 個統計數字，每個代表不同狀況："),
    Spacer(),
    simpleTable(
        ["欄位", "代表意思", "要做什麼"],
        [
            ["總筆數", "縮寫檔裡共有幾團", "—"],
            ["字典命中", "字典裡有記錄，直接替換", "✓ 不用動"],
            ["參考檔匹配", "從你上傳的參考檔查到的", "✓ 不用動"],
            ["新學到的縮寫", "這次學到、以後免上傳的", "✓ 不用動"],
            ["未匹配", "找不到完整名的（需手動）", "⚠️ 要手動補"],
        ],
        [1800, 3560, 4000],
    ),
    Spacer(),

    H2("關於「未匹配」的團"),
    P("如果統計裡有顯示未匹配數字，表示有幾個團找不到完整名稱。工具會："),
    ...BulletList([
        "在網頁下方列出這些團的名單（包含日期、營收）",
        [
            { text: "在下載的 Excel 裡把這幾列 " },
            { text: "A 欄用黃底紅字標記", bold: true, color: RED, highlight: "yellow" },
            { text: "，讓你一眼看到" },
        ],
        "原本的縮寫文字保留不動（方便你還知道原來的縮寫是什麼）",
    ]),
    Spacer(),
    Callout("📝 未匹配的處理建議",
        [
            "1. 打開下載的 Excel，找到黃底紅字的列",
            "2. 從其他地方（上週的 GRC、合約、Delphi 系統）查到正確的完整名稱",
            "3. 直接在 Excel 裡把 A 欄改成完整名稱，把黃底去掉",
            "4. 如果以後這個縮寫還會再出現，建議到「字典管理」頁面手動新增",
        ],
        'warn'),
    Spacer(),

    H2("下載修正版"),
    ...NumberedList([
        [
            { text: "按 " },
            { text: "「⬇️ 下載修正版檔案」", bold: true, color: BLUE },
            { text: " 按鈕" },
        ],
        "檔案會存到你瀏覽器預設的下載資料夾",
        [
            { text: "檔名會自動加上日期時間，例如：" },
            { text: "Grc (7)_已修正_20260421_143022.xlsx", bold: true, color: BLUE },
        ],
        "打開 Excel 檢查，沒問題就可以存檔使用",
    ]),
    new Paragraph({ children: [new PageBreak()] }),
);

// =============================================================
// 五、字典管理
// =============================================================
sections.push(
    H1("五、字典管理"),
    H2("什麼是字典？"),
    P("工具會記住看過的「縮寫 → 完整名稱」對應，存在一個叫做「字典」的資料庫。"),
    P("下次同一個縮寫再出現，工具不用問 0.xlsx 就能直接替換。這就是工具會「越用越聰明」的原因。"),
    Spacer(),

    H2("打開字典管理頁"),
    P("在網頁上方的導覽列，點「字典管理」連結。"),
    Spacer(),

    H2("你可以做的事"),
    Spacer(),
    simpleTable(
        ["功能", "用途", "怎麼做"],
        [
            ["搜尋", "找特定的縮寫或名稱", "在右上角搜尋框打關鍵字"],
            ["新增", "手動補一個新對應", "按「➕ 新增對應」按鈕"],
            ["刪除", "移除錯誤的對應", "每列最右邊有「刪除」按鈕"],
        ],
        [1600, 3800, 3960],
    ),
    Spacer(),

    H2("什麼時候要手動新增？"),
    ...BulletList([
        [
            { text: "工具標為" },
            { text: "「未匹配」", bold: true, color: RED },
            { text: "的縮寫 → 你手動查到完整名稱後，來這裡補進去，下次就認識了" },
        ],
        "Delphi 系統改版後出現新的縮寫格式",
        "某個團的縮寫記錯了，先刪掉再新增正確的",
    ]),
    Spacer(),

    H2("手動新增教學"),
    ...NumberedList([
        [
            { text: "點 " },
            { text: "「➕ 新增對應」", bold: true, color: BLUE },
            { text: " 按鈕" },
        ],
        [
            { text: "「縮寫」欄：" },
            { text: "完整複製 Delphi A 欄文字", bold: true },
            { text: "（包含 ... 或 ^ 符號）" },
        ],
        "「完整名稱」欄：填你查到的正式團名",
        "「備註」欄：想記下的來源或說明（可不填）",
        [
            { text: "按 " },
            { text: "「儲存」", bold: true, color: BLUE },
            { text: "，搞定" },
        ],
    ]),
    Spacer(),
    Callout("💡 小技巧",
        [
            "複製縮寫時要「完整複製」，包含空格和特殊符號（... 和 ^）",
            "Delphi 產出的截斷方式是固定的，只要複製一次就會一直對",
        ],
        'info'),
    new Paragraph({ children: [new PageBreak()] }),
);

// =============================================================
// 六、常見狀況說明
// =============================================================
sections.push(
    H1("六、常見狀況說明"),

    H2("狀況 1：字典命中很多、參考檔匹配很少"),
    P("這是最理想的狀況！表示字典已經累積夠多對應，大部分團不用上傳參考檔就能認得。"),
    Callout("🎯 建議做法",
        "繼續這樣用，下次甚至可以不用上傳參考檔。只有遇到新團才需要。",
        'success'),
    Spacer(),

    H2("狀況 2：字典命中 0、全部靠參考檔"),
    P("這通常是第一次用這個工具、或 dictionary.db 被刪掉重建。參考檔匹配後會自動學入字典。"),
    Callout("🎯 建議做法",
        "正常現象。每週用完字典會越長越大，之後就會漸漸變成「字典命中為主」。",
        'info'),
    Spacer(),

    H2("狀況 3：未匹配的團特別多"),
    P("可能原因："),
    ...BulletList([
        "沒上傳參考檔，而這週有新團（字典還不認識）",
        "參考檔和縮寫檔的日期對不上（例如一個是 4 月 Plan、一個是 5 月更新版）",
        "這些團真的不在 0.xlsx 裡（例如內部活動、待處理團）",
    ]),
    Callout("🎯 建議做法",
        [
            "1. 先試著上傳這週的 0.xlsx 當參考檔，重跑一次",
            "2. 還是未匹配的，表示 0.xlsx 裡也沒有，需要手動補",
            "3. 手動補完後，別忘了到「字典管理」新增對應，下次就認識了",
        ],
        'warn'),
    Spacer(),

    H2("狀況 4：工具替換錯了團名"),
    P("極少發生，但可能的原因：有兩團到達日、營收都剛好一模一樣，系統配錯了。"),
    Callout("🎯 建議做法",
        [
            "1. 到「字典管理」搜尋那個錯誤的縮寫，刪掉",
            "2. 手動新增正確的對應",
            "3. 重新上傳檔案跑一次",
        ],
        'warn'),
    new Paragraph({ children: [new PageBreak()] }),
);

// =============================================================
// 七、疑難排解 Q&A
// =============================================================
sections.push(
    H1("七、疑難排解 Q&A"),
    Spacer(),

    H3("Q1：雙擊 exe 沒反應？"),
    P("A：可能是："),
    ...BulletList([
        "防毒軟體誤判擋住了 → 到防毒設定把這個 exe 加到白名單",
        "Windows SmartScreen 擋住 → 點「其他資訊」→「仍要執行」",
        "路徑有特殊字元 → 把資料夾移到 C:\\ 或 D:\\ 根目錄試試",
    ]),
    Spacer(),

    H3("Q2：黑色視窗一閃就消失？"),
    P("A：程式遇到錯誤但還沒來得及顯示就退出了。試試："),
    ...NumberedList([
        "按 Windows 鍵 + R，輸入 cmd，按 Enter",
        "在黑色視窗輸入：cd「你工具資料夾的完整路徑」",
        "再輸入：GRC-團名修正工具.exe",
        "這樣錯誤訊息就不會消失了，把訊息拍給主管看",
    ]),
    Spacer(),

    H3("Q3：瀏覽器沒自動打開？"),
    P("A：手動打開 Chrome / Edge，網址列貼上："),
    Code("http://127.0.0.1:5000"),
    Spacer(),

    H3("Q4：說 port 5000 被佔用？"),
    P("A：有其他程式在用 5000 port。關掉其他程式（特別是其他開發工具、Skype 舊版）再試。如果還是不行，請主管改 port 號碼。"),
    Spacer(),

    H3("Q5：下載的檔案打不開 / 格式怪怪？"),
    P("A：99% 是沒下載完整。檢查檔案大小是否合理（應該 50KB 以上）。如果確實壞掉，關掉工具重開再跑一次。"),
    Spacer(),

    H3("Q6：想把字典備份到其他電腦？"),
    P("A：整個字典就是一個檔案 "),
    Code("dictionary.db"),
    P("複製這個檔案到另一台電腦的工具資料夾（覆蓋原本的那個），就完成搬家。"),
    Spacer(),

    H3("Q7：字典會不會爆炸（檔案太大）？"),
    P("A：不會。每筆資料只有幾十個 bytes，就算記了 10000 個對應，整個 db 也才幾 MB。完全不用擔心。"),
    Spacer(),

    H3("Q8：同一個縮寫的完整名稱改變了怎辦？"),
    P("A：到「字典管理」搜尋該縮寫 → 刪掉 → 下次跑的時候若有新的參考檔，會重新學到正確的對應。"),
    Spacer(),

    H3("Q9：工具支援 Mac 嗎？"),
    P("A：目前版本只支援 Windows。如需要 Mac 版請找主管。"),
    Spacer(),

    H3("Q10：可以同時處理多份檔案嗎？"),
    P("A：目前一次只能處理一份。處理完下載後，按「處理下一份」就可以跑下一份。"),
    Spacer(),

    HorizontalRule(),
    Spacer(),
    Pmix([
        { text: "其他問題或建議，請聯絡：", color: GRAY },
        { text: "業務部系統管理員", bold: true, color: BLUE },
    ]),
    Spacer(),
    Pmix([
        { text: "© 2026 台中勤美洲際酒店 · 版本 1.0", color: GRAY, size: 20 },
    ], { align: AlignmentType.CENTER }),
);

// ============ 組裝文件 ============

const doc = new Document({
    creator: "台中勤美洲際酒店",
    title: "GRC 團名修正工具使用教學",
    description: "GRC 報表團名自動修正工具使用手冊",
    styles: {
        default: {
            document: { run: { font: FONT, size: 22 } },
        },
        paragraphStyles: [
            {
                id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
                quickFormat: true,
                run: { size: 36, bold: true, color: BLUE, font: FONT },
                paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 },
            },
            {
                id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
                quickFormat: true,
                run: { size: 28, bold: true, color: DARK, font: FONT },
                paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 },
            },
            {
                id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal",
                quickFormat: true,
                run: { size: 24, bold: true, color: DARK, font: FONT },
                paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 },
            },
        ],
    },
    numbering: {
        config: [
            {
                reference: "bullets",
                levels: [
                    { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                      style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
                ],
            },
            {
                reference: "numbers",
                levels: [
                    { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                      style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
                ],
            },
        ],
    },
    sections: [
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },  // A4
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
                },
            },
            headers: {
                default: new Header({
                    children: [new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [new TextRun({
                            text: "GRC 團名修正工具使用教學",
                            size: 18, color: GRAY, font: FONT,
                        })],
                    })],
                }),
            },
            footers: {
                default: new Footer({
                    children: [new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun({ text: "— 第 ", size: 18, color: GRAY, font: FONT }),
                            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GRAY, font: FONT }),
                            new TextRun({ text: " 頁 —", size: 18, color: GRAY, font: FONT }),
                        ],
                    })],
                }),
            },
            children: sections,
        },
    ],
});

const outputPath = path.join(__dirname, "使用教學.docx");
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(outputPath, buffer);
    console.log("✓ 文件已產出：" + outputPath);
    console.log("  大小：" + (buffer.length / 1024).toFixed(1) + " KB");
}).catch(err => {
    console.error("✗ 產出失敗：", err);
    process.exit(1);
});
