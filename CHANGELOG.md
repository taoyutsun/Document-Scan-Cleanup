# 版本紀錄 / Changelog

[繁體中文說明](README.md) · [English README](README.en.md)

## 3.6.0 — 2026-11-25

- PowerPoint 批次處理新增「清除動畫效果」，可獨立勾選並用於合併輸出。保留投影片切換效果與媒體檔案。
- PowerPoint、Word 和 Excel 的個別批次副本統一使用 `.cleaned` 檔名後綴。

Adds optional PowerPoint animation removal for individual presentations and merged output. Slide transitions and media files are preserved. Individual PowerPoint, Word, and Excel batch copies use the `.cleaned` filename suffix.

## 3.5.3 — 2026-11-20

- 本次公開版納入 3.4.0 至 3.5.2 的 Word／Excel 批次清理及合併、可調整的檔案清單窗格和拖曳排序。
- 修正 Word 接受修訂時因自訂 XML 零件重新編號而被誤擋的問題。Word 合併後會核對兩份以上來源的自訂 XML 資料、屬性與關聯。

This public release includes the Word and Excel batch cleanup and merge features, resizable file lists and drag-to-reorder from 3.4.0 through 3.5.2. It fixes a false positive when Word renumbers custom XML parts while accepting revisions. Word merges verify the custom XML data, properties and relationships from all input files.

## 3.5.2（本機測試版 / Local test build）

- Office 批次清單支援直接拖曳排序，合併依更新後的清單順序處理。

Adds drag-to-reorder in the Office batch list. Merges use the resulting list order.

## 3.5.1（本機測試版 / Local test build）

- 關鍵字掃描的檔案清單預設加高。關鍵字掃描與 Office 批次處理都可拖曳清單下方的分隔線調整高度。

Increases the default height of the keyword scan file list and adds a draggable split between the file list and lower panel in both tabs.

## 3.5.0（本機測試版 / Local test build）

- Word 新增依序分節合併。Excel 新增依序合併工作表群組。同名工作表加上編號。
- 新增各格式與跨格式的清理全選／清除。各格式的合併及關鍵字操作仍可獨立選擇。

Adds ordered Word document and Excel worksheet merges. Adds per-format and all-Office controls for selecting or clearing cleanup options.

## 3.4.0（本機測試版 / Local test build）

- Office 批次處理支援在同一批次中加入 PowerPoint、Word 和 Excel，並將副本集中另存。
- Word 新增移除註解及接受全部修訂並停止追蹤。
- Excel 新增分別移除串列註解和附註，並可預覽隱藏工作表。

Adds mixed Office batches, Word comment removal and revision acceptance, and separate Excel threaded comment and note removal with a hidden sheet preview.

## 3.3.1 — 2026-11-20

- 自 3.2.1 起的公開更新。包含 3.3.0 試用版新增的 PowerPoint 批次處理、檔案拖曳及共用清單。
- 區分已移除的敏感度標籤紀錄、仍有效的標籤及 ActiveX。處理有相容標籤的簡報時保留原有分類，另存後驗證標籤狀態。

The public update from 3.2.1 includes the PowerPoint batch tools, shared file list and drag-and-drop support introduced in the 3.3.0 test build. It distinguishes removed label records, active sensitivity labels and ActiveX. The batch checks that active label metadata survives saving before delivering the output.

## 3.3.0 — 2026-11-20（本機測試版 / Local test build）

- 新增 PowerPoint 批次處理，可選擇移除備忘稿、隱藏投影片或註解，並依清單順序合併。
- 兩個分頁共用檔案清單及檔案拖曳功能，PowerPoint 批次也可套用已勾選的關鍵字項目。
- 更新視窗樣式，保留作者與原始碼連結，批次輸出另存並附處理報告。

Adds share-copy tools for PowerPoint notes, hidden slides, comments, and ordered merging. The two tabs share a drag-and-drop file list, and the PowerPoint batch can also apply checked keyword findings.

## 3.2.1 — 2026-10-01

- 新增「重新載入啟動清單」「開啟清單資料夾」與已儲存清單數量，方便確認私人設定位置與載入狀態。
- 改用可縮放視窗配置，保留底部作者連結；結果清單及紀錄區隨視窗縮放。
- 重新載入前保護未儲存的關鍵字修改，私人清單格式與公開預設維持不變。

Adds startup-profile reload, a profile-folder button and saved-list count. A responsive layout reserves the author footer while resizing results/logs. Reload asks before discarding unsaved keyword edits; profile format and public defaults are unchanged.

## 3.2.0 — 2026-10-01

- 新增本機關鍵字清單：載入、儲存、切換、設為 GUI 啟動預設及恢復通用預設。
- 設定保存在使用者資料夾，更新 EXE 後保留；公開原始碼與發布套件不包含個人清單。
- 新增 `--keywords-profile`；CLI 不自動套用 GUI 個人啟動設定。
- 清單與設定驗證、原子儲存及損壞時的通用預設提示。
- 公開示例與測試只使用通用或虛構字詞。

Local keyword profiles with import/save/switch/startup/reset actions; user settings survive executable updates. Explicit CLI profile selection, atomic writes and validation/fallback. Personal profiles are excluded from public source and release packages; examples use general or fictional terms only.

## 3.1.0 — 2026-10-01

首個 GitHub 公開版本，以交付文件前的指定內容去敏感化與範本整理為主題。

- 提供自訂關鍵字、逐項檢視、清理副本與變更報告。
- 支援 PPTX、DOCX、XLSX、PDF、EML 與常見文字格式；舊格式透過桌面 Office 轉檔。
- 加入通用預設範例與中文／英文 README，互相連結。
- 作者區塊連到 GitHub 專屬原始碼頁面。
- 公開版本僅含合成測試與一般驗證方法；不包含私人文件或工作紀錄。
- PPT 未使用範本整理不會額外改名使用中、未選取的主題。

First public GitHub release for user-directed content cleanup and template housekeeping before document delivery.

- Custom literal keywords, reviewed selections, separate outputs and change reports.
- Modern Office, PDF, EML and common text formats; desktop Office conversion for legacy formats.
- General defaults and interlinked Traditional Chinese / English READMEs.
- Author section links to the dedicated GitHub repository.
- Synthetic tests only; no private documents or work logs included.
- Unused PPT template pruning leaves active, unselected theme names unchanged.
