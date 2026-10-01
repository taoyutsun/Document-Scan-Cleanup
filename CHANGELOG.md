# 版本紀錄 / Changelog

[繁體中文說明](README.md) · [English README](README.en.md)

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
