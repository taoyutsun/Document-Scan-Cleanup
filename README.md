# 文件去敏感化與範本清理工具

**Document Scan & Cleanup** · Windows 本機工具 · MIT License

**繁體中文** | [English](README.en.md)

交付文件給客戶、供應商或合作夥伴前，依自己指定的關鍵字檢查內容、選擇要移除的項目，再另存整理好的副本。工具支援 PowerPoint、Word、Excel、PDF、郵件與常見文字格式，盡量保留原有格式與版面。

## 為什麼開發這個工具？

許多個人、公司或單位會共用一套文件範本，未必另備內部與對外使用的版本。文件製作完成後，若要提供給客戶、供應商或其他合作夥伴，可能需要移除內部註記、草稿標示、特定人名、聯絡資訊、專案代號，或其他不適合出現在交付副本中的內容。

手動處理常常比想像中花時間。文字可能藏在頁首頁尾、簡報母片、未使用的版面配置，甚至文件內部名稱與屬性。把 PowerPoint 內容直接貼進其他範本，又可能遇到字型、圖片或文字框跑版，需要逐頁調整。

這個工具把「掃描 → 檢視並勾選 → 清理另存 → 檢查副本」放在同一個流程，讓使用者在既有文件上整理指定內容，減少反覆搬移與手動尋找的工作。

**去敏感化範圍由使用者決定。** 工具比對指定文字，並非自動辨識所有個資或機密資訊的系統，也不會重新判定文件的公開等級。

## 下載與快速開始

1. 到 [Releases](https://github.com/taoyutsun/Document-Scan-Cleanup/releases/latest) 下載 Windows x64 可攜式 ZIP。
2. 完整解壓縮，執行 `DocumentScanCleanup.exe`。處理新版文件不需安裝 Python 或 Office。
3. 按「選擇檔案」，加入需要檢查的文件。
4. 輸入關鍵字，以分號分隔，例如：`Draft; Internal Use Only; 專案代號; client@example.invalid`。範例是示意內容，請換成自己要處理的文字。
5. 按「掃描」，檢視結果；雙擊可查看詳細內容，點第一欄或按空白鍵調整勾選。
6. 按「清理選取項目並另存」，選擇結果位置。**原檔保留，不覆寫。**
7. 開啟副本，檢查內容、換行及版面，再進行分享。

GUI 使用繁體中文。EXE 尚未做程式碼簽署，Windows 可能顯示 SmartScreen 提示；是否允許執行仍依裝置的軟體管理設定。

## 個人關鍵字清單（3.2.0 起）

工具內建通用範例，也可保存自己的工作清單。**個人清單放在使用者設定資料夾，不寫入 EXE、GitHub 原始碼或發布 ZIP。** 更換／更新 EXE 不會刪除清單。

1. 在關鍵字欄輸入自己的詞，以分號分隔。
2. 按「儲存清單」並命名。儲存會更新目前載入的本機清單；通用預設會另建個人清單。
3. 按「設為啟動預設」，下次啟動 GUI 即可自動帶入。若有未儲存的修改，會先要求命名並儲存。
4. 下拉選單可切換已保存的清單。「載入清單」可匯入 UTF-8 JSON，驗證後複製至本機設定，不修改匯入的原檔。
5. 「恢復通用預設」會恢復目前與下次啟動的通用字詞；個人清單仍保留。

Windows 儲存位置：`%APPDATA%\DocumentScanCleanup\profiles\`；啟動選擇存在同一上層資料夾的 `keyword-settings.json`。每份清單使用獨立識別碼，因此名稱相同也不會彼此覆寫。修改文字不會自動保存；清單或啟動設定損壞時，會顯示提示並使用通用預設，不改寫原設定。

「開啟清單資料夾」可直接找到實際 JSON 檔案；介面顯示已儲存清單數量。外部新增清單或更改啟動設定後，按「重新載入啟動清單」即可重新讀取；若目前有未儲存修改，會先詢問是否放棄。清單放在 `profiles` 子資料夾，`keyword-settings.json` 本身只記錄啟動選擇。

視窗使用可縮放配置，結果清單與紀錄區會隨視窗大小調整，並保留底部作者連結；不需最大化才看得到。

匯入格式範例（以下均為虛構字詞）：

```json
{
  "schema_version": 1,
  "name": "範例工作清單",
  "keywords": ["Draft", "Internal Use Only", "Project Cedar", "範例詞"]
}
```

支援 1～100 個非空白關鍵字，每個最多 150 字；單一字詞不可含分號或換行。個人清單與含關鍵字／命中文字的報告皆應留在私人工作區，勿加入公開文件、截圖或問題回報。

**CLI 使用獨立且明確的選擇**：未指定關鍵字時使用通用預設，不自動套用 GUI 的個人啟動設定。可用 `--keywords-profile .\my-list.json` 指定清單，與 `--keywords` 二擇一；搭配 `--gui` 則只套用於該次介面啟動，不更改保存的啟動預設。

## 可以處理哪些格式？

| 格式 | 掃描 | 清理範圍與條件 |
| --- | --- | --- |
| PPTX | 投影片、母片、版面、備忘稿、名稱與屬性 | 選定文字／名稱；可另選移除未使用的母片、版面與主題，保留使用中的背景及格式依賴 |
| DOCX | 正文、頁首頁尾、表格、文字方塊、註解、修訂、樣式與屬性 | 選定文字及支援的名稱；同步更新樣式引用，保留原格式結構 |
| XLSX | 一般／隱藏儲存格、頁首頁尾、傳統註解、部分圖形與屬性 | 選定文字；保留公式、數值及未選取儲存格，共用文字的其他引用保留 |
| PDF | 可解析文字區塊、一般文件屬性，列出頁碼 | **刪除勾選的整個文字顯示區塊**；保留其他文字位置、圖片與頁面尺寸 |
| EML | 標題、純文字／HTML 正文與支援格式附件 | 選定標題／正文；附件與內嵌圖片保留，附件中的命中仍會列出 |
| PPT／PPS | 先轉成 PPTX | 需要桌面版 PowerPoint；另存新版格式，不改寫舊二進位檔 |
| DOC／RTF | 先轉成 DOCX | 需要桌面版 Word；另存新版格式 |
| XLS | 先轉成 XLSX | 需要桌面版 Excel；另存新版格式 |
| TXT／MD／LOG | 常見文字編碼與逐行文字 | 關鍵字或選定文字行；保留編碼、BOM 與換行 |
| CSV／TSV | 文字內容 | 僅移除關鍵字，保留引號、分隔符號與資料列 |
| JSON／XML | 文字與屬性中的關鍵字 | **只掃描**，避免破壞結構與欄位名稱 |

目前不支援 MSG、巨集格式、OpenDocument、圖片 OCR，以及內嵌 Office／OLE 文件的自動清理。

## 比對與清理方式

- **自訂文字比對**：不分大小寫、使用字面文字，支援中文、英文與其他字元；不使用正規表示式。可命中較長字串內的片段，請看前後文。
- **一般預設範例**：包含 Confidential、Internal Use Only、Draft、機密、草稿等標示，可直接改成要整理的人名、完整聯絡文字、專案代號或其他指定內容。
- **關鍵字模式**：只移除選定項目中的匹配文字。PPT／Word 可識別跨格式分段的字詞；郵件 HTML 也可處理跨標籤的正文文字。
- **段落／文字行模式**：清空選定的支援文字區塊，保留框架、段落與格式，不刪除整個文字框或表格。
- **PDF 固定使用區塊刪除**：可能連同同一句的其他文字一起移除；區塊也可能只是長段落的一部分。請雙擊看完整範圍，工具不會自動補寫句子或重排文字。
- **PDF／Excel／郵件預設不勾選清理項目**，由使用者先檢視再選擇。CLI 的 `--clean` 則預設處理全部可清理命中項目。
- **PPT 未使用範本整理與關鍵字選取分開**：移除未使用結構，不會因為啟用此選項就另行改名使用中的、未選取的主題。

Excel 頁首頁尾、郵件標題及名稱／屬性只移除指定字樣；CSV／TSV 不提供整行清空。移除文字可能改變 Word 換行或分頁，請檢查副本。

## 結果與原檔保留

GUI 會建立新的 `Document_Cleanup_日期時間` 資料夾。每份文件都有自己的編號子資料夾，內含清理副本、TXT／JSON 變更報告與剩餘命中；批次另有 `batch.report.json`。只需檢查時，可選「儲存掃描報告」。

來源文件在掃描後被編輯時，工具會要求重新掃描；輸出已存在時停止，不覆寫已有結果。不同儲存位置或引用可能重複命中，數量不等於不同敏感內容的數量。

**報告會包含文件路徑及文字片段。** 請與來源文件同樣妥善保存；回報問題時使用合成範例，不要附上真實客戶文件或未整理的報告。

## 本機處理與功能界線

- 文件在本機處理，不主動上傳文件，不需要 API key，不含遙測或更新檢查。作者／原始碼連結只有在使用者點擊時開啟預設瀏覽器。
- 正式敏感度標籤、未知 `customXml`、Office 欄位指令及未支援的識別碼保留；工具不會移除存取控制，也不會把文件重新分類為公開資料。
- 簽章、加密、啟用保護或不支援的結構不清理。EML 含 DKIM、ARC、S/MIME、PGP 或封裝郵件附件時只掃描。
- 不做 OCR。圖片、PDF 特殊巢狀內容／表單／附件、Office 內嵌文件等未完整解析；介面區分檢查完成、部分未檢查與無法掃描。**零命中不代表已移除所有敏感資訊。**
- 透過 Office 轉檔時會停用巨集／事件與連結更新，只使用可確認的獨立程序。Office 自身設定仍由裝置環境管理；轉檔可能影響字型或相容性物件。PowerPoint 已開啟時，請先手動另存 PPTX 或關閉後再處理。
- XLSX 會保留公式及快取值，不負責重新計算公式結果；必要時請在 Excel 開啟副本重新計算。
- 一般檔案／Office 解壓後上限 600 MB，ZIP 零件上限 25,000；舊格式自動轉檔上限 100 MB、90 秒；郵件附件遞迴深度 2、單附件 100 MB、總解碼量 200 MB。

## 原始碼執行與建置

建議使用 **Windows x64／Python 3.10**（本版驗證環境）。

```powershell
git clone https://github.com/taoyutsun/Document-Scan-Cleanup.git
cd Document-Scan-Cleanup
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -X utf8 document_scan_cleanup.py --gui
```

CLI 範例（請改成自己的文件位置）：

```powershell
python -X utf8 document_scan_cleanup.py --scan .\example.docx .\example.pdf --keywords 'Draft;專案代號' --output-dir .\scan-results
python -X utf8 document_scan_cleanup.py --clean .\example.pptx --keywords 'Internal Use Only;僅限內部使用' --prune --output-dir .\clean-results
```

CLI 支援 `--mode keyword|block`、`--prune` 與 `--selection-json`（完整來源路徑對應命中 ID 陣列）。GUI 版 EXE 不顯示終端文字，但可產生指定的檔案與報告；失敗結果寫入 `errors.report.json`。

測試與建置：

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -p 'test_*.py' -v
.\scripts\build.ps1 -Python .\.venv\Scripts\python.exe
.\scripts\package.ps1 -Python .\.venv\Scripts\python.exe
```

建置產物在 `dist/`，不寫入原始碼版本庫。測試使用程式產生的合成文件，沒有任何私人或實際組織文件。[測試說明](docs/TESTING.md) · [版本紀錄](CHANGELOG.md) · [第三方授權](THIRD_PARTY_NOTICES.md)

## 作者與授權

本工具由 **Arthur Tao** 設計與維護，採 [MIT License](LICENSE) 授權。歡迎使用與分享，並保留原作者與來源資訊。

- [亞瑟 ASK 部落格](https://taoyutsun.blogspot.com/)
- [Facebook](https://facebook.com/arthurtaoyutsun)
- [GitHub 原始碼](https://github.com/taoyutsun/Document-Scan-Cleanup)

第三方元件遵循各自的授權，不因本工具採 MIT 而改變。請見 `third-party-licenses/` 與 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
