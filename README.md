# 設定驅動 xlsx 報表產生器（Prototype）

以 ASP.NET Core MVC (.NET 8) 建立的單一專案 prototype，驗證「設定驅動的報表產生」流程是否可行：
上傳範本 → 掃描標記 → 設定資料來源與 JSONPath 映射 → 依參數打 API 取資料 → 套版產生 xlsx。

不含登入、不含正式資料庫（設定檔存在磁碟 JSON，範本存在磁碟 xlsx），UI 為陽春的 Razor + Bootstrap（CDN）。

## 技術棧

- ASP.NET Core 8 MVC（Razor Views，無 SPA 框架，Bootstrap 5 走 CDN）
- MiniExcel 1.34.2（`SaveAsByTemplate` 套版；版本鎖定以對齊正式專案的 .NET Core 3.1 相容性）
- Newtonsoft.Json（`SelectToken` / `SelectTokens` 做 JSONPath 映射）
- DocumentFormat.OpenXml（產生示範範本、控制框線/合併儲存格/列印版面設定）
- 設定檔：`App_Data/configs/{code}.json`（無資料庫）
- 範本檔：`App_Data/templates/*.xlsx`
- 產生紀錄：`App_Data/logs/report-{yyyyMMdd}.log`

## 啟動方式

```bash
dotnet restore
dotnet run
```

啟動時會自動：

1. 若 `App_Data/templates/` 是空的，用 OpenXML 產生一份示範範本 `sample_inspection.xlsx`（外部檢驗紀錄表版面，含表頭欄位標記、明細標記列、框線、合併儲存格與列印設定）。
2. 若尚無任何設定檔，seed 一份對應示範範本的完整設定 `IPQC_EXT_INSPECT`（外部檢驗紀錄表），資料來源指向本機的 Mock API。

瀏覽器開啟 `http://localhost:5047`（實際埠號以主控台顯示的 `Now listening on` 為準）即可開始操作。

## 三個頁面的操作順序

### 1. 先看/測試 Mock 資料來源（非必要，但方便理解資料形狀）

`GET /api/mock/inspection/{docNo}`，例如 `http://localhost:5047/api/mock/inspection/0040`。
明細筆數依 `docNo` 內容做穩定雜湊變化（3~15 筆不等），同一個 docNo 每次呼叫結果一致，方便重複測試版面在不同明細筆數下是否正確。

### 2. `/Report/Index`：產生報表（驗收的主要入口）

1. 左側清單選擇「外部檢驗紀錄表」
2. 右側輸入參數「單據號」，例如 `0040`
3. 按「產生並下載 xlsx」，會在新分頁下載檔案（`{報表名稱}_{docNo}_{yyyyMMddHHmmss}.xlsx`）
4. 換一個 docNo（例如 `0099`）重新產生，可觀察明細筆數變多時版面是否仍正確（框線、合併儲存格、重複標題列不跑掉）

也可以直接呼叫：`GET /Report/Generate?code=IPQC_EXT_INSPECT&docNo=0040`

### 3. `/ReportConfig/Index`：報表設定管理

清單顯示所有設定，可「新增 / 編輯 / 複製 / 刪除」。

進入「編輯」(`/ReportConfig/Edit/{code}`) 後：

1. **範本檔案**：可從既有範本下拉選擇，或上傳新的 `.xlsx`。上傳後立即呼叫後端掃描 API，顯示範本內所有 `{{xxx}}` 標記（一般欄位 / 集合欄位）。
2. **資料來源**：填寫 URL（可用 `{docNo}` 這種佔位符對應到參數 key）、Method、自訂 Headers、逾時秒數。
3. **參數定義**：新增/移除輸入參數（key/label/type/required），供 `/Report/Index` 動態產生輸入欄位使用。
4. **測試連線**：依「參數定義」帶入測試值（例如 docNo=0040），按下「測試連線」會由後端代打一次資料來源 API，並把原始 JSON 顯示在頁面上，方便確認資料形狀與映射是否對得上。
5. **映射設定**：
   - 根節點 JSONPath（`root`，例如 `$.data`）
   - 一般欄位：每個範本標記一行，填入對應 JSONPath；**留空的話，預設用 `$.{標記名稱}`**（例如 `{{docNo}}` 沒填就當作 `$.docNo`），不強制每個標記都要手動填。
   - 集合欄位：陣列來源路徑（`path`，相對於 root）、每個子欄位的 JSONPath（相對於單筆項目，留空同樣預設用 `$.{欄位名稱}`）、以及 spread 設定（陣列來源路徑 `from`、欄位前綴 `prefix`、最大欄數 `max`）— 只要 spread 設定了，範本中符合 `{prefix}{1..max}` 命名的欄位（例如 `v1`~`v13`）會自動由 spread 產生，不需要也不會出現在「需個別映射」的清單中。
   - 所有欄位值一律轉成文字寫入儲存格（不分型別）。若範本標記獨佔一個儲存格且取到的值長得像數字（例如 `"0040"`），MiniExcel 套版時仍會依內容自動判斷成數值而遺失前導零／改變格式——這是 MiniExcel 本身的行為，與我們傳入的 .NET 型別無關，本專案目前不會另外處理。
   - **圖片欄位**：範本裡用方括號語法 `[[標記]]`（與一般 `{{標記}}` 不同）。JSONPath 以 `$` 開頭時，對應的值可以是圖片 URL 或 base64 內容（可用 `data:image/...;base64,` 前綴，會自動去除）；不是 `$` 開頭則直接當成固定值使用（例如寫死一個圖片 URL，不透過資料來源）。「來源型別」選 `auto` 會依內容自動判斷、也可以明確指定 `url` 或 `base64`。找不到節點、抓取或解碼失敗時該張圖片直接略過（標記文字會清空），不會擋下報表產生。MiniExcel 完全不處理 `[[]]` 語法，是套版完成後由 `ReportEngine` 另外用 OpenXML 掃描、下載/解碼並嵌入圖片。
6. 按「儲存設定」即可，不會檢查是否每個標記都已映射——留空的一般/集合欄位在產生報表時會自動用預設路徑嘗試解析，找不到節點就寫入空字串，不會擋下儲存或報表產生。

## 專案結構

```
Controllers/
  MockDataController.cs      模擬外部檢驗資料來源 API
  ReportController.cs        /Report/Index、/Report/Generate
  ReportConfigController.cs  /ReportConfig/*、範本上傳掃描、測試連線
Services/
  IReportEngine.cs / ReportEngine.cs          核心：取數 → JSONPath 映射 → 集合攤平 → 套版
  IReportConfigStore.cs / ReportConfigStore.cs 設定檔的磁碟 JSON 存取
  IDataSourceClient.cs / DataSourceClient.cs   代打外部資料來源 API（逾時/非 2xx 錯誤處理）
  IReportTemplateScanner.cs / ReportTemplateScanner.cs  掃描範本 {{}} 標記
  TemplateSeeder.cs          啟動時用 OpenXML 產生示範範本
  ConfigSeeder.cs            啟動時 seed 示範設定
Models/                      設定檔、掃描結果、Mock 資料、視圖模型等 POCO
Views/Report, Views/ReportConfig, Views/Shared
App_Data/
  templates/                 範本 xlsx
  configs/                   報表設定 JSON
  logs/                      每次產生報表的紀錄
```

`ReportEngine` 完全獨立於 Controller、只依賴介面（`IReportConfigStore`、`IDataSourceClient`），日後要整包搬到別的專案，只需搬動 `Services/` 與 `Models/` 中對應的檔案即可。

## 已知限制（prototype 範圍）

- 映射設定為純文字 JSONPath 輸入，無拖拉樹狀 UI。
- 編輯頁在「選擇既有範本」切換範本時會重新整理頁面（伺服端重新掃描），尚未儲存的其他欄位編輯會遺失；上傳新範本則是純前端 AJAX 更新，不會遺失其他欄位內容。
- 所有一般/集合欄位一律當文字處理（不分型別）。範本作者若讓某個標記獨自佔滿整個儲存格且取到的值恰好長得像數字（例如 `"0040"`），MiniExcel 套版時仍會依內容自動轉型成數值而遺失前導零／改變顯示格式；這是 MiniExcel 本身讀取儲存格內容做的判斷，跟我們傳入的 .NET 型別無關，目前沒有機制可以關閉（過去用過在標記後方多放一個保留標記讓儲存格「內含多個標記」來規避，但已依需求移除）。若特定欄位的原始格式很重要（例如帶前導零的單據號、超過安全整數範圍的長數字 ID），要注意這個風險。
- 不支援 `date` 型別的格式化，日期時間一律以來源 API 回傳的原始文字寫入儲存格。
- 圖片欄位為固定尺寸嵌入（約 120x70px，不依來源圖片實際尺寸/長寬比縮放），只支援 URL 或 base64 兩種來源；示範範本目前只在右上角保留一個 `[[photo]]` 標記，若要多張圖片需自行在範本中加入更多 `[[標記]]` 並在映射設定畫面對應。
- MockDataController 僅供本機自我測試使用，非正式資料來源。
