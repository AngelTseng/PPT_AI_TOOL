# PPT AI Tool

以 **Streamlit + OpenAI + PowerPoint COM** 為核心的投影片生成工具。  
此專案可把文字需求或既有 PPT 內容轉成結構化 `deck_spec`，再依公司模板輸出 `.pptx`。

---

## 1) 系統定位（前端 / 後端）

雖然此專案是單一 Python repo，但可以分成兩層：

- **前端（UI 層）**：`UI.py`（Streamlit）
  - 提供三種操作模式：
    1. 從 Prompt 生成 PPT
    2. 上傳既有 PPT 後美化
    3. 上傳 Spec JSON 直接渲染
  - 顯示進度、錯誤訊息、下載按鈕與結果區塊。

- **後端（服務邏輯層）**：多個 Python 模組
  - LLM 生成 / 美化：`llm_generate_spec.py`, `llm_beautify_spec.py`
  - 規格正規化：`spec_normalizer.py`
  - 規格驗證：`spec_validator.py`
  - 實際渲染 PPT：`ppt_renderer.py`（透過 Windows PowerPoint COM）
  - 版型註冊：`slide_registry.py`
  - 既有 PPT 解析：`extract_ppt_content.py`

> 簡單說：`UI.py` 負責互動，其他模組負責資料處理與輸出。

---

## 2) 架構總覽

```text
[User / Browser]
      |
      v
[Streamlit UI: UI.py]
      |
      +--> generate_spec() / beautify_spec()  (OpenAI)
      |
      +--> normalize_beautified_spec()
      |
      +--> validate_deck_spec()
      |
      +--> render_deck()  (PowerPoint COM on Windows)
      v
 [Generated PPTX + JSON artifacts]
```

### 2.1 核心資料模型：`deck_spec`

大部分流程都圍繞一個 JSON：

- `slides: []`
- 每張 slide 含 `type`（例如 `cover`, `section`, `content_2_a`, `content_4_a`, `flow`, `table`, `end`）
- 不同 type 需要對應欄位（例如 `cards`, `steps`, `columns/rows`）

這個格式同時是：
- LLM 的輸出目標
- validator 的檢查對象
- renderer 的渲染輸入

---

## 3) 前端（UI）說明

### 檔案
- `UI.py`

### 功能
- 以 Streamlit 建立操作界面
- 提供結果下載與狀態顯示
- 可顯示中間 JSON（debug 模式）
- 內建三種工作流（Generate / Beautify / Render from JSON）

### 輸出
- 所有 UI 產物預設輸出到：`ui_outputs/`

---

## 4) 後端模組分工

### 4.1 LLM Spec 生成
- `llm_generate_spec.py`
- 依 `template_map.json` + 提示詞產生 `deck_spec`
- 使用 `response_format=json_schema` 限制輸出形狀

### 4.2 LLM Spec 美化
- `llm_beautify_spec.py`
- 將解析出的舊 PPT spec 轉為更規範、更可渲染的 spec

### 4.3 Spec 正規化
- `spec_normalizer.py`
- 清理字串、補齊缺欄、統一型別變體
- 強化卡片內容的完整度
- **保證投影片序列包含首頁 `cover` 與尾頁 `end`**

### 4.4 Spec 驗證
- `spec_validator.py`
- 檢查必要欄位
- 回傳 `errors` / `warnings`
- 在渲染前擋掉不合法 spec

### 4.5 PPT 渲染
- `ppt_renderer.py`
- 將 spec 套入模板的命名 shape
- 透過 `win32com.client` 呼叫本機 PowerPoint
- 進行文字填充、表格填充、SmartArt 步驟填充

### 4.6 模板配置
- `slide_registry.py`：slide type 對應模板頁索引與欄位要求
- `template_map.json`：模板 shape 名稱摘要（供生成提示參考）

---

## 5) 環境需求

## OS / Runtime
- **Windows**（必要，因為渲染依賴 Microsoft PowerPoint COM）
- Python 3.10+
- Microsoft PowerPoint（桌面版，需可被 COM 呼叫）

## Python 套件
見 `requirements.txt`：
- `streamlit`
- `pywin32`
- `python-pptx`
- `pydantic`
- `openai`
- `tqdm`

> 注意：若只在 Linux 容器內跑，通常可做 spec 相關流程，但 **PPT COM 渲染無法完整運作**。

---

## 6) 環境變數

至少需要：

- `OPENAI_API_KEY`：OpenAI SDK 使用
- `OPENAI_MODEL`（可選）
  - 預設在 `config.py` 中是：`gpt-4.1-mini`

PowerShell 範例：

```powershell
$env:OPENAI_API_KEY="sk-..."
$env:OPENAI_MODEL="gpt-4.1-mini"
```

CMD 範例：

```cmd
set OPENAI_API_KEY=sk-...
set OPENAI_MODEL=gpt-4.1-mini
```

---

## 7) 安裝與啟動

## 7.1 安裝依賴

```bash
pip install -r requirements.txt
```

或使用專案提供 bat：
- `insatll.bat`（檔名目前如此拼寫）

## 7.2 啟動 UI

```bash
python -m streamlit run UI.py
```

或使用：
- `run_ui.bat`

預設網址：
- `http://localhost:8501`

---

## 8) 使用流程（建議）

1. 準備 `template.pptx`（含正確 shape 命名）
2. 啟動 UI
3. 選擇模式：
   - Prompt 生成
   - 上傳 PPT 美化
   - 上傳 JSON 渲染
4. 若有警告，先檢視中間 spec
5. 下載輸出 PPT 與 JSON

---

## 9) 模板設計重點

本專案高度依賴 shape 名稱對映，請確保模板中關鍵 shape 名稱固定，例如：
- `Topic`, `speaker_name`
- `agenda_name`, `title`, `item_1`, `content_1`...
- `flow_chart_1`, `sheet_1` 等

命名若改動，請同步調整：
- `slide_registry.py`
- renderer / extractor 對應邏輯
- `template_map.json`（建議重建）

---

## 10) 常見問題（FAQ）

### Q1: 為什麼渲染失敗？
最常見原因：
- 非 Windows 環境
- 未安裝 PowerPoint
- 模板 shape 名稱與程式預期不一致

### Q2: 為什麼有些內容看起來太短或空？
- spec 會先經過 normalizer/validator
- 若來源資料太少，系統會盡量補齊，但仍建議在 prompt 或 input spec 提供更完整內容

### Q3: 可以只當 JSON 規格處理器使用嗎？
可以。你可以只使用 `generate_spec / beautify_spec / normalize / validate`，不做 COM 渲染。

---

## 11) 主要檔案索引

- `UI.py`：Streamlit 前端入口
- `config.py`：模型與共用設定
- `llm_generate_spec.py`：從 prompt 生成 spec
- `llm_beautify_spec.py`：美化既有 spec
- `spec_normalizer.py`：規格正規化
- `spec_validator.py`：規格驗證
- `ppt_renderer.py`：PowerPoint 渲染
- `extract_ppt_content.py`：既有 PPT 解析
- `slide_registry.py`：slide type 與模板索引映射
- `template_map.json`：模板 shape 摘要

---

## 12) 後續可擴充方向

- 將 COM 渲染抽象化，支援非 Windows 後端渲染策略
- 將 validator 規則拆分為可配置策略檔
- 新增單元測試（normalizer/validator/renderer 的 mock 測試）
- 增加 CI（lint + test + schema checks）

