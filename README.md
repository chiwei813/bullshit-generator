# Bullshit 論文生成器 (Bullshit Paper Generator)

一個基於 AI 生成技術的學術論文自動生成工具，可以根據用戶提供的主題快速生成結構完整的學術論文。

## 功能特點

- **快速生成**: 只需輸入研究主題，幾分鐘內即可生成完整論文
- **學術格式**: 自動遵循學術論文寫作規範和格式
- **自動引用**: 自動搜尋和引用相關學術文獻
- **完整章節**: 包含摘要、緒論、研究方法、研究結果、討論與建議等完整章節
- **專業排版**: 輸出為格式精美的 Word 文檔(.docx)

## 系統需求

- Python 3.8 或以上
- 與網路連接（用於 API 呼叫和文獻搜尋）
- Google Gemini API Key

## 安裝步驟

1. **克隆專案到本地**

```bash
git clone https://github.com/your-username/Bullshit-paper-generator.git
cd Bullshit-paper-generator
```

2. **安裝所需套件**

```bash
pip install -r requests.txt
```

3. **設定 API Key**

在 `api.txt` 檔案中設定您的 Google Gemini API Key：

```
GEMINI_API_KEY=您的_API_KEY
```

您可以從 https://makersuite.google.com/app/apikey 獲取 API Key。

## 使用方法

### 使用網頁界面

1. **啟動網頁服務**

```bash
python app.py
```

2. **打開瀏覽器**

訪問 `http://localhost:5000` 

3. **生成論文**
   - 在輸入框中輸入研究主題
   - 點擊「生成論文」按鈕
   - 等待進度條完成
   - 點擊「下載論文」獲取生成的 Word 文檔

### 使用 Python 腳本

您也可以直接在 Python 程式中使用論文生成功能：

```python
from paper_generator import PaperGenerator

generator = PaperGenerator()
filepath = generator.generate_paper("您的研究主題")
print(f"論文已生成: {filepath}")
```

## 檔案結構

```
├── api.txt                # API Key 設定檔
├── app.py                 # Flask 網頁應用程式
├── paper_generator.py     # 論文生成核心程式
├── requests.txt           # 所需套件清單
├── static/                # 靜態資源文件夾
│   └── style.css          # 網頁樣式
└── templates/             # 網頁模板文件夾
    └── index.html         # 首頁模板
```

## 工作原理

1. **資料收集**: 使用 Google Gemini API 和 Google Scholar API 搜尋與主題相關的最新學術文獻
2. **論文生成**: 使用 Gemini API 根據提供的主題和參考文獻生成各個章節內容
3. **文獻引用**: 自動添加適當的引用和參考文獻
4. **文檔排版**: 使用 python-docx 套件生成格式規範的學術論文 Word 文檔

## 注意事項

- 本工具生成的內容僅供參考和學習研究使用
- 請勿將生成的內容直接用於學術發表或提交作為學術作業
- API 使用受到 Google Gemini 的使用政策限制，過度使用可能導致配額限制

## 許可證

[MIT License](LICENSE)

## 貢獻與反饋

歡迎提出建議、報告問題或提交 Pull Request 來改進此專案。