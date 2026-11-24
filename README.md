# MP3 會議記錄產生器

這是一個 Streamlit Web App，可將上傳的 MP3 會議錄音轉為文字，
並自動整理成繁體中文、有標點、分段與會議記錄格式的 Word 檔。

## 功能

- 上傳 MP3 檔案
- 使用 OpenAI Whisper API 轉錄語音
- 將逐字稿整理為繁體中文
- 自動加上標點符號與分段
- 產生會議記錄，包括：
  - 會議資訊
  - 會議目的
  - 討論重點
  - 決議事項
  - 待辦事項
  - 逐字稿原文
- 產生並下載 Word (.docx) 檔案

## 部署方式

1. 將本專案上傳到 GitHub。
2. 前往 Streamlit Community Cloud。
3. 選擇「New app」。
4. 選擇你的 GitHub repository。
5. 選擇 `app.py`。
6. 在 Streamlit Cloud 的 Secrets 中設定：

```toml
OPENAI_API_KEY = "你的_OpenAI_API_Key"
```

7. 部署完成後即可使用。

## 注意事項

- 本 App 需要 OpenAI API Key。
- API Key 不應寫在程式碼或公開的 GitHub 檔案中。
- 建議使用 Streamlit Cloud Secrets 管理 API Key。
