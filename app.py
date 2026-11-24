import os
from datetime import datetime

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================
# 頁面設定
# ============================================
st.set_page_config(
    page_title="MP3 會議記錄產生器",
    page_icon="🎤",
    layout="centered"
)

st.title("🎤 MP3 會議記錄產生器")
st.write(
    "上傳 MP3 檔案，自動轉錄為繁體中文、加上標點與分段，"
    "並產生可下載的 Word 會議記錄。"
)

# ============================================
# 讀取 OpenAI API Key
# ============================================
api_key = st.secrets.get("OPENAI_API_KEY", "")

if not api_key:
    st.error(
        "尚未設定 OpenAI API Key。"
        "請在 Streamlit Cloud 的 Secrets 中設定 OPENAI_API_KEY。"
    )
    st.stop()

client = OpenAI(api_key=api_key)

# ============================================
# 上傳 MP3
# ============================================
uploaded_file = st.file_uploader(
    "請上傳 MP3 檔案",
    type=["mp3"]
)

if uploaded_file is not None:
    st.success(f"已上傳檔案：{uploaded_file.name}")

    # 儲存上傳的 MP3
    upload_folder = "uploads"
    os.makedirs(upload_folder, exist_ok=True)

    # ============================================
    # 取得安全的英文檔名
    # 避免中文、空格或特殊符號造成 UnicodeEncodeError
    # ============================================
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()

    # 若副檔名不是 .mp3，強制設為 .mp3
    if file_extension != ".mp3":
        file_extension = ".mp3"

    safe_filename = (
        "uploaded_audio_"
        + datetime.now().strftime("%Y%m%d_%H%M%S")
        + file_extension
    )

    audio_path = os.path.join(
        upload_folder,
        safe_filename
    )

    with open(audio_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # ============================================
    # 開始轉錄
    # ============================================
    if st.button("開始轉錄並產生會議記錄"):

        with st.spinner("正在轉錄 MP3，請稍候……"):

            with open(audio_path, "rb") as audio_file:
                transcription = client.audio.transcriptions.create(
                    model="whisper-1",
                    file=audio_file,
                    response_format="text"
                )

        raw_transcript = transcription.strip()

        st.subheader("原始逐字稿")
        st.write(raw_transcript)

        # ============================================
        # 整理為會議記錄
        # ============================================
        with st.spinner(
            "正在整理為繁體中文、標點、分段與會議記錄……"
        ):

            system_prompt = """你是一位專業的繁體中文會議記錄撰寫助理。
你的任務是將會議語音逐字稿整理成清晰、專業、可閱讀的會議記錄。

規則：
1. 全部使用繁體中文。
2. 使用正確且自然的標點符號。
3. 依議題與討論內容合理分段。
4. 保留會議中實際提到的內容，不要虛構、添加或遺漏重要資訊。
5. 若逐字稿沒有提供會議名稱、日期、時間、地點、出席人員，請寫「未提供」。
6. 若無法確認是誰發言，請不要自行猜測姓名。
7. 將重複、口頭語與無意義的贅字適度刪除，但不得改變原意。
8. 將討論內容整理為重點，將已決定事項整理為決議，將後續工作整理為待辦事項。
9. 待辦事項若能判斷負責人或期限，請列出；無法判斷時請寫「待確認」。
10. 只輸出會議記錄內容，不要加入額外說明。"""

            user_prompt = f"""請將以下會議逐字稿整理成繁體中文會議記錄，格式如下：

# 會議記錄

## 一、會議資訊
- 會議名稱：未提供
- 會議日期：未提供
- 會議時間：未提供
- 會議地點：未提供
- 出席人員：未提供
- 記錄產生時間：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

## 二、會議目的
請依逐字稿內容整理會議目的；若無法判斷，請寫「未提供」。

## 三、討論重點
請依議題分段整理討論重點。
每個重點請使用條列方式呈現。

## 四、決議事項
請列出會議中已達成共識或決定的事項。
若沒有明確決議，請寫「本次會議未形成明確決議」。

## 五、待辦事項
請列出後續需要處理的事項，格式如下：
- 待辦事項：  
  負責人：待確認  
  預計完成期限：待確認  

若沒有待辦事項，請寫「本次會議未列出待辦事項」。

## 六、其他事項
請列出會議中提到但未歸類於上述項目的重要資訊。
若沒有，請寫「無」。

## 七、逐字稿原文
請將整理後的逐字稿附在後面，保留分段與標點，並全部使用繁體中文。

會議逐字稿如下：
{raw_transcript}"""

            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": system_prompt
                    },
                    {
                        "role": "user",
                        "content": user_prompt
                    }
                ],
                temperature=0.2
            )

            meeting_minutes = (
                response.choices[0].message.content.strip()
            )

        st.subheader("會議記錄")
        st.markdown(meeting_minutes)

        # ============================================
        # 產生 Word 檔
        # ============================================
        def create_word_document(
            meeting_minutes
        ):
            document = Document()

            style = document.styles["Normal"]
            style.font.name = "Microsoft JhengHei"
            style.font.size = Pt(12)

            title = document.add_heading(
                "會議記錄",
                level=0
            )
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            info = document.add_paragraph()
            info.add_run(
                "產生時間："
                f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            )

            document.add_paragraph()

            document.add_heading(
                "會議記錄內容",
                level=1
            )

            paragraphs = [
                paragraph.strip()
                for paragraph in meeting_minutes.split("\n")
                if paragraph.strip()
            ]

            for paragraph_text in paragraphs:
                document.add_paragraph(paragraph_text)

            # 產生安全的英文 Word 檔名
            word_base_name = (
                "meeting_minutes_"
                + datetime.now().strftime("%Y%m%d_%H%M%S")
            )

            word_filename = f"{word_base_name}.docx"

            document.save(word_filename)

            return word_filename

        word_filename = create_word_document(
            meeting_minutes
        )

        # ============================================
        # 下載 Word 檔
        # ============================================
        with open(word_filename, "rb") as f:
            word_bytes = f.read()

        st.download_button(
            label="下載 Word 會議記錄",
            data=word_bytes,
            file_name=word_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            )
        )

        # 刪除暫存 MP3 與 Word 檔
        os.remove(audio_path)
        os.remove(word_filename)
