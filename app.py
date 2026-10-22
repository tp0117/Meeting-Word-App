import streamlit as st
from google import genai
from google.genai import types
import time
from docx import Document
from docx.shared import Pt
import io
import os
import tempfile
from PIL import Image

# --- 頁面設定 ---
st.set_page_config(page_title="營運部會議紀錄生成器", page_icon="📝", layout="centered")

st.title("📝 東隧營運部 - 會議紀錄自動生成器 (Pro 頂級專業版)")
st.markdown("支援**議程格式圖**、**多頁手寫筆記**及**會議錄音檔**。系統已鎖定使用官方最強推理模型 `gemini-1.5-pro`，確保最高準確度及排版邏輯。")

# --- 側邊欄：設定與讀取 API Key ---
with st.sidebar:
    st.header("⚙️ 設定")
    
    if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ 已自動載入預設 API Key")
    else:
        api_key = st.text_input("請輸入你的 Google Gemini API Key", type="password")
        st.markdown("[按此獲取 API Key](https://aistudio.google.com/app/apikey)")

# --- 轉換 Word 檔的功能 ---
def create_word_doc(text):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.size = Pt(12)
    
    for line in text.split('\n'):
        if line.strip():
            doc.add_paragraph(line)
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 主畫面：上載與生成 ---
st.subheader("1. 上載會議資料")

uploaded_agenda_img = st.file_uploader("📋 1. 上載會議議程 / 格式次序圖片 (系統將嚴格依照此圖次序排列)", type=['png', 'jpg', 'jpeg'])
uploaded_images = st.file_uploader("📸 2. 上載手寫筆記圖片 (支援多頁選取)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
uploaded_audio = st.file_uploader("🎙️ 3. 上載會議錄音檔 (支援 MP3, M4A, WAV, AAC, MP4)", type=['mp3', 'm4a', 'wav', 'aac', 'mp4'])
agenda_text = st.text_area("✍️ 4. 有冇文字版議程或補充資料？(選填)", height=80, placeholder="例如：1. 人手安排 2. 日常運作 3. 安全事項...")

if st.button("🚀 立即生成會議紀錄 (使用 Pro 專業版)"):
    if not api_key:
        st.error("❌ 未設定 API Key，請在 Secrets 設定或左側欄位輸入！")
    elif not uploaded_agenda_img and not uploaded_images and not uploaded_audio and not agenda_text:
        st.error("❌ 請至少上載一張圖片、一個錄音檔，或輸入文字議程！")
    else:
        try:
            status_placeholder = st.empty()
            status_placeholder.info("⏳ 正在連接 AI 系統 (Gemini 1.5 Pro)，請稍候...")
            
            clean_api_key = str(api_key).strip().strip('"').strip("'")
            
            # 設定 15 分鐘超時上限，確保 Pro 模型有足夠時間思考長錄音
            try:
                client = genai.Client(
                    api_key=clean_api_key,
                    http_options=types.HttpOptions(timeout=900000)
                )
            except Exception:
                client = genai.Client(api_key=clean_api_key)
            
            contents = []
            
            # 提示詞 Prompt
            prompt_text = """
            你是一個專業的行政秘書。請根據提供的內容（會議議程格式圖片、手寫筆記圖片、會議錄音，或文字議程），整理成一份極度專業的會議紀錄。

            請嚴格依照以下格式及要求：
            1. 標題格式：
               泰富基建有限公司 Pacific Infrastructure Ltd
               東區海底隧道 營運部會議紀錄
            2. 必須包含標頭資訊：日期、時間、地點、主席、出席者、秘書。
            3. 職銜及用語規範：
               - TM = 隧道經理 (關炳輝先生)
               - 主席/副總經理/土木工程及營運總監 = 韓俊文先生
               - COC = 總營運主任 (陳智彬先生)
               - OC = 營運主任
               - AOC = 助理營運主任 (如 AOC 170 鍾建基先生、AOC 912 陶國強先生、AOC 潘子乾等)
               - 員工代號 (如 TO 2178 等) 請準確紀錄。
            4. 內容排列核心要求：
               - 【最高優先級】如果有提供「會議議程/格式次序圖片」，必須**完全依照該圖片內的章節次序、目錄編號與結構**來輸出會議紀錄內容。
               - 請將手寫筆記及錄音檔中討論的內容，準確歸納並填入對應的議程章節之下。
               - 語氣必須正式且精練。
               - 如有私人敏感問題（如欠債），請修飾為「私人問題」。
               - 使用條列式（Bullet points）列出每條重點。
            
            補充參考/文字議程資訊：
            """ + agenda_text
            
            contents.append(prompt_text)
            
            if uploaded_agenda_img:
                agenda_image = Image.open(uploaded_agenda_img)
                contents.append("【以下這張圖片是會議紀錄的標準議程次序與格式要求，請務必完全依照此圖片的章節架構排列內容】：")
                contents.append(agenda_image)
            
            if uploaded_images:
                contents.append("【以下是會議手寫筆記圖片內容】：")
                for img_file in uploaded_images:
                    image = Image.open(img_file)
                    contents.append(image)
            
            temp_audio_path = None
            if uploaded_audio:
                status_placeholder.info("⏳ 正在上載錄音檔至 Google 伺服器...")
                ext = os.path.splitext(uploaded_audio.name)[1]
                if not ext:
                    ext = ".m4a"
                    
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_file:
                    tmp_file.write(uploaded_audio.getvalue())
                    temp_audio_path = tmp_file.name
                
                audio_file_ref = client.files.upload(file=temp_audio_path)
                
                status_placeholder.info("⏳ 錄音檔已上載，正在等候 Google 伺服器完成音訊轉碼，請稍等約 10-20 秒...")
                while audio_file_ref.state.name == "PROCESSING":
                    time.sleep(2)
                    audio_file_ref = client.files.get(name=audio_file_ref.name)
                
                if audio_file_ref.state.name == "FAILED":
                    raise Exception("錄音檔上載轉碼失敗，請重新嘗試。")
                
                contents.append(audio_file_ref)

            st.subheader("📄 預覽結果")
            preview_placeholder = st.empty()
            
            # 正式鎖定使用 API 官方最高級別的 gemini-1.5-pro 模型
            candidate_models = ['gemini-1.5-pro']
            result_text = ""
            last_err = None
            
            for model_name in candidate_models:
                success = False
                for attempt in range(1, 4):
                    try:
                        status_placeholder.info(f"🤖 正在使用 {model_name} 深度分析並生成會議紀錄... (嘗試 {attempt}/3)")
                        response_stream = client.models.generate_content_stream(
                            model=model_name,
                            contents=contents
                        )
                        
                        result_text = ""
                        for chunk in response_stream:
                            if chunk.text:
                                result_text += chunk.text
                                preview_placeholder.markdown(result_text)
                        
                        if result_text:
                            success = True
                            break
                    except Exception as err:
                        last_err = err
                        err_str = str(err)
                        if "503" in err_str or "high demand" in err_str.lower() or "UNAVAILABLE" in err_str:
                            status_placeholder.warning(f"⚠️ Google 伺服器繁忙，正在等候 3 秒後自動重試 ({attempt}/3)...")
                            time.sleep(3)
                        else:
                            break 
                
                if success and result_text:
                    break
            
            if temp_audio_path and os.path.exists(temp_audio_path):
                os.remove(temp_audio_path)
            
            status_placeholder.empty()
            
            if not result_text:
                raise Exception(f"無法生成內容，請稍後再試。最後錯誤細節：{last_err}")
            
            st.success("✅ 會議紀錄生成成功！(已成功使用 Pro 專業版模型)")
            
            word_file = create_word_doc(result_text)
            st.download_button(
                label="📥 下載 Word 檔 (.docx)",
                data=word_file,
                file_name="東隧營運部會議紀錄_Pro版.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"發生錯誤：{e}")
