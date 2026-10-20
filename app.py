import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
import io
import os
import tempfile
from PIL import Image

# --- 頁面設定 ---
st.set_page_config(page_title="營運部會議紀錄生成器", page_icon="📝", layout="centered")

st.title("📝 東隧營運部 - 會議紀錄自動生成器")
st.markdown("支援**多頁手寫筆記圖片**及**會議錄音檔**，系統會自動生成標準 Word 會議紀錄。")

# --- 側邊欄：設定與讀取 API Key ---
with st.sidebar:
    st.header("⚙️ 設定")
    
    # 優先由 Streamlit Secrets 自動讀取 API Key
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
st.subheader("1. 上載會議資料 (可選圖片、錄音或兩者皆備)")

uploaded_images = st.file_uploader("📸 上載手寫筆記圖片 (支援多頁選取)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
uploaded_audio = st.file_uploader("🎙️ 上載會議錄音檔 (支援 MP3, M4A, WAV, AAC, MP4)", type=['mp3', 'm4a', 'wav', 'aac', 'mp4'])

agenda_text = st.text_area("📋 有冇特定議程次序或補充資料？(選填)", height=100, placeholder="例如：1. 人手安排 2. 日常運作 3. 安全事項...")

if st.button("🚀 立即生成會議紀錄"):
    if not api_key:
        st.error("❌ 未設定 API Key，請在 Secrets 設定或左側欄位輸入！")
    elif not uploaded_images and not uploaded_audio and not agenda_text:
        st.error("❌ 請至少上載一張圖片、一個錄音檔，或輸入文字議程！")
    else:
        try:
            with st.spinner("AI 正在高精度分析會議內容，請稍候..."):
                # 自動清理 API Key 前後的多餘空白及雙引號，避免格式錯誤
                clean_api_key = str(api_key).strip().strip('"').strip("'")
                genai.configure(api_key=clean_api_key)
                
                # 直接使用 Gemini 1.5 Pro（若無付費權限自動備用 1.5 Flash）
                try:
                    model = genai.GenerativeModel('gemini-1.5-pro')
                except Exception:
                    model = genai.GenerativeModel('gemini-1.5-flash')
                
                # 建立傳送給 AI 的內容清單
                prompt_content = [
                    """
                    你是一個專業的行政秘書。請根據提供的內容（手寫筆記圖片、會議錄音，或兩者結合），整理成一份極度專業的會議紀錄。
                    請依照以下格式及要求：
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
                    4. 內容要求：
                       - 語氣必須正式且精練。
                       - 如有私人敏感問題（如欠債），請修飾為「私人問題」。
                       - 按議程次序分段清晰，並使用條列式（Bullet points）列出每條重點。
                    
                    附加參考/議程資訊：
                    """ + agenda_text
                ]
                
                # 處理圖片 (支援多張)
                if uploaded_images:
                    for img_file in uploaded_images:
                        image = Image.open(img_file)
                        prompt_content.append(image)
                
                # 處理錄音檔
                temp_audio_path = None
                if uploaded_audio:
                    file_extension = os.path.splitext(uploaded_audio.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as tmp_file:
                        tmp_file.write(uploaded_audio.getvalue())
                        temp_audio_path = tmp_file.name
                    
                    audio_file = genai.upload_file(path=temp_audio_path)
                    prompt_content.append(audio_file)
                
                response = model.generate_content(prompt_content)
                result_text = response.text
                
                if temp_audio_path and os.path.exists(temp_audio_path):
                    os.remove(temp_audio_path)
                
                st.success("✅ 會議紀錄生成成功！")
                
                st.subheader("📄 預覽結果")
                st.markdown(result_text)
                
                word_file = create_word_doc(result_text)
                st.download_button(
                    label="📥 下載 Word 檔 (.docx)",
                    data=word_file,
                    file_name="東隧營運部會議紀錄.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"發生錯誤：{e}")
