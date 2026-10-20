import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
import io
from PIL import Image

# --- 頁面設定 ---
st.set_page_config(page_title="營運部會議紀錄生成器", page_icon="📝", layout="centered")

st.title("📝 東隧營運部 - 會議紀錄自動生成器")
st.markdown("上載手寫筆記照片，AI 會自動幫你轉換成標準格式的會議紀錄 Word 檔。")

# --- 側邊欄：設定 API Key ---
with st.sidebar:
    st.header("⚙️ 設定")
    api_key = st.text_input("請輸入你的 Google Gemini API Key", type="password")

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
uploaded_file = st.file_uploader("上載手寫筆記圖片 (支援 JPG, PNG)", type=['png', 'jpg', 'jpeg'])
agenda_text = st.text_area("有冇特定議程次序或補充資料？(選填)", height=100, placeholder="例如：1. 人手安排 2. 日常運作 3. 安全事項...")

if st.button("🚀 立即生成會議紀錄"):
    if not api_key:
        st.error("請先在左側欄位輸入 Gemini API Key！")
    elif not uploaded_file:
        st.error("請先上載手寫筆記圖片！")
    else:
        try:
            with st.spinner("AI 正在努力辨識及整理中，請等大約半分鐘..."):
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-1.5-pro')
                image = Image.open(uploaded_file)
                
                prompt = f"""
                你是一個專業的行政秘書。請根據提供的圖片內容（手寫筆記），整理成一份極度專業的會議紀錄。
                請依照以下格式及要求：
                1. 標題為「泰富基建有限公司 Pacific Infrastructure Ltd 東區海底隧道 營運部會議紀錄」
                2. 必須包含：日期、時間、地點、主席、出席者、秘書。
                3. 使用專業術語，將俗稱轉為正式職銜（例如 TM = 隧道經理，主 = 主席，OC = 營運主任，AOC = 助理營運主任）。
                4. 確保所有人名及員工編號正確。
                5. 必須分段清晰，並使用條列式（Bullet points）列出重點。
                
                附加參考/議程資訊：
                {agenda_text}
                """
                
                response = model.generate_content([prompt, image])
                result_text = response.text
                
                st.success("✅ 會議紀錄生成成功！")
                
                st.subheader("📄 預覽結果")
                st.markdown(result_text)
                
                word_file = create_word_doc(result_text)
                st.download_button(
                    label="📥 下載 Word 檔 (.docx)",
                    data=word_file,
                    file_name="營運部會議紀錄.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"發生錯誤：請檢查 API Key 是否正確，或稍後再試。系統訊息：{e}")
