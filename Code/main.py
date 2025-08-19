
# streamlit run Code/main.py
from utils import st, os
from config import save_path
from create_ppt import generate_ppt_from_report, generate_report
from ppt_draw import create_node, generate_diagram_to_ppt

def main():
    st.title("📊 分析檔案ppt產生器")

    uploaded_files = st.file_uploader("**請上傳 Python (.py) 檔案**", type=["py"], accept_multiple_files=True)  # 讓使用者上傳檔案, 支援多個檔案上傳
    
    if not uploaded_files:  # 若無檔案上傳則結束流程
        return
    
    # 之後將這些參數直接放在LLM的system prompt即可
    with st.expander("⚙️ **PPT相關設定**"):
        ppt_pages = st.slider("**請選擇要產出的 PPT 頁數**", min_value=1, max_value=20, value=5)    
        level = st.radio("**請選擇PPT的知識層級**", ["入門者", "初學者", "中階者", "高階者", "專家"], index=0)
        language = st.radio("**請選擇語言**", ["繁體中文", "English", "Japanese", "Korean"], index=0)

        if st.button("✅ 儲存設定"):
            st.success(f"已儲存設定! 🎉")

    if st.button("📥 產出簡報 PPT"):
        with st.spinner("正在分析並產生報告..."):
            all_codes = []
            for uploaded_file in uploaded_files:                
                code = uploaded_file.read().decode("utf-8")     # 讀取上傳內容

                all_codes.append(code)  # 將多個code合併
             # 加入動態狀態區塊
            status = st.empty()
            structure = generate_report(all_codes, st_status=status, num_pages=ppt_pages, level=level, language=language)

            st.json(structure)  # 顯示結構以便調試

            # 依結構生成 PPT
            generate_ppt_from_report(structure, save_path)

            nodes = create_node(structure)
            st.markdown("**以下是 nodes 部分**")
            st.json(nodes)
            generate_diagram_to_ppt(save_path, status, nodes)

        st.success(f"✅ 報表已儲存至：`{save_path}`")
        # 3️⃣ 提供下載按鈕
        with open(save_path, "rb") as f:
            st.download_button(
                "📥 下載PPT",
                f,
                file_name=os.path.basename(save_path),
            )
        
if __name__ == "__main__":
    main()


