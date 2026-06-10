import streamlit as st
import pandas as pd

st.set_page_config(page_title="急診自動排班系統", layout="wide")
st.title("🏥 急診護理人員自動排班系統 (自訂義介面版)")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    training_file = st.file_uploader("📂 1. 上傳班表 (training 檔案)", type=['xlsx', 'csv'])
with col2:
    template_file = st.file_uploader("📂 2. 上傳空白檔", type=['xlsx', 'csv'])

if training_file and template_file:
    try:
        if 'csv' in training_file.name.lower():
            df_shift = pd.read_csv(training_file)
        else:
            df_shift = pd.read_excel(training_file)
            
        if 'csv' in template_file.name.lower():
            df_template = pd.read_csv(template_file)
        else:
            df_template = pd.read_excel(template_file)
            
        # 重新命名基礎欄位
        base_columns = {df_shift.columns[1]: '組別', df_shift.columns[2]: '性別', df_shift.columns[3]: '姓名'}
        df_shift = df_shift.rename(columns=base_columns)
        
        # 取得所有人員名單，供後續選單使用
        all_staff = df_shift['姓名'].dropna().unique().tolist()
        
        st.success("✅ 檔案讀取成功！請至左側邊欄設定本月訓練與組長規則。")
        
        # ==========================================
        # 🌟 左側邊欄動態設定控制台
        # ==========================================
        st.sidebar.header("⚙️ 本月特殊排班規則設定")
        
        st.sidebar.subheader("🎓 訓練人員綁定 (整月固定區)")
        train_s2 = st.sidebar.multiselect("S2 訓練名單", all_staff)
        train_b1 = st.sidebar.multiselect("B1/R/C2/S 訓練名單", all_staff)
        train_t = st.sidebar.multiselect("檢傷(T) 訓練名單", all_staff)
        
        st.sidebar.markdown("---")
        st.sidebar.subheader("👑 當月組長 (L) 第一順位設定")
        st.sidebar.caption("若該組有人休假，系統將自動尋找次順位")
        l_chain_1 = st.sidebar.selectbox("第一組首選", ["--預設(黃麗婷)--"] + all_staff)
        l_chain_2 = st.sidebar.selectbox("第二組首選", ["--預設(蕭惠澤)--"] + all_staff)
        l_chain_3 = st.sidebar.selectbox("第三組首選", ["--預設(許慧芳)--"] + all_staff)
        
        # ==========================================
        # 排班運算按鈕
        # ==========================================
        st.markdown("---")
        if st.button("🚀 開始自動排班運算 (套用上述規則)"):
            st.info("💡 這裡將會接上我們之前討論的演算法，並優先將你剛才選取的訓練人員填入對應區域。")
            st.write(f"系統收到指令：S2訓練有 {len(train_s2)} 人，檢傷訓練有 {len(train_t)} 人...")
            
    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
else:
    st.info("💡 請上傳檔案以解鎖規則設定介面。")