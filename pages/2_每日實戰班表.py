import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("🖨️ 每日實戰班表產生器 (關鍵字對位版)")

col1, col2, col3 = st.columns(3)
with col1: t_file = st.file_uploader("📂 1. 原始班表", type=['xlsx', 'csv'])
with col2: r_file = st.file_uploader("📂 2. 排班結果", type=['xlsx', 'csv'])
with col3: a_file = st.file_uploader("📂 3. 護佐班表", type=['xlsx', 'csv'])

if t_file and r_file and a_file:
    try:
        # 自動偵測並讀取資料
        df_t = pd.read_csv(t_file) if 'csv' in t_file.name.lower() else pd.read_excel(t_file)
        df_r = pd.read_csv(roster_file) if 'csv' in roster_file.name.lower() else pd.read_excel(roster_file)
        df_a = pd.read_excel(a_file)

        # 【核心關鍵】去除欄位前後空格，並強制將「姓名」欄位對齊
        for df in [df_t, df_r]:
            df.columns = df.columns.str.strip()
        
        # 取得日期 (假設欄位名稱為純數字)
        date_cols = [str(c) for c in df_t.columns if str(c).strip().isdigit()]
        
        if st.button("🚀 產出報表"):
            # 建立區域對照表
            zone_map = {str(row['姓名']).strip(): {str(d): str(row.get(d, "")).strip() for d in date_cols} for _, row in df_r.iterrows()}
            
            output_rows = []
            for d in date_cols:
                output_rows.append({"Name_1": f"=== 6/{d} ==="})
                for shift, time in [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]:
                    output_rows.append({"Name_1": f"【{shift}班】", "Zone_1": time})
                    
                    staff_list = []
                    # 護理師
                    for _, row in df_t.iterrows():
                        if shift in str(row.get(d, "")):
                            name = str(row['姓名']).strip()
                            zone = zone_map.get(name, {}).get(d, "")
                            staff_list.append((name, zone if zone != 'nan' else ""))
                    
                    # 護佐 (自動抓取名為「姓名」的欄位)
                    if '姓名' in df_a.columns:
                        for _, a_row in df_a.iterrows():
                            if shift in str(a_row.get(int(d), "")):
                                staff_list.append((str(a_row['姓名']).strip(), "護佐"))
                    
                    # 填入排版
                    for i in range(0, len(staff_list), 8):
                        chunk = staff_list[i:i+8]
                        n_row = {f"Name_{idx+1}": n for idx, (n, z) in enumerate(chunk)}
                        z_row = {f"Zone_{idx+1}": z for idx, (n, z) in enumerate(chunk)}
                        output_rows.extend([n_row, z_row, {}])
            
            df_final = pd.DataFrame(output_rows)
            st.dataframe(df_final)
            
            buf = io.BytesIO()
            with pd.ExcelWriter(buf) as w: df_final.to_excel(w, index=False)
            st.download_button("📥 下載 Excel", data=buf.getvalue(), file_name="每日實戰班表.xlsx")
    except Exception as e:
        st.error(f"檔案欄位讀取失敗，請確認標題列是否有「姓名」及「1, 2, 3...」: {e}")
