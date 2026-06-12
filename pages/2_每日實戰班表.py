import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日實戰班表生成器", layout="wide")
st.title("🖨️ 每日實戰班表產生器 (獨立運作版)")

col1, col2, col3 = st.columns(3)
with col1: training_file = st.file_uploader("📂 1. 原始班表", type=['xlsx', 'csv'])
with col2: roster_file = st.file_uploader("📂 2. 排班結果", type=['xlsx', 'csv'])
with col3: aide_file = st.file_uploader("📂 3. 護佐班表 (.xlsx)", type=['xlsx', 'csv'])

if training_file and roster_file and aide_file:
    try:
        # 讀取資料
        df_train = pd.read_csv(training_file) if 'csv' in training_file.name.lower() else pd.read_excel(training_file)
        df_roster = pd.read_csv(roster_file) if 'csv' in roster_file.name.lower() else pd.read_excel(roster_file)
        df_aide = pd.read_excel(aide_file)
        
        # 取得日期
        date_cols = [str(c) for c in df_train.columns if str(c).isdigit()]
        
        if st.button("🚀 產出每日實戰班表"):
            # 建立區域對照表
            zone_map = {}
            for _, row in df_roster.iterrows():
                name = str(row['姓名']).strip()
                zone_map[name] = {str(d): str(row.get(d, "")).strip() for d in date_cols}
            
            output_rows = []
            for d in date_cols:
                # 建立日期區塊標題
                output_rows.append({"Name_1": f"=== 6/{d} ==="})
                
                for shift, time in [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]:
                    output_rows.append({"Name_1": f"【{shift}班】", "Zone_1": time})
                    
                    staff_list = []
                    # 抓取護理人員
                    for _, row in df_train.iterrows():
                        name = str(row['姓名']).strip()
                        if str(row.get(d, "")).strip().upper() == shift:
                            zone = zone_map.get(name, {}).get(d, "")
                            staff_list.append((name, zone if zone != 'nan' else ""))
                    # 抓取護佐
                    if '姓名' in df_aide.columns:
                        for _, a_row in df_aide.iterrows():
                            if str(a_row.get(int(d), "")).strip().upper() == shift:
                                staff_list.append((str(a_row['姓名']).strip(), "護佐"))
                    
                    # 排版 (8人一列)
                    for i in range(0, len(staff_list), 8):
                        chunk = staff_list[i:i+8]
                        n_row = {f"Name_{idx+1}": n for idx, (n, z) in enumerate(chunk)}
                        z_row = {f"Zone_{idx+1}": z for idx, (n, z) in enumerate(chunk)}
                        output_rows.extend([n_row, z_row, {}])
            
            # --- 瀏覽功能 ---
            df_final = pd.DataFrame(output_rows)
            st.markdown("### 👀 初步預覽")
            st.dataframe(df_final)
            
            # --- 下載功能 ---
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 下載 Excel 檔案",
                data=buffer.getvalue(),
                file_name="每日實戰班表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ 產出完成！請點擊上方按鈕下載。")
            
    except Exception as e:
        st.error(f"發生錯誤: {e}")
