import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日實戰班表生成器", layout="wide")
st.title("🖨️ 第二階段：產出每日實戰班表 (直向 A4)")
st.markdown("---")

# 這裡加入一個萬用讀取函數，解決 xls 的相容性問題
def safe_read_excel(file):
    try:
        return pd.read_excel(file, engine='openpyxl')
    except:
        # 如果 openpyxl 讀不動 (可能是舊版 xls)，嘗試用更底層的讀取方式
        return pd.read_excel(file)

st.markdown("### 📂 上傳來源檔案")
col1, col2, col3 = st.columns(3)
with col1:
    training_file = st.file_uploader("📂 1. 原始班表 (Training)", type=['xlsx', 'csv'])
with col2:
    roster_file = st.file_uploader("📂 2. 排班結果 (含 HN/工讀生)", type=['xlsx', 'csv'])
with col3:
    aide_file = st.file_uploader("📂 3. 護佐當月班表", type=['xls', 'xlsx', 'csv'])

if training_file and roster_file and aide_file:
    try:
        # 讀取檔案
        df_train = pd.read_csv(training_file) if 'csv' in training_file.name.lower() else safe_read_excel(training_file)
        df_roster = pd.read_csv(roster_file) if 'csv' in roster_file.name.lower() else safe_read_excel(roster_file)
        # 護佐班表特別處理 header=None
        df_aide_raw = pd.read_excel(aide_file, header=None)

        # 清理邏輯... (維持原本的程式碼結構)
        shift_cols = list(df_train.columns)
        shift_cols[0] = '編號'
        shift_cols[1], shift_cols[2], shift_cols[3] = '組別', '性別', '姓名'
        date_length = len(shift_cols) - 4
        date_cols = [str(i) for i in range(1, date_length + 1)]
        df_train.columns = shift_cols
        
        train_shift_map = {}
        for _, row in df_train.iterrows():
            name = str(row['姓名']).strip()
            train_shift_map[name] = {str(d): str(row.get(str(d), "")).strip().upper() for d in date_cols}

        df_roster = df_roster[~df_roster['姓名'].isin(['各班獨立人數', 'D', 'E', 'N'])]
        df_roster['姓名'] = df_roster['姓名'].astype(str).str.strip()

        # 護佐解析器
        name_row_idx, name_col_idx = None, None
        for i in range(min(15, len(df_aide_raw))):
            for j in range(min(15, len(df_aide_raw.columns))):
                if str(df_aide_raw.iloc[i, j]).strip() == '姓名':
                    name_row_idx, name_col_idx = i, j
                    break
            if name_row_idx is not None: break

        aide_dict = {}
        if name_row_idx is not None:
            date_col_map = {}
            for j in range(name_col_idx + 1, len(df_aide_raw.columns)):
                val = str(df_aide_raw.iloc[name_row_idx - 1, j]).replace('.0', '').strip()
                if val.isdigit(): date_col_map[val] = j
            for i in range(name_row_idx + 1, len(df_aide_raw)):
                aide_name = str(df_aide_raw.iloc[i, name_col_idx]).strip()
                if aide_name and aide_name.lower() != 'nan':
                    aide_dict[aide_name] = {d: str(df_aide_raw.iloc[i, c]).strip().upper() for d, c in date_col_map.items()}

        if st.button("🚀 開始生成每日實戰班表"):
            output_cols = []
            for i in range(1, 9): output_cols.extend([f'Name_{i}', f'Zone_{i}'])
            final_roster_data = []

            for day in date_cols:
                final_roster_data.append({k: (f"=== 6/{day} ===" if k == 'Name_1' else "") for k in output_cols})
                for shift_code, time_label in [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]:
                    final_roster_data.append({'Name_1': f"【{shift_code}班】", 'Zone_1': time_label})
                    staff_list = []
                    # 護理人員
                    for _, row in df_roster.iterrows():
                        name, zone = row['姓名'], str(row.get(int(day), "")).strip()
                        shift = train_shift_map.get(name, {}).get(day, "")
                        if shift_code in shift: staff_list.append((name, zone if zone != 'nan' else ""))
                    # 護佐
                    for a_name, sch in aide_dict.items():
                        if shift_code in sch.get(day, ""): staff_list.append((a_name, "護佐"))
                    
                    for i in range(0, len(staff_list), 8):
                        chunk = staff_list[i:i+8]
                        name_row, zone_row = {k: "" for k in output_cols}, {k: "" for k in output_cols}
                        for idx, (n, z) in enumerate(chunk):
                            name_row[f'Name_{idx+1}'] = n
                            zone_row[f'Zone_{idx+1}'] = z
                        final_roster_data.extend([name_row, zone_row, {k: "" for k in output_cols}])

            df_final = pd.DataFrame(final_roster_data, columns=output_cols)
            st.success("✅ 產出成功！")
            st.dataframe(df_final)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer: df_final.to_excel(writer, index=False)
            st.download_button("📥 下載 Excel", data=output.getvalue(), file_name="每日實戰班表.xlsx")

    except Exception as e:
        st.error(f"發生錯誤: {e}")
