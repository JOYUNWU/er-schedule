import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日實戰班表生成器", layout="wide")
st.title("🖨️ 第二階段：產出每日實戰班表 (直向 A4)")
st.markdown("---")

st.info("💡 這個頁面完全獨立！您隨時可以手動修改「排班結果」後直接上傳，不需重新跑自動排班。")

st.markdown("### 📂 上傳來源檔案")
col1, col2, col3 = st.columns(3)
with col1:
    training_file = st.file_uploader("📂 1. 原始班表 (Training)", type=['xlsx', 'csv'], help="用來對照每天誰上 D/E/N 班")
with col2:
    roster_file = st.file_uploader("📂 2. 排班結果 (含 HN/工讀生)", type=['xlsx', 'csv'], help="排班長手動修改/安插後的最終結果檔")
with col3:
    aide_file = st.file_uploader("📂 3. 護佐當月班表", type=['xls', 'xlsx', 'csv'], help="護佐專用的原汁原味排班表")

if training_file and roster_file and aide_file:
    try:
        # 1. 讀取所有檔案
        df_train = pd.read_csv(training_file) if 'csv' in training_file.name.lower() else pd.read_excel(training_file)
        df_roster = pd.read_csv(roster_file) if 'csv' in roster_file.name.lower() else pd.read_excel(roster_file)
        df_aide_raw = pd.read_excel(aide_file, header=None)

        # 2. 清理原始班表 (抓取每日 D/E/N)
        shift_cols = list(df_train.columns)
        shift_cols[1], shift_cols[2], shift_cols[3] = '組別', '性別', '姓名'
        date_length = len(shift_cols) - 4
        shift_cols[4:] = [str(i) for i in range(1, date_length + 1)]
        df_train.columns = shift_cols
        df_train = df_train.dropna(subset=['姓名'])
        df_train['姓名'] = df_train['姓名'].astype(str).str.strip()
        
        date_cols = [str(i) for i in range(1, date_length + 1)]

        # 建立快查表
        train_shift_map = {}
        for _, row in df_train.iterrows():
            name = str(row['姓名']).strip()
            train_shift_map[name] = {str(d): str(row.get(d, row.get(int(d), ""))).strip().upper() for d in date_cols}

        # 3. 清理排班結果表 (Roster)
        df_roster = df_roster[~df_roster['姓名'].isin(['各班獨立人數', 'D', 'E', 'N'])]
        df_roster = df_roster.dropna(subset=['姓名'])
        df_roster['姓名'] = df_roster['姓名'].astype(str).str.strip()

        # 4. 護佐班表 (Aide) 智慧解析器
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
                if val.isdigit():
                    date_col_map[val] = j

            for i in range(name_row_idx + 1, len(df_aide_raw)):
                aide_name = str(df_aide_raw.iloc[i, name_col_idx]).strip()
                if aide_name and aide_name.lower() != 'nan':
                    aide_dict[aide_name] = {}
                    for d_str, c_idx in date_col_map.items():
                        aide_dict[aide_name][d_str] = str(df_aide_raw.iloc[i, c_idx]).strip().upper()

        if st.button("🚀 開始生成每日實戰班表 (直向 A4)"):
            with st.spinner("正在執行直向 A4 排版..."):
                
                output_cols = []
                for i in range(1, 9):
                    output_cols.extend([f'Name_{i}', f'Zone_{i}'])

                final_roster_data = []

                for day in date_cols:
                    final_roster_data.append({k: (f"=== 6/{day} ===" if k == 'Name_1' else "") for k in output_cols})
                    final_roster_data.append({k: "" for k in output_cols})

                    shifts_info = [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]

                    for shift_code, time_label in shifts_info:
                        final_roster_data.append({
                            'Name_1': f"【{shift_code}班】",
                            'Zone_1': time_label,
                            **{k: "" for k in output_cols[2:]}
                        })

                        shift_staff_list = []

                        # A. 護理人員 (包含 HN 與工讀生)
                        for _, row in df_roster.iterrows():
                            name = row['姓名']
                            zone_val = str(row.get(day, row.get(int(day), ""))).strip()
                            if zone_val.upper() in ['X', 'NAN', 'NONE', 'OFF']:
                                zone_val = ""

                            actual_shift = ""
                            if name in train_shift_map:
                                actual_shift = train_shift_map[name].get(day, "")
                            else:
                                actual_shift = str(row.get('當月班別', '')).strip().upper()
                            
                            is_off = actual_shift in ['OFF', 'X', 'NAN', 'NONE', '']

                            if not is_off and (shift_code in actual_shift or actual_shift == shift_code):
                                shift_staff_list.append((name, zone_val))

                        # B. 護佐 (自動跟隨在後)
                        for aide_name, schedule in aide_dict.items():
                            aide_shift = schedule.get(day, "")
                            if shift_code == 'D' and ('D' in aide_shift):
                                shift_staff_list.append((aide_name, "護佐"))
                            elif shift_code == 'E' and ('E' in aide_shift):
                                shift_staff_list.append((aide_name, "護佐"))
                            elif shift_code == 'N' and ('N' in aide_shift):
                                shift_staff_list.append((aide_name, "護佐"))

                        # C. 執行 8 人換行排版
                        chunk_size = 8
                        for i in range(0, len(shift_staff_list), chunk_size):
                            chunk = shift_staff_list[i:i+chunk_size]
                            name_row = {k: "" for k in output_cols}
                            zone_row = {k: "" for k in output_cols}

                            for idx, (s_name, s_zone) in enumerate(chunk):
                                col_idx = idx + 1
                                name_row[f'Name_{col_idx}'] = s_name
                                zone_row[f'Zone_{col_idx}'] = s_zone

                            final_roster_data.append(name_row)
                            final_roster_data.append(zone_row)
                            final_roster_data.append({k: "" for k in output_cols}) 

                        final_roster_data.append({k: "" for k in output_cols}) 

                    final_roster_data.append({k: "------------------------" for k in output_cols})

                df_final = pd.DataFrame(final_roster_data, columns=output_cols)

                st.success("✅ 每日班表產出成功！已自動將護佐接續於名單後方。")
                st.dataframe(df_final.head(25)) 

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='每日實戰班表')
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 下載每日實戰班表 (Excel)",
                    data=excel_data,
                    file_name="每日實戰班表_A4直向.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"處理失敗，請檢查檔案格式是否正確：{e}")
