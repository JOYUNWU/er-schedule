import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日實戰班表生成器", layout="wide")
st.title("🖨️ 每日實戰班表產生器 (AI自動對位版)")
st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1: t_file = st.file_uploader("📂 1. 原始班表", type=['xlsx', 'csv'])
with col2: r_file = st.file_uploader("📂 2. 排班結果 (含HN/PT)", type=['xlsx', 'csv'])
with col3: a_file = st.file_uploader("📂 3. 護佐當月班表", type=['xlsx', 'csv'])

def extract_grid(df):
    date_row_idx = -1
    start_col_idx = -1
    date_cols = {}
    
    # 自動尋找日期列 (1, 2, 3...)
    for idx in range(min(15, len(df))):
        row_vals = [str(x).strip().replace('.0', '') for x in df.iloc[idx]]
        if '1' in row_vals and '2' in row_vals and '3' in row_vals:
            date_row_idx = idx
            for c_idx, val in enumerate(df.iloc[idx]):
                v_str = str(val).strip().replace('.0', '')
                if v_str.isdigit() and 1 <= int(v_str) <= 31:
                    date_cols[v_str] = c_idx
                    if v_str == '1':
                        start_col_idx = c_idx
            break
            
    if start_col_idx == -1:
        return {}, {}
        
    # 姓名欄位永遠在 1號 的前一欄
    name_col_idx = start_col_idx - 1
    # 當月班別永遠在姓名的前一欄
    shift_col_idx = start_col_idx - 2 if start_col_idx >= 2 else -1
    
    result = {}
    for idx in range(date_row_idx + 1, len(df)):
        name = str(df.iloc[idx, name_col_idx]).strip()
        # 過濾掉無意義的列
        if name and name.lower() != 'nan' and not any(x in name for x in ['姓名', '星期', '各班獨立', '人數', '總計', '休假']):
            if name in ['D', 'E', 'N']: continue 
            
            row_data = {'shifts': {}, 'default_shift': ""}
            if shift_col_idx != -1:
                ds = str(df.iloc[idx, shift_col_idx]).strip().upper()
                row_data['default_shift'] = ds if ds != 'NAN' else ""
                
            for d_str, c_idx in date_cols.items():
                row_data['shifts'][d_str] = str(df.iloc[idx, c_idx]).strip()
                
            result[name] = row_data
    return result, date_cols

if t_file and r_file and a_file:
    try:
        df_t = pd.read_csv(t_file, header=None) if 'csv' in t_file.name.lower() else pd.read_excel(t_file, header=None)
        df_r = pd.read_csv(r_file, header=None) if 'csv' in r_file.name.lower() else pd.read_excel(r_file, header=None)
        df_a = pd.read_csv(a_file, header=None) if 'csv' in a_file.name.lower() else pd.read_excel(a_file, header=None)

        train_map, date_cols_t = extract_grid(df_t)
        roster_map, date_cols_r = extract_grid(df_r)
        aide_map, date_cols_a = extract_grid(df_a)

        sorted_days = sorted(list(date_cols_t.keys()), key=int)

        if st.button("🚀 開始轉出每日實戰班表 (6人直式完美對齊)"):
            
            # ✅ 修改點 1：將欄位簡化為單純的 6 個直行 (Col_1 到 Col_6)
            output_cols = [f'Col_{i}' for i in range(1, 7)]
            final_rows = []

            for day in sorted_days:
                final_rows.append({output_cols[0]: f"=== 本月 {day} 日 ==="})

                for shift_code, time_label in [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]:
                    final_rows.append({
                        output_cols[0]: f"【{shift_code}班】",
                        output_cols[1]: time_label
                    })

                    shift_staff_list = []

                    # A. 撈出護理同仁與工讀生
                    for name, r_data in roster_map.items():
                        zone_val = r_data['shifts'].get(day, "")
                        if zone_val.upper() in ['X', 'NAN', 'NONE', 'OFF', '']:
                            continue

                        if name in train_map:
                            today_shift = train_map[name]['shifts'].get(day, "").upper()
                        else:
                            today_shift = r_data['default_shift']

                        if shift_code in today_shift:
                            shift_staff_list.append((name, zone_val))

                    # B. 撈出護佐
                    for aide_name, a_data in aide_map.items():
                        today_aide_shift = a_data['shifts'].get(day, "").upper()
                        if shift_code in today_aide_shift:
                            shift_staff_list.append((aide_name, "護佐"))

                    # C. ✅ 修改點 2：強制名字與區域放在同一個欄位上下對齊
                    for i in range(0, len(shift_staff_list), 6):
                        chunk = shift_staff_list[i:i+6]
                        name_row = {k: "" for k in output_cols}
                        zone_row = {k: "" for k in output_cols}

                        for idx, (s_name, s_zone) in enumerate(chunk):
                            col_key = output_cols[idx] # 對準 Col_1, Col_2...
                            name_row[col_key] = s_name
                            zone_row[col_key] = s_zone

                        final_rows.append(name_row)
                        final_rows.append(zone_row)

                    final_rows.append({k: "" for k in output_cols}) # 班別間留白

                final_rows.append({output_cols[0]: "------------------------"}) # 每日中斷線

            df_final = pd.DataFrame(final_rows, columns=output_cols)
            
            st.markdown("### 👀 每日實戰班表初步預覽")
            st.dataframe(df_final)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, header=False, sheet_name='每日實戰班表')
            
            st.download_button(
                label="📥 下載實戰班表 Excel 檔案",
                data=buffer.getvalue(),
                file_name="每日實戰班表_完美對齊版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("🎉 實戰班表已成功生成！人員與區域已百分之百完美垂直對齊！")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
