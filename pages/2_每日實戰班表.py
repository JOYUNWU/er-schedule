import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日班表生成器", layout="wide")
st.title("🖨️ 每日班表生成器")
st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1: t_file = st.file_uploader("📂 1. 原始班表", type=['xlsx', 'csv'])
with col2: r_file = st.file_uploader("📂 2. 排班結果 (含HN/PT)", type=['xlsx', 'csv'])
with col3: a_file = st.file_uploader("📂 3. 護佐班表", type=['xlsx', 'csv'])

def extract_grid(df):
    date_row_idx = -1
    start_col_idx = -1
    date_cols = {}
    year_month = ""
    weekdays = {}
    
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
            
            # 💡 往前掃描，尋找「年」與「月」的字眼
            for r in range(max(0, idx-2), idx+1):
                for c in range(len(df.columns)):
                    cell_val = str(df.iloc[r, c]).strip()
                    if '年' in cell_val and '月' in cell_val:
                        year_month = cell_val
            break
            
    if start_col_idx == -1:
        return {}, {}, "", {}
        
    # 💡 往下一列掃描，尋找「星期」
    if date_row_idx + 1 < len(df):
        for d_str, c_idx in date_cols.items():
            wd = str(df.iloc[date_row_idx + 1, c_idx]).strip()
            # 清理字眼，確保輸出時不會變成 "星期星期六"
            wd = wd.replace('星期', '').replace('周', '') 
            weekdays[d_str] = wd

    # 姓名欄位與班別欄位對位
    name_col_idx = start_col_idx - 1
    shift_col_idx = start_col_idx - 2 if start_col_idx >= 2 else -1
    
    result = {}
    for idx in range(date_row_idx + 1, len(df)):
        name = str(df.iloc[idx, name_col_idx]).strip()
        if name and name.lower() != 'nan' and not any(x in name for x in ['姓名', '星期', '各班獨立', '人數', '總計', '休假']):
            if name in ['D', 'E', 'N']: continue 
            
            row_data = {'shifts': {}, 'default_shift': ""}
            if shift_col_idx != -1:
                ds = str(df.iloc[idx, shift_col_idx]).strip().upper()
                row_data['default_shift'] = ds if ds != 'NAN' else ""
                
            for d_str, c_idx in date_cols.items():
                row_data['shifts'][d_str] = str(df.iloc[idx, c_idx]).strip()
                
            result[name] = row_data
    return result, date_cols, year_month, weekdays

if t_file and r_file and a_file:
    try:
        df_t = pd.read_csv(t_file, header=None) if 'csv' in t_file.name.lower() else pd.read_excel(t_file, header=None)
        df_r = pd.read_csv(r_file, header=None) if 'csv' in r_file.name.lower() else pd.read_excel(r_file, header=None)
        df_a = pd.read_csv(a_file, header=None) if 'csv' in a_file.name.lower() else pd.read_excel(a_file, header=None)

        # 這裡會從 training.xlsx 中精準抽出年月與星期
        train_map, date_cols_t, ym_t, wd_t = extract_grid(df_t)
        roster_map, date_cols_r, _, _ = extract_grid(df_r)
        aide_map, date_cols_a, _, _ = extract_grid(df_a)

        sorted_days = sorted(list(date_cols_t.keys()), key=int)

        # 清理年月格式 (例如把 "2024年06月" 變成 "2024年6月")
        ym_str = ym_t if ym_t else "本月"
        ym_str = ym_str.replace("年0", "年")

        if st.button("🚀 開始生成每日班表"):
            
            output_cols = [f'Col_{i}' for i in range(1, 7)]
            final_rows = []

            for day in sorted_days:
                # 💡 組裝最終完美的標題：2024年6月1日 星期六
                day_wd = wd_t.get(day, "")
                if day_wd:
                    title_str = f"{ym_str}{day}日 星期{day_wd}"
                else:
                    title_str = f"{ym_str}{day}日"

                final_rows.append({output_cols[0]: title_str})

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

                    # C. 6 人一列垂直完美對齊
                    for i in range(0, len(shift_staff_list), 6):
                        chunk = shift_staff_list[i:i+6]
                        name_row = {k: "" for k in output_cols}
                        zone_row = {k: "" for k in output_cols}

                        for idx, (s_name, s_zone) in enumerate(chunk):
                            col_key = output_cols[idx] 
                            name_row[col_key] = s_name
                            zone_row[col_key] = s_zone

                        final_rows.append(name_row)
                        final_rows.append(zone_row)

                    final_rows.append({k: "" for k in output_cols}) 

                final_rows.append({output_cols[0]: "------------------------"}) 

            df_final = pd.DataFrame(final_rows, columns=output_cols)
            
            st.markdown("### 👀 每日班表初步預覽")
            st.dataframe(df_final)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, header=False, sheet_name='每日班表')
            
            st.download_button(
                label="📥 下載每日班表 Excel 檔案",
                data=buffer.getvalue(),
                file_name="每日班表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("🎉 每日班表已成功生成！")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
