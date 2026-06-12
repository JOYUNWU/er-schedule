import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日實戰班表生成器", layout="wide")
st.title("🖨️ 每日實戰班表產生器 (護佐精準鎖定版)")
st.markdown("---")

st.info("💡 運作邏輯：由「原始班表」決定誰當天上班，「排班結果」決定上班區域，「護佐班表」則依指定位置精準抓取代號。")

# 萬用 Excel 讀取函數
def safe_read(file):
    try:
        return pd.read_excel(file, header=None, engine='openpyxl')
    except:
        return pd.read_excel(file, header=None)

col1, col2, col3 = st.columns(3)
with col1: t_file = st.file_uploader("📂 1. 原始班表 (Training)", type=['xlsx'])
with col2: r_file = st.file_uploader("📂 2. 排班結果 (含HN/工讀生)", type=['xlsx'])
with col3: a_file = st.file_uploader("📂 3. 護佐當月班表", type=['xlsx'])

# 智慧型護理人員解析器
def parse_nurse_table(df):
    name_row, name_col = None, None
    for i in range(min(15, len(df))):
        for j in range(min(15, len(df.columns))):
            if "姓名" in str(df.iloc[i, j]):
                name_row, name_col = i, j
                break
        if name_row is not None: break
    
    date_map = {}
    for r in [name_row, name_row - 1]:
        if r is None or r < 0: continue
        for j in range(len(df.columns)):
            val = str(df.iloc[r, j]).replace('.0', '').strip()
            if val.isdigit():
                date_map[val] = j
                
    data_dict = {}
    for i in range(name_row + 1, len(df)):
        name = str(df.iloc[i, name_col]).strip()
        if name and name.lower() != 'nan' and not any(x in name for x in ['各班獨立', '總計', '人數', 'D', 'E', 'N']):
            data_dict[name] = {}
            for d, j in date_map.items():
                data_dict[name][d] = str(df.iloc[i, j]).strip()
    return data_dict, list(date_map.keys())

if t_file and r_file and a_file:
    try:
        df_t_raw = safe_read(t_file)
        df_r_raw = safe_read(r_file)
        df_a_raw = safe_read(a_file)

        # 1. 解析護理人員
        train_map, date_cols = parse_nurse_table(df_t_raw)
        roster_map, _ = parse_nurse_table(df_r_raw)
        
        # 確保日期排序正確 (1日、2日...30日)
        date_cols.sort(key=int)

        # 2. 依照排班長指定規則：精準解析護佐班表
        aide_date_map = {}
        date_row = df_a_raw.iloc[1]  # 第 2 列 (index 1)
        for j in range(4, len(df_a_raw.columns)):  # 從 E 欄 (index 4) 開始
            val = str(date_row.iloc[j]).replace('.0', '').strip()
            if val.isdigit():
                aide_date_map[val] = j

        aide_map = {}
        for i in range(2, len(df_a_raw)):  # 從第 3 列往下掃描
            aide_name = str(df_a_raw.iloc[i, 3]).strip()  # D 欄 (index 3) 是姓名
            if aide_name and aide_name.lower() != 'nan' and aide_name != '姓名':
                aide_map[aide_name] = {}
                for d_str, col_idx in aide_date_map.items():
                    aide_map[aide_name][d_str] = str(df_a_raw.iloc[i, col_idx]).strip().upper()

        if st.button("🚀 開始生成每日實戰班表"):
            output_cols = []
            for i in range(1, 9):
                output_cols.extend([f'Name_{i}', f'Zone_{i}'])

            final_rows = []

            for day in date_cols:
                # 每日大標題
                final_rows.append({k: (f"=== 本月 {day} 日 ===" if k == 'Name_1' else "") for k in output_cols})
                final_rows.append({k: "" for k in output_cols})

                for shift_code, time_label in [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]:
                    final_rows.append({
                        'Name_1': f"【{shift_code}班】",
                        'Zone_1': time_label,
                        **{k: "" for k in output_cols[2:]}
                    })

                    shift_staff_list = []

                    # A. 撈出該班別上班的護理人員
                    for name, dates in train_map.items():
                        today_shift = dates.get(day, "").upper()
                        if shift_code in today_shift:
                            # 去 Roster 字典找他今天被分配到哪個核定區域
                            zone_val = roster_map.get(name, {}).get(day, "")
                            if zone_val.upper() in ['X', 'NAN', 'NONE', 'OFF', '']:
                                zone_val = ""
                            shift_staff_list.append((name, zone_val))

                    # B. 撈出該班別上班的護佐 (精準鎖定)
                    for aide_name, dates in aide_map.items():
                        today_aide_shift = dates.get(day, "")
                        if shift_code in today_aide_shift:
                            shift_staff_list.append((aide_name, "護佐"))

                    # C. 執行 A4 直向 8 人換行排版
                    for i in range(0, len(shift_staff_list), 8):
                        chunk = shift_staff_list[i:i+8]
                        name_row = {k: "" for k in output_cols}
                        zone_row = {k: "" for k in output_cols}

                        for idx, (s_name, s_zone) in enumerate(chunk):
                            col_idx = idx + 1
                            name_row[f'Name_{col_idx}'] = s_name
                            zone_row[f'Zone_{col_idx}'] = s_zone

                        final_rows.append(name_row)
                        final_rows.append(zone_row)
                        final_rows.append({k: "" for k in output_cols}) # 人的下方留空一行

                    final_rows.append({k: "" for k in output_cols}) # 班別間隔

                final_rows.append({k: "------------------------" for k in output_cols}) # 每日中斷線

            df_final = pd.DataFrame(final_rows, columns=output_cols)
            
            # --- 瀏覽功能 ---
            st.markdown("### 👀 初步預覽 (全自動對位完成)")
            st.dataframe(df_final)

            # --- 下載功能 (Excel) ---
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='每日實戰班表')
            
            st.download_button(
                label="📥 下載實戰班表 Excel 檔案",
                data=buffer.getvalue(),
                file_name="每日實戰班表_鎖定解鎖版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("🎉 實戰班表已成功生成！請點擊上方按鈕下載。")

    except Exception as e:
        st.error(f"執行失敗，請確保護佐班表格式正確。錯誤訊息: {e}")
