import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日實戰班表生成器", layout="wide")
st.title("🖨️ 第二階段：產出每日實戰班表 (直向 A4)")
st.markdown("---")

st.info("💡 運作邏輯：由「原始班表」決定誰上班、「排班結果」決定每人守區、「護佐班表」由指定 D 欄與第 2 列日期精準定位。")

col1, col2, col3 = st.columns(3)
with col1: t_file = st.file_uploader("📂 1. 原始班表 (Training)", type=['xlsx', 'csv'])
with col2: r_file = st.file_uploader("📂 2. 排班結果 (含HN/工讀生)", type=['xlsx', 'csv'])
with col3: a_file = st.file_uploader("📂 3. 護佐當月班表", type=['xlsx', 'csv'])

if t_file and r_file and a_file:
    try:
        # 強制讀取最原始的矩陣結構，不預設表頭
        df_t = pd.read_csv(t_file, header=None) if 'csv' in t_file.name.lower() else pd.read_excel(t_file, header=None)
        df_r = pd.read_csv(r_file, header=None) if 'csv' in r_file.name.lower() else pd.read_excel(r_file, header=None)
        df_a = pd.read_csv(a_file, header=None) if 'csv' in a_file.name.lower() else pd.read_excel(a_file, header=None)

        # ----------------------------------------------------
        # 解析 1：原始班表 (Training) - 姓名在 index 3, 日期在數字列
        # ----------------------------------------------------
        train_map = {}
        date_cols_t = {}
        # 尋找含有日期的那一列
        for idx in range(len(df_t)):
            row_vals = [str(x).strip().replace('.0', '') for x in df_t.iloc[idx]]
            if '1' in row_vals and '2' in row_vals and '3' in row_vals:
                for col_idx, val in enumerate(df_t.iloc[idx]):
                    v_str = str(val).strip().replace('.0', '')
                    if v_str.isdigit() and 1 <= int(v_str) <= 31:
                        date_cols_t[v_str] = col_idx
                break
        
        # 抓取護理人員每日班別
        for idx in range(3, len(df_t)): # 從第 4 行開始是人員
            name = str(df_t.iloc[idx, 3]).strip() # 姓名固定在第 4 欄 (index 3)
            if name and name.lower() != 'nan' and name != '姓名' and name != '星期':
                train_map[name] = {d: str(df_t.iloc[idx, c_idx]).strip().upper() for d, c_idx in date_cols_t.items()}

        # ----------------------------------------------------
        # 解析 2：排班結果 (Roster) - 姓名在 index 4, 日期在第 1 行
        # ----------------------------------------------------
        roster_map = {}
        date_cols_r = {}
        for col_idx, val in enumerate(df_r.iloc[0]):
            v_str = str(val).strip().replace('.0', '')
            if v_str.isdigit() and 1 <= int(v_str) <= 31:
                date_cols_r[v_str] = col_idx

        for idx in range(2, len(df_r)): # 從人員開始
            name = str(df_r.iloc[idx, 4]).strip() # 姓名固定在第 5 欄 (index 4)
            if name and name.lower() != 'nan' and not any(x in name for x in ['各班獨立', '總計', '人數', 'D', 'E', 'N']):
                roster_map[name] = {d: str(df_r.iloc[idx, c_idx]).strip() for d, c_idx in date_cols_r.items()}

        # ----------------------------------------------------
        # 解析 3：護佐班表 (Aide) - 姓名在 index 3 (D欄), 日期在 index 1 (第2列)
        # ----------------------------------------------------
        aide_map = {}
        date_cols_a = {}
        date_row_a = df_a.iloc[1] # 第 2 列 (index 1)
        for j in range(4, len(df_a.columns)): # 從 E 欄 (index 4) 開始是日期
            v_str = str(date_row_a.iloc[j]).strip().replace('.0', '')
            if v_str.isdigit():
                date_cols_a[v_str] = j

        for idx in range(3, len(df_a)): # 從第 4 行開始是護佐名單
            aide_name = str(df_a.iloc[idx, 3]).strip() # D 欄 (index 3) 是姓名
            if aide_name and aide_name.lower() != 'nan' and aide_name != '姓名':
                aide_map[aide_name] = {d: str(df_a.iloc[idx, c_idx]).strip().upper() for d, c_idx in date_cols_a.items()}

        # ----------------------------------------------------
        # ⚡ 開始交叉對位與直式排版
        # ----------------------------------------------------
        sorted_days = sorted(list(date_cols_t.keys()), key=int)

        if st.button("🚀 開始轉出每日實戰班表 (直向 A4)"):
            output_cols = []
            for i in range(1, 9):
                output_cols.extend([f'Name_{i}', f'Zone_{i}'])

            final_rows = []

            for day in sorted_days:
                # 每日大標題
                final_rows.append({k: (f"=== 6月 {day} 日 ===" if k == 'Name_1' else "") for k in output_cols})
                final_rows.append({k: "" for k in output_cols})

                for shift_code, time_label in [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]:
                    final_rows.append({
                        'Name_1': f"【{shift_code}班】",
                        'Zone_1': time_label,
                        **{k: "" for k in output_cols[2:]}
                    })

                    shift_staff_list = []

                    # A. 撈出護理同仁
                    for name, dates in train_map.items():
                        today_shift = dates.get(day, "")
                        if shift_code in today_shift: # 包含 D, E, N
                            zone_val = roster_map.get(name, {}).get(day, "")
                            if zone_val.upper() in ['X', 'NAN', 'NONE', 'OFF', '']:
                                zone_val = ""
                            shift_staff_list.append((name, zone_val))

                    # B. 撈出護佐同仁
                    for aide_name, dates in aide_map.items():
                        today_aide_shift = dates.get(day, "")
                        if shift_code in today_aide_shift: # 包含 D, D2, E, N
                            shift_staff_list.append((aide_name, "護佐"))

                    # C. 每 8 人分切成一列，名字在上方，守區在下方
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
                        final_rows.append({k: "" for k in output_cols}) # 人員下方留空一行

                    final_rows.append({k: "" for k in output_cols}) # 班別間隔

                final_rows.append({k: "------------------------" for k in output_cols}) # 每日中斷線

            df_final = pd.DataFrame(final_rows, columns=output_cols)
            
            # 呈現預覽
            st.markdown("### 👀 每日實戰班表初步預覽")
            st.dataframe(df_final)

            # 下載 Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='每日實戰班表')
            
            st.download_button(
                label="📥 下載實戰班表 Excel 檔案",
                data=buffer.getvalue(),
                file_name="每日實戰班表_直向A4版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("🎉 實戰班表已成功生成！請點擊上方按鈕下載。")

    except Exception as e:
        st.error(f"處理失敗，請確認檔案格式是否與本月相同。錯誤訊息: {e}")
