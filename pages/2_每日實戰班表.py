import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日實戰班表生成器", layout="wide")
st.title("🖨️ 每日實戰班表產生器 (6人直向完美版)")
st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1: t_file = st.file_uploader("📂 1. 原始班表", type=['xlsx', 'csv'])
with col2: r_file = st.file_uploader("📂 2. 排班結果 (含HN/PT)", type=['xlsx', 'csv'])
with col3: a_file = st.file_uploader("📂 3. 護佐當月班表", type=['xlsx', 'csv'])

if t_file and r_file and a_file:
    try:
        # 強制讀取最原始的矩陣結構
        df_t = pd.read_csv(t_file, header=None) if 'csv' in t_file.name.lower() else pd.read_excel(t_file, header=None)
        df_r = pd.read_csv(r_file, header=None) if 'csv' in r_file.name.lower() else pd.read_excel(r_file, header=None)
        df_a = pd.read_csv(a_file, header=None) if 'csv' in a_file.name.lower() else pd.read_excel(a_file, header=None)

        # ----------------------------------------------------
        # 解析 1：原始班表 (抓取每日班別)
        # ----------------------------------------------------
        train_map = {}
        date_cols_t = {}
        for idx in range(len(df_t)):
            row_vals = [str(x).strip().replace('.0', '') for x in df_t.iloc[idx]]
            if '1' in row_vals and '2' in row_vals and '3' in row_vals:
                for col_idx, val in enumerate(df_t.iloc[idx]):
                    v_str = str(val).strip().replace('.0', '')
                    if v_str.isdigit() and 1 <= int(v_str) <= 31:
                        date_cols_t[v_str] = col_idx
                break
        
        for idx in range(3, len(df_t)): 
            name = str(df_t.iloc[idx, 3]).strip() 
            if name and name.lower() != 'nan' and name not in ['姓名', '星期']:
                train_map[name] = {d: str(df_t.iloc[idx, c_idx]).strip().upper() for d, c_idx in date_cols_t.items()}

        # ----------------------------------------------------
        # 解析 2：排班結果 (這是主力！包含 HN 與 PT 及上班區域)
        # ----------------------------------------------------
        roster_data = [] 
        date_cols_r = {}
        for col_idx, val in enumerate(df_r.iloc[0]):
            v_str = str(val).strip().replace('.0', '')
            if v_str.isdigit() and 1 <= int(v_str) <= 31:
                date_cols_r[v_str] = col_idx

        for idx in range(2, len(df_r)): 
            name = str(df_r.iloc[idx, 4]).strip() # 第 5 欄是姓名
            default_shift = str(df_r.iloc[idx, 3]).strip().upper() # 第 4 欄是當月班別 (HN/PT用)
            if name and name.lower() != 'nan' and not any(x in name for x in ['各班獨立', '總計', '人數', 'D', 'E', 'N']):
                zones = {d: str(df_r.iloc[idx, c_idx]).strip() for d, c_idx in date_cols_r.items()}
                roster_data.append({
                    'name': name,
                    'default_shift': default_shift,
                    'zones': zones
                })

        # ----------------------------------------------------
        # 解析 3：護佐班表
        # ----------------------------------------------------
        aide_map = {}
        date_cols_a = {}
        date_row_a = df_a.iloc[1] 
        for j in range(4, len(df_a.columns)): 
            v_str = str(date_row_a.iloc[j]).strip().replace('.0', '')
            if v_str.isdigit():
                date_cols_a[v_str] = j

        # 修正：從第 3 列 (index 2) 就開始抓，避免漏掉第一位護佐
        for idx in range(2, len(df_a)): 
            aide_name = str(df_a.iloc[idx, 3]).strip() 
            if aide_name and aide_name.lower() != 'nan' and aide_name != '姓名':
                aide_map[aide_name] = {d: str(df_a.iloc[idx, c_idx]).strip().upper() for d, c_idx in date_cols_a.items()}

        # ----------------------------------------------------
        # ⚡ 執行 6人一列 排版
        # ----------------------------------------------------
        sorted_days = sorted(list(date_cols_t.keys()), key=int)

        if st.button("🚀 開始轉出每日實戰班表 (6人直式版)"):
            output_cols = []
            for i in range(1, 7): # ✅ 已修改為 6 個人
                output_cols.extend([f'Name_{i}', f'Zone_{i}'])

            final_rows = []

            for day in sorted_days:
                # 緊湊排版，去掉多餘空白
                final_rows.append({k: (f"=== 6月 {day} 日 ===" if k == 'Name_1' else "") for k in output_cols})

                for shift_code, time_label in [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]:
                    final_rows.append({
                        'Name_1': f"【{shift_code}班】",
                        'Zone_1': time_label,
                        **{k: "" for k in output_cols[2:]}
                    })

                    shift_staff_list = []

                    # A. 撈出護理同仁與工讀生 (以 Roster 為基準尋找)
                    for r_item in roster_data:
                        name = r_item['name']
                        zone_val = r_item['zones'].get(day, "")
                        
                        if zone_val.upper() in ['X', 'NAN', 'NONE', 'OFF', '']:
                            continue # 今天沒排區域就是沒上班

                        # 如果是護理師，找 Training 的班別；如果是 HN/PT，找 Default Shift
                        today_shift = train_map.get(name, {}).get(day, r_item['default_shift'])

                        if shift_code in today_shift:
                            shift_staff_list.append((name, zone_val))

                    # B. 撈出護佐
                    for aide_name, dates in aide_map.items():
                        if shift_code in dates.get(day, ""):
                            shift_staff_list.append((aide_name, "護佐"))

                    # C. 6 人一列排版
                    for i in range(0, len(shift_staff_list), 6):
                        chunk = shift_staff_list[i:i+6]
                        name_row = {k: "" for k in output_cols}
                        zone_row = {k: "" for k in output_cols}

                        for idx, (s_name, s_zone) in enumerate(chunk):
                            col_idx = idx + 1
                            name_row[f'Name_{col_idx}'] = s_name
                            zone_row[f'Zone_{col_idx}'] = s_zone

                        final_rows.append(name_row)
                        final_rows.append(zone_row)

                    final_rows.append({k: "" for k in output_cols}) # 班別間隔

                final_rows.append({k: "------------------------" for k in output_cols}) # 每日中斷線

            df_final = pd.DataFrame(final_rows, columns=output_cols)
            
            st.markdown("### 👀 每日實戰班表初步預覽")
            st.dataframe(df_final)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # ✅ 關鍵：設定 header=False，這樣匯出 Excel 時就不會印出 Name_1, Zone_1 等醜醜的標題，直接從日期開始！
                df_final.to_excel(writer, index=False, header=False, sheet_name='每日實戰班表')
            
            st.download_button(
                label="📥 下載實戰班表 Excel 檔案",
                data=buffer.getvalue(),
                file_name="每日實戰班表_6人直向完美版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("🎉 實戰班表已成功生成！已修正區域遺失問題，改為 6 人一列，並移除多餘的標題列。")

    except Exception as e:
        st.error(f"發生錯誤，請截圖給工程師確認：{e}")
