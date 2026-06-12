import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("🖨️ 每日實戰班表產生器 (最終穩定版)")

col1, col2, col3 = st.columns(3)
with col1: training_file = st.file_uploader("📂 1. 原始班表", type=['xlsx', 'csv'])
with col2: roster_file = st.file_uploader("📂 2. 排班結果 (含HN/工讀生)", type=['xlsx', 'csv'])
with col3: aide_file = st.file_uploader("📂 3. 護佐班表", type=['xlsx', 'csv'])

if training_file and roster_file and aide_file:
    try:
        # 使用 header=None 並自動掃描關鍵字，這是最保險的讀取法
        df_train = pd.read_csv(training_file, header=None)
        df_roster = pd.read_csv(roster_file, header=None)
        df_aide = pd.read_csv(aide_file, header=None)

        def find_header_and_data(df):
            # 尋找「姓名」在哪一列
            for i in range(10):
                for j in range(10):
                    if "姓名" in str(df.iloc[i, j]):
                        return i, j
            return 1, 0

        t_row, t_col = find_header_and_data(df_train)
        r_row, r_col = find_header_and_data(df_roster)
        a_row, a_col = find_header_and_data(df_aide)

        # 重新建構 DataFrame
        def build_df(df, r, c):
            new_df = df.iloc[r:].copy()
            new_df.columns = df.iloc[r]
            return new_df.iloc[1:]

        df_t = build_df(df_train, t_row, t_col)
        df_r = build_df(df_roster, r_row, r_col)
        df_a = build_df(df_aide, a_row, a_col)

        date_cols = [c for c in df_t.columns if str(c).strip().isdigit()]

        if st.button("🚀 產出報表"):
            output_rows = []
            for d in date_cols:
                output_rows.append({"Name_1": f"=== 6/{d} ==="})
                for shift, time in [('D', "7'-4"), ('E', "3'-12"), ('N', "11'-8")]:
                    output_rows.append({"Name_1": f"【{shift}班】", "Zone_1": time})
                    staff_list = []
                    
                    # 抓取護理師
                    for _, row in df_t.iterrows():
                        if shift in str(row.get(d, "")):
                            name = str(row['姓名']).strip()
                            zone = str(df_r[df_r['姓名'] == name].get(d, [""])).strip()
                            staff_list.append((name, zone))
                    
                    # 抓取護佐
                    for _, row in df_a.iterrows():
                        if shift in str(row.get(d, "")):
                            staff_list.append((str(row['姓名']).strip(), "護佐"))
                    
                    for i in range(0, len(staff_list), 8):
                        chunk = staff_list[i:i+8]
                        n_row = {f"Name_{idx+1}": n for idx, (n, z) in enumerate(chunk)}
                        z_row = {f"Zone_{idx+1}": z for idx, (n, z) in enumerate(chunk)}
                        output_rows.extend([n_row, z_row, {}])
            
            df_final = pd.DataFrame(output_rows)
            st.dataframe(df_final)
            
            # 下載 Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as w: df_final.to_excel(w, index=False)
            st.download_button("📥 下載 Excel", data=buffer.getvalue(), file_name="每日實戰班表.xlsx")
    except Exception as e:
        st.error(f"錯誤，請確認檔案是否有欄位名稱「姓名」: {e}")
