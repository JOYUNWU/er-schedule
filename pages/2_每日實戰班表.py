import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="每日實戰班表生成器", layout="wide")
st.title("🖨️ 每日實戰班表產生器 (AI自動對位完美版)")
st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1: t_file = st.file_uploader("📂 1. 原始班表", type=['xlsx', 'csv'])
with col2: r_file = st.file_uploader("📂 2. 排班結果 (含HN/PT)", type=['xlsx', 'csv'])
with col3: a_file = st.file_uploader("📂 3. 護佐當月班表", type=['xlsx', 'csv'])

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
            wd = str(df.iloc[date_row_idx
