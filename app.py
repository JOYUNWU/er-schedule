import streamlit as st
import pandas as pd
import numpy as np
import random
import io
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

# ==========================================
# 0. 頂層小幫手函數 (保持不變)
# ==========================================
def get_team_of(staff_name, df_daily):
    ts = df_daily[df_daily['姓名'] == staff_name]['組別'].values
    return str(ts[0]).strip().upper() if len(ts) > 0 else ""

def get_team_priority(staff_name, df_daily):
    t = get_team_of(staff_name, df_daily)
    if t == 'G': return 1
    elif t == 'F': return 2
    elif t == 'E': return 3
    elif t == 'D': return 4
    elif t == 'C': return 5
    elif t == 'B': return 6
    elif t == 'A': return 7
    else: return 8

def get_zone_count(z, m_counts, s_name):
    if z in ['A2', 'B2', 'C2']: return sum(m_counts[s_name].get(x, 0) for x in ['A2', 'B2', 'C2'])
    if z in ['MO', 'MO1', 'MO2']: return sum(m_counts[s_name].get(x, 0) for x in ['MO', 'MO1', 'MO2'])
    if z in ['R1', 'R3']: return sum(m_counts[s_name].get(x, 0) for x in ['R1', 'R3'])
    if z in ['GB', 'GC']: return sum(m_counts[s_name].get(x, 0) for x in ['GB', 'GC'])
    return m_counts[s_name].get(z, 0)

def get_macro(zone):
    if zone in ["MO", "MO1", "MO2"]: return "OBS"
    if zone in ["S", "S1", "S2"]: return "SURG"
    if zone in ["R", "R1", "R2", "R3"]: return "RESUS"
    if zone in ["A1", "A2", "B1", "B2", "C1", "C2"]: return "MED"
    if zone in ["GB", "GC", "GD"]: return "G"
    if zone in ["T", "T2"]: return "TRIAGE"
    if zone in ["P", "P2"]: return "PEDS"
    return zone

def is_severe(zone):
    return zone in ['R', 'S', 'C2']

# ==========================================
# 🌟 AI 權重計分系統
# ==================================================
def get_zone_score(zone, name, team, day_idx, df_result, date_columns, monthly_counts, work_blocks, current_day_macro_teams, ab_gb_gc_fulfilled):
    
    score = get_zone_count(zone, monthly_counts, name) * 10000 

    if team == 'A' and zone in ['R2', 'A2', 'S1']:
        score += 5000000 

    past_zones = []
    for i in [1, 2]: 
        if day_idx - i >= 0:
            pz = str(df_result.loc[df_result['姓名'] == name, date_columns[day_idx - i]].values[0]).strip()
            if pz not in ['OFF', '', 'X', 'NAN', 'L', 'L2', '職代', 'None']:
                past_zones.append(pz)

    past_macros = [get_macro(pz) for pz in past_zones]
    past_severe = any(is_severe(pz) for pz in past_zones)

    if get_macro(zone) in past_macros: score += 5000000
    if is_severe(zone) and past_severe: score += 5000000

    macro_z = get_macro(zone)
    if team in current_day_macro_teams.get(macro_z, set()):
        score += 20000 

    if zone in ['T', 'P']:
        prev_zone = past_zones[0] if len(past_zones) > 0 else ""
        if prev_zone not in ['T', 'P']: 
            days_to_off = work_blocks[name][day_idx]
            if days_to_off not in [2, 3]: 
                score += 5000000 

    if zone in ['GB', 'GC']:
        c = get_zone_count(zone, monthly_counts, name)
        if team in ['A', 'B'] and c == 0:
            score -= 100000 
        elif team in ['C', 'D', 'E'] and not ab_gb_gc_fulfilled:
            score += 20000 
        if team == 'D' and c >= 1: score += 5000000

    if zone == 'B1':
        if team == 'B': score -= 400
        elif team == 'A': score -= 300
        elif team == 'C': score -= 200
        
    elif zone == 'C1':
        if team == 'E': score -= 500
        elif team == 'D': score -= 400
        elif team == 'C': score -= 300
        elif team == 'B': score -= 200
        elif team == 'A': score -= 100
        
    elif zone == 'R':
        if team == 'A': score -= 400
        elif team == 'B': score -= 300
        elif team == 'C': score -= 200
        
    elif zone == 'S':
        if team == 'A': score -= 400
        elif team == 'B': score -= 300
        elif team == 'C': score -= 200
        
    elif zone in ['GB', 'GC']:
        if team == 'A': score -= 400
        elif team == 'B': score -= 300
        elif team == 'C': score -= 200

    # 💡 階梯式溢出權重 (A2+B2+C2)
    elif zone in ['A2', 'B2', 'C2']:
        if team == 'D': score -= 400
        elif team == 'C': score -= 300
        elif team == 'B': score -= 200
        elif team == 'A': score -= 100

    return score

# ==========================================
# 網頁 UI 初始化 (修正了顯示邏輯)
# ==========================================
st.set_page_config(page_title="急診自動排班系統", layout="wide")
st.title("急診護理人員自動初步排班系統")
st.markdown("---")

col1, col2 = st.columns(2)
with col1: training_file = st.file_uploader("📂 1. 上傳班表", type=['xlsx', 'csv'])
with col2: template_file = st.file_uploader("📂 2. 上傳空白檔", type=['xlsx', 'csv'])

# ... (檔案讀取邏輯保持原樣，中間排班邏輯省略以節省空間)

# 為了節省空間，請直接將上述 `if st.button("🚀 開始自動排班運算")` 的內容接續在下面...
# 確保排班後的 df_result 生成後，執行以下的修正顯示：

if st.button("🚀 開始自動排班運算", disabled=not data_ready):
    with st.spinner("🚀 生成中，已修正顯示 BUG..."):
        try:
            # ... (中間排班邏輯保持不變) ...
            
            # (排班邏輯執行結束後，df_result 已經產生)
            df_result = df_result.fillna("")

            # 💡 修正顯示錯誤：先取 head(10)，再轉 style
            st.success("🎉 排班完成！")
            st.dataframe(df_result.head(10)) 

            # (後續 Excel 寫入公式的邏輯保持不變)
            # ...
