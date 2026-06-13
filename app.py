import streamlit as st
import pandas as pd
import numpy as np
import random
import io
from openpyxl.styles import PatternFill # 🌟 匯入顏色標記套件

# ==========================================
# 0. 頂層小幫手函數
# ==========================================
def get_team_of(staff_name, df_daily):
    ts = df_daily[df_daily['姓名'] == staff_name]['組別'].values
    return str(ts[0]).strip().upper() if len(ts) > 0 else ""

def get_team_priority(staff_name, df_daily):
    t = get_team_of(staff_name, df_daily)
    if t == 'F': return 1
    elif t == 'E': return 2
    elif t == 'D': return 3
    elif t == 'C': return 4
    elif t == 'B': return 5
    elif t == 'A': return 6
    else: return 7

def get_zone_count(z, m_counts, s_name):
    if z in ['A2', 'B2', 'C2']: return sum(m_counts[s_name].get(x, 0) for x in ['A2', 'B2', 'C2'])
    if z in ['MO', 'MO1', 'MO2']: return sum(m_counts[s_name].get(x, 0) for x in ['MO', 'MO1', 'MO2'])
    if z in ['GC', 'GB', 'T2']: return sum(m_counts[s_name].get(x, 0) for x in ['GC', 'GB', 'T2'])
    return m_counts[s_name].get(z, 0)

def get_macro(zone):
    if zone in ["MO", "MO1", "MO2"]: return "OBS"
    if zone in ["S", "S1", "S2"]: return "SURG"
    if zone in ["R", "R1", "R2", "R3"]: return "RESUS"
    if zone in ["A2", "B1", "B2", "C1", "C2", "A1"]: return "MED"
    if zone in ["GB", "GC", "GD"]: return "G"
    if zone in ["T", "T2"]: return "TRIAGE"
    if zone in ["P", "P2"]: return "PEDS"
    return zone

def is_severe(zone):
    return zone in ['R', 'S', 'C2']

# 💡 智慧次數管控引擎
def apply_limit(c, base, max_limit):
    if c < base: return -8000
    elif c >= max_limit: return 50000
    else: return c * 100 

# 🌟 終極 AI 權重計分系統
def get_zone_score(zone, name, team, day_idx, df_result, date_columns, monthly_counts, work_blocks):
    score = get_zone_count(zone, monthly_counts, name) * 10 

    past_zones = []
    for i in [1, 2]: 
        if day_idx - i >= 0:
            pz = str(df_result.loc[df_result['姓名'] == name, date_columns[day_idx - i]].values[0]).strip()
            if pz not in ['OFF', '', 'X', 'NAN', 'L', 'L2', '職代', 'None']:
                past_zones.append(pz)

    past_macros = [get_macro(pz) for pz in past_zones]
    past_severe = any(is_severe(pz) for pz in past_zones)

    if get_macro(zone) in past_macros: score += 20000
    if is_severe(zone) and past_severe: score += 20000

    # ==================================================
    # 🌟 各組精確上下限鎖定
    # ==================================================
    if team in ['A', 'B']:
        if zone in ['MO', 'MO1', 'MO2']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['MO', 'MO1', 'MO2'])
            score += apply_limit(c, 2, 3)
        elif zone in ['GB', 'GC', 'T2']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['GB', 'GC', 'T2'])
            score += apply_limit(c, 2, 3)
        elif zone == 'R':
            c = monthly_counts[name].get('R', 0)
            score += apply_limit(c, 2, 3)
        elif zone == 'B1':
            c = monthly_counts[name].get('B1', 0)
            score += apply_limit(c, 2, 3)
        elif zone in ['R1', 'R3']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['R1', 'R3'])
            score += apply_limit(c, 2, 3)
        elif zone in ['A2', 'B2', 'C2']: 
            c = sum(monthly_counts[name].get(x, 0) for x in ['A2', 'B2', 'C2'])
            score += apply_limit(c, 3, 4)
        elif zone == 'S':
            c = monthly_counts[name].get('S', 0)
            score += apply_limit(c, 2, 3)
        elif zone == 'S2':
            c = monthly_counts[name].get('S2', 0)
            score += apply_limit(c, 2, 3)

    elif team == 'C':
        if zone in ['MO', 'MO1', 'MO2']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['MO', 'MO1', 'MO2'])
            score += apply_limit(c, 2, 3)
        elif zone in ['GB', 'GC']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['GB', 'GC'])
            score += apply_limit(c, 1, 2)
        elif zone in ['R1', 'R3']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['R1', 'R3'])
            score += apply_limit(c, 2, 3)
        elif zone in ['A2', 'B2', 'C2']: 
            c = sum(monthly_counts[name].get(x, 0) for x in ['A2', 'B2', 'C2'])
            score += apply_limit(c, 3, 4)
        elif zone == 'S2':
            c = monthly_counts[name].get('S2', 0)
            score += apply_limit(c, 1, 2)
        elif zone == 'C1':
            c = monthly_counts[name].get('C1', 0)
            score += apply_limit(c, 2, 3)
        elif zone == 'A1':
            c = monthly_counts[name].get('A1', 0)
            score += apply_limit(c, 1, 2)
        elif zone == 'S1':
            c = monthly_counts[name].get('S1', 0)
            score += apply_limit(c, 1, 2)

    elif team == 'D':
        if zone in ['GB', 'GC']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['GB', 'GC'])
            score += apply_limit(c, 1, 1) 
        elif zone in ['A2', 'B2']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['A2', 'B2'])
            score += apply_limit(c, 3, 4)
        elif zone == 'R1':
            c = monthly_counts[name].get('R1', 0)
            score += apply_limit(c, 2, 3)
        elif zone in ['MO', 'MO1', 'MO2']:
            c = sum(monthly_counts[name].get(x, 0) for x in ['MO', 'MO1', 'MO2'])
            score += apply_limit(c, 2, 3)
        elif zone == 'S1':
            c = monthly_counts[name].get('S1', 0)
            score += apply_limit(c, 2, 3)

    # ==================================================
    # 🌟 A1, S1, R2 絕對保留給 E 與 F
    if zone in ['A1', 'S1', 'R2']: 
        if team == 'A': score += 50000 # A組絕對封殺
        elif team not in ['E', 'F']: score += 8000 # B, C, D組不鼓勵
        elif team in ['E', 'F']: score -= 8000 # 強力拉攏 E, F

    # 🌟 T的壓線防護與A組任務解鎖
    if zone == 'T':
        is_continuing_T = (day_idx > 0 and df_result.loc[df_result['姓名'] == name, date_columns[day_idx - 1]].values[0] == 'T')
        if not is_continuing_T:
            if work_blocks[name][day_idx] < 2:
                score += 50000 # 嚴格封殺：連續上班天數不夠，不准開始T班
        if team == 'A' and monthly_counts[name].get('T', 0) == 0:
            score -= 5000 # 優先讓A組解鎖T

    # 🌟 P的壓線防護
    if zone == 'P':
        is_continuing_P = (day_idx > 0 and df_result.loc[df_result['姓名'] == name, date_columns[day_idx - 1]].values[0] == 'P')
        if not is_continuing_P and work_blocks[name][day_idx] < 2:
            score += 50000

    # 通用天花板保護
    explicit_zones = ["MO","MO1","MO2","GB","GC","T2","R","B1","R1","R3","A2","B2","C2","S","S2","C1","A1","S1","P","R2"]
    if zone not in explicit_zones and monthly_counts.get(name, {}).get(zone, 0) >= 3:
        score += 8000

    return score

# ==========================================
# 網頁 UI 初始化
# ==========================================
st.set_page_config(page_title="急診自動排班系統", layout="wide")
st.title("🏥 急診護理人員自動排班系統 (次數鎖定與螢光標示版)")
st.markdown("---")

col1, col2 = st.columns(2)
with col1: training_file = st.file_uploader("📂 1. 上傳班表", type=['xlsx', 'csv'])
with col2: template_file = st.file_uploader("📂 2. 上傳空白檔", type=['xlsx', 'csv'])

all_staff, df_shift, df_template, data_ready, date_columns = [], None, None, False, []

if training_file and template_file:
    try:
        df_shift = pd.read_csv(training_file) if 'csv' in training_file.name.lower() else pd.read_excel(training_file)
        df_template = pd.read_csv(template_file) if 'csv' in template_file.name.lower() else pd.read_excel(template_file)
            
        shift_cols = list(df_shift.columns)
        shift_cols[0], shift_cols[1], shift_cols[2], shift_cols[3] = '編號', '組別', '性別', '姓名'
        date_length = len(shift_cols) - 4
        shift_cols[4:] = [str(i) for i in range(1, date_length + 1)]
        df_shift.columns = shift_cols
        
        template_cols = list(df_template.columns)
        template_cols[0], template_cols[1], template_cols[2], template_cols[3] = '編號', '組別', '性別', '姓名'
        template_cols[4:4+date_length] = [str(i) for i in range(1, date_length + 1)]
        df_template.columns = template_cols
        
        df_shift, df_template = df_shift.dropna(subset=['姓名']), df_template.dropna(subset=['姓名'])
        df_shift['姓名'], df_template['姓名'] = df_shift['姓名'].astype(str).str.strip(), df_template['姓名'].astype(str).str.strip()
        all_staff, date_columns, data_ready = df_shift['姓名'].unique().tolist(), df_shift.columns[4:], True
        st.success("✅ 檔案載入成功！")
    except Exception as e:
        st.error(f"檔案讀取失敗：{e}")

# 左側設定
st.sidebar.header("⚙️ 本月特殊排班規則設定")
train_s2 = st.sidebar.multiselect("S2 訓練名單", options=all_staff if data_ready else [])
train_b1_r = st.sidebar.multiselect("B1/R/C2/S 訓練名單", options=all_staff if data_ready else [])
train_t = st.sidebar.multiselect("檢傷(T) 訓練名單", options=all_staff if data_ready else [])

leader_options = ["請點選"] + all_staff if data_ready else ["請點選"]

st.sidebar.subheader("👑 各班組長順位")
d_l_1 = st.sidebar.selectbox("D班 第一順位", leader_options)
d_l_2 = st.sidebar.selectbox("D班 第二順位", leader_options)
d_l_3 = st.sidebar.selectbox("D班 第三順位", leader_options)
e_l_1 = st.sidebar.selectbox("E班 第一順位", leader_options)
e_l_2 = st.sidebar.selectbox("E班 第二順位", leader_options)
e_l_3 = st.sidebar.selectbox("E班 第三順位", leader_options)
n_l_1 = st.sidebar.selectbox("N班 第一順位", leader_options)
n_l_2 = st.sidebar.selectbox("N班 第二順位", leader_options)
n_l_3 = st.sidebar.selectbox("N班 第三順位", leader_options)

st.markdown("---")
if st.button("🚀 開始自動排班運算", disabled=not data_ready):
    with st.spinner("🧠 嚴格執行各組次數軟保底與硬天花板鎖定中..."):
        try:
            df_result = df_template.copy()
            monthly_counts = {name: {} for name in all_staff}
            progress_bar = st.progress(0)
            
            work_blocks = {name: [0]*len(date_columns) for name in all_staff}
            for name in all_staff:
                for day_idx in range(len(date_columns)):
                    count = 0
                    for future_idx in range(day_idx, len(date_columns)):
                        col = date_columns[future_idx]
                        val = str(df_shift.loc[df_shift['姓名'] == name, col].values[0]).strip().upper()
                        if val == 'OFF': break
                        count += 1
                    work_blocks[name][day_idx] = count

            team_allowed_zones = {
                'A': ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2", "GC", "GB", "GD"],
                'B': ["A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "GC", "GB", "GD"],
                'C': ["A1", "B1", "C1", "A2", "B2", "C2", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "GC", "GB", "GD"],
                'D': ["A1", "C1", "A2", "B2", "R1", "R2", "P", "MO", "MO1", "MO2", "S1", "GB", "GC"],
                'E': ["A1", "A2", "R2", "MO", "MO1", "P", "S1"],
                'F': ["A1", "R2", "MO", "MO1", "S1"]
            }
            
            for day_idx, date_col in enumerate(date_columns):
                todays_shifts = df_shift[['姓名', '組別', '性別', date_col]].copy()
                todays_shifts.columns = ['姓名', '組別', '性別', '班別']
                todays_template = df_template[['姓名', date_col]].copy()
                todays_template.columns = ['姓名', '預設區域']
                daily_data = pd.merge(todays_shifts, todays_template, on='姓名')
                
                for shift_type in ['D', 'E', 'N']:
                    shift_staff = daily_data[daily_data['班別'].astype(str).str.upper() == shift_type].copy()
                    if len(shift_staff) == 0: continue
                    
                    unassigned_staff = shift_staff['姓名'].tolist()
                    assignments = {}
                    pre_assigned_zones = []
                    
                    for _, row in shift_staff.iterrows():
                        name, preset = row['姓名'], str(row['預設區域']).strip()
                        if preset.upper() not in ['X', 'NAN', 'NONE', ''] and name in unassigned_staff:
                            assignments[name] = preset
                            unassigned_staff.remove(name)
                            pre_assigned_zones.append(preset)
                    
                    if "L" not in assignments.values():
                        l_chain = [d_l_1, d_l_2, d_l_3] if shift_type == 'D' else [e_l_1, e_l_2, e_l_3] if shift_type == 'E' else [n_l_1, n_l_2, n_l_3]
                        for l_cand in l_chain:
                            if l_cand in unassigned_staff:
                                assignments[l_cand] = "L"
                                unassigned_staff.remove(l_cand)
                                break

                    target_zone_count = len(unassigned_staff)
                    full_priority_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "P", "MO", "MO1", "MO2", "S1", "S", "S2", "R3", "T2", "GB", "GC", "GD"]
                    available_zones = []
                    
                    for z in full_priority_zones:
                        if len(available_zones) >= target_zone_count: break
                        if z not in pre_assigned_zones:
                            available_zones.append(z)

                    filler_zones = ["MO", "A2", "B2", "C1", "C2", "R1", "S1", "MO1", "B1"]
                    f_idx = 0
                    while len(available_zones) < target_zone_count:
                        available_zones.append(filler_zones[f_idx % len(filler_zones)])
                        f_idx += 1
                    
                    # 🌟 N連啟倫 霸王鎖定 (要在排班初期優先鎖定)
                    if "N連啟倫" in unassigned_staff:
                        mo_zones = [z for z in ["MO", "MO1"] if z in available_zones]
                        if mo_zones:
                            chosen = min(mo_zones, key=lambda z: monthly_counts.get("N連啟倫", {}).get(z, 0))
                            assignments["N連啟倫"] = chosen
                            unassigned_staff.remove("N連啟倫")
                            available_zones.remove(chosen)

                    # 連續排班機制
                    continuous_reqs = []
                    for name in list(unassigned_staff):
                        if day_idx > 0:
                            prev_shift = str(df_shift.loc[df_shift['姓名'] == name, date_columns[day_idx - 1]].values[0]).strip().upper()
                            if prev_shift != 'OFF':
                                prev_zone = str(df_result.loc[df_result['姓名'] == name, date_columns[day_idx - 1]].values[0]).strip()
                                if prev_zone not in ['OFF', 'L', 'X', 'NAN', '']:
                                    is_train = (name in train_s2 and prev_zone == "S2") or (name in train_t and prev_zone == "T") or (name in train_b1_r and prev_zone in ["B1", "R", "S", "C2"])
                                    is_TP = prev_zone in ["T", "P"]
                                    is_team_F = (get_team_of(name, shift_staff) == 'F')
                                    if is_train or is_TP or is_team_F:
                                        continuous_reqs.append({'name': name, 'zone': prev_zone, 'is_train': is_train})
                    
                    continuous_reqs.sort(key=lambda x: not x['is_train']) 
                    for req in continuous_reqs:
                        name, z = req['name'], req['zone']
                        if name in unassigned_staff and z in available_zones:
                            assignments[name] = z
                            unassigned_staff.remove(name)
                            available_zones.remove(z)

                    # 新訓防撞分配
                    for t_list, t_zone in [(train_s2, "S2"), (train_t, "T")]:
                        cands = [n for n in list(unassigned_staff) if n in t_list]
                        if cands and t_zone in available_zones:
                            cands.sort(key=lambda n: monthly_counts[n].get(t_zone, 0))
                            assignments[cands[0]] = t_zone
                            unassigned_staff.remove(cands[0])
                            available_zones.remove(t_zone)

                    brcs_cands = [n for n in list(unassigned_staff) if n in train_b1_r]
                    for name in brcs_cands:
                        v_zones = [z for z in ["B1", "R", "S", "C2"] if z in available_zones]
                        if v_zones:
                            v_zones.sort(key=lambda z: monthly_counts[name].get(z, 0))
                            assignments[name] = v_zones[0]
                            unassigned_staff.remove(name)
                            available_zones.remove(v_zones[0])
                    
                    # F組/E組優先配區
                    for name in [n for n in list(unassigned_staff) if get_team_of(n, shift_staff) == 'F']:
                        a_zones = [z for z in ["A1", "R2", "MO", "MO1", "S1"] if z in available_zones]
                        if a_zones:
                            z_scores = [(z, get_zone_score(z, name, 'F', day_idx, df_result, date_columns, monthly_counts, work_blocks)) for z in a_zones]
                            z_scores.sort(key=lambda x: x[1])
                            best_z, best_s = z_scores[0]
                            if best_s < 20000: 
                                assignments[name] = best_z
                                unassigned_staff.remove(name)
                                available_zones.remove(best_z)

                    for name in [n for n in list(unassigned_staff) if get_team_of(n, shift_staff) == 'E']:
                        if "A2" in available_zones:
                            if get_zone_score("A2", name, 'E', day_idx, df_result, date_columns, monthly_counts, work_blocks) < 20000:
                                assignments[name] = "A2"
                                unassigned_staff.remove(name)
                                available_zones.remove("A2")

                    # AI 計分歸建分配
                    random.shuffle(unassigned_staff)
                    unassigned_staff.sort(key=lambda x: get_team_priority(x, shift_staff))

                    for name in list(unassigned_staff):
                        gender = shift_staff[shift_staff['姓名'] == name]['性別'].values[0] if len(shift_staff[shift_staff['姓名'] == name]) > 0 else ""
                        team = get_team_of(name, shift_staff)
                        team_allowed = team_allowed_zones.get(team, available_zones)

                        valid_cands = [z for z in available_zones if z in team_allowed and not (z == "S2" and str(gender).upper() == "M")]
                        if not valid_cands: valid_cands = [z for z in available_zones if not (z == "S2" and str(gender).upper() == "M")]
                        if not valid_cands: valid_cands = available_zones.copy()

                        z_scores = [(z, get_zone_score(z, name, team, day_idx, df_result, date_columns, monthly_counts, work_blocks)) for z in valid_cands]
                        z_scores.sort(key=lambda x: x[1])
                        
                        chosen_zone = z_scores[0][0]
                        assignments[name] = chosen_zone
                        unassigned_staff.remove(name)
                        if chosen_zone in available_zones: available_zones.remove(chosen_zone)

                    for name, assigned_zone in assignments.items():
                        df_result.loc[df_result['姓名'] == name, date_col] = assigned_zone
                        monthly_counts[name][assigned_zone] = monthly_counts[name].get(assigned_zone, 0) + 1

                progress_bar.progress((day_idx + 1) / len(date_columns))

            name_col_index = df_result.columns.get_loc('姓名')
            majority_shift_dict = {name: max(set([x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']]), key=[x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']].count) if [x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']] else "" for name in all_staff}
            df_result.insert(name_col_index, '當月班別', df_result['姓名'].map(majority_shift_dict))
            
            # 月底總計生成
            zone_count_order = ["L", "L2", "T", "T2", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "S1", "S2", "S", "P", "P2", "MO", "MO1", "MO2", "GB", "GC", "GD", "職代"]
            for count_zone in zone_count_order:
                df_result[count_zone] = df_result[date_columns].apply(lambda row: (row == count_zone).sum(), axis=1)
            
            df_result['A2+B2+C2計'] = df_result[['A2', 'B2', 'C2']].sum(axis=1)
            df_result['MO系計'] = df_result[['MO', 'MO1', 'MO2']].sum(axis=1)
            df_result['GC+GB+T2計'] = df_result[['GC', 'GB', 'T2']].sum(axis=1)

            df_result = df_result.fillna("")

            st.success("🎉 排班完成！已啟動次數封閉鎖定，且當 B1 或 R 次數超過 3 次時，下載的 Excel 將自動以「螢光黃」網底標記該儲存格。")
            st.dataframe(df_result.head(10))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='排班結果')
                
                # 🌟 實作：匯出 Excel 時，幫 B1 或 R > 3 的儲存格加上螢光黃網底
                workbook = writer.book
                worksheet = writer.sheets['排班結果']
                b1_idx = df_result.columns.get_loc('B1') + 1
                r_idx = df_result.columns.get_loc('R') + 1
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                for row in range(2, len(df_result) + 2): # 跳過標題列
                    for col_idx in [b1_idx, r_idx]:
                        val = worksheet.cell(row=row, column=col_idx).value
                        if isinstance(val, (int, float)) and val > 3:
                            worksheet.cell(row=row, column=col_idx).fill = yellow_fill

                df_shift.to_excel(writer, index=False, sheet_name='原始班表')

            st.download_button("📥 下載最終排班表 (Excel)", data=output.getvalue(), file_name="排班結果_螢光鎖定版.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"發生內部錯誤：{e}")
