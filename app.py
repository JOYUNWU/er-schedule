import streamlit as st
import pandas as pd
import numpy as np
import random
import io

# ==========================================
# 0. 頂層小幫手函數
# ==========================================
def get_team_of(staff_name, df_daily):
    ts = df_daily[df_daily['姓名'] == staff_name]['組別'].values
    return str(ts[0]).strip().upper() if len(ts) > 0 else ""

def get_team_priority(staff_name, df_daily):
    t = get_team_of(staff_name, df_daily)
    # 優先順序：限制越多的組別越先挑
    if t == 'G': return 1
    elif t == 'F': return 2
    elif t == 'E': return 3
    elif t == 'D': return 4
    elif t == 'C': return 5
    elif t == 'B': return 6
    elif t == 'A': return 7
    else: return 8

# 🌟 合算區域平均化
def get_zone_count(z, m_counts, s_name):
    if z in ['A1', 'B2', 'C2']: return sum(m_counts[s_name].get(x, 0) for x in ['A1', 'B2', 'C2'])
    if z in ['MO', 'MO1', 'MO2']: return sum(m_counts[s_name].get(x, 0) for x in ['MO', 'MO1', 'MO2'])
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
# 🌟 AI 權重計分系統 (加入強力均分與動態解鎖)
# ==================================================
def get_zone_score(zone, name, team, day_idx, df_result, date_columns, monthly_counts, work_blocks, current_day_macro_teams, ab_needs_gbgc):
    
    # 🌟 絕對均分條款：針對 B1, C1, R 施加 10 倍的均分強迫症權重
    if zone in ['B1', 'C1', 'R']:
        score = get_zone_count(zone, monthly_counts, name) * 1000 
    else:
        score = get_zone_count(zone, monthly_counts, name) * 100 

    past_zones = []
    for i in [1, 2]: 
        if day_idx - i >= 0:
            pz = str(df_result.loc[df_result['姓名'] == name, date_columns[day_idx - i]].values[0]).strip()
            if pz not in ['OFF', '', 'X', 'NAN', 'L', 'L2', '職代', 'None']:
                past_zones.append(pz)

    past_macros = [get_macro(pz) for pz in past_zones]
    past_severe = any(is_severe(pz) for pz in past_zones)

    # 1. 嚴格跳兩天大區懲罰 (絕對防撞)
    if get_macro(zone) in past_macros: score += 50000
    # 2. 重症錯開霸王條款 (絕對防撞)
    if is_severe(zone) and past_severe: score += 50000

    # 3. 區域涵蓋不同組別 (資深帶資淺混排)
    macro_z = get_macro(zone)
    if team in current_day_macro_teams.get(macro_z, set()):
        score += 2000 # 同大區已經有同組的人，加罰分數讓給其他組

    # 4. T與P的壓線進場機制 (未來預判)
    if zone in ['T', 'P']:
        prev_zone = past_zones[0] if len(past_zones) > 0 else ""
        if prev_zone not in ['T', 'P']: # 若是「第一天」要進場
            days_to_off = work_blocks[name][day_idx]
            if days_to_off not in [2, 3]: # 只有剩2天或3天時准許進場
                score += 50000 

    # 5. 🌟 GB/GC 動態公平競爭解鎖
    if zone in ['GB', 'GC']:
        c = sum(monthly_counts[name].get(x, 0) for x in ['GB', 'GC'])
        if team in ['A', 'B'] and c == 0:
            score -= 10000 # 極強拉力，確保 A/B 優先解鎖
        elif team in ['C', 'D', 'E']:
            if ab_needs_gbgc:
                score += 5000 # A/B 還沒全員解鎖前，CDE 乖乖讓路
            # 如果 ab_needs_gbgc 變為 False (代表 A/B 已經全員上過)，此處就不加懲罰，CDE 加入公平競爭！

    # 6. 避讓條款
    if zone in ['A1', 'S1', 'R2']: 
        if team == 'A': score += 20000 
        elif team in ['B', 'C', 'D']: score += 8000 

    return score

# ==========================================
# 網頁 UI 初始化
# ==========================================
st.set_page_config(page_title="急診自動排班系統", layout="wide")
st.title("🏥 急診護理人員自動排班系統 (動態公平與強迫均分版)")
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

# ==========================================
# 左側設定區 
# ==========================================
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

def highlight_b1_r(val):
    try:
        if pd.to_numeric(val) > 3: return 'background-color: #FFFF00'
    except: pass
    return ''

st.markdown("---")
if st.button("🚀 開始自動排班運算", disabled=not data_ready):
    with st.spinner("🧠 執行天然均分與任務動態解鎖中..."):
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

            # 🌟 更新：精準白名單 (補上 T2 等)
            team_allowed_zones = {
                'A': ["L", "T", "T2", "B1", "C1", "B2", "C2", "R", "R1", "R3", "MO", "MO1", "MO2", "S2", "S", "GC", "GB", "GD"],
                'B': ["T", "T2", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "GC", "GB", "GD"],
                'C': ["A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "GC", "GB", "GD"],
                'D': ["A1", "B1", "C1", "A2", "B2", "C2", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "GC", "GB", "GD"],
                'E': ["A1", "C1", "A2", "B2", "R1", "R2", "P", "MO", "MO1", "MO2", "S1", "GB", "GC"],
                'F': ["A1", "A2", "R2", "MO", "MO1", "P", "S1"],
                'G': ["A1", "R2", "MO", "MO1", "S1"]
            }
            
            master_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "P", "MO", "MO1", "MO2", "S1", "S", "S2", "R3", "T2", "GB", "GC", "GD"]

            for day_idx, date_col in enumerate(date_columns):
                todays_shifts = df_shift[['姓名', '組別', '性別', date_col]].copy()
                todays_shifts.columns = ['姓名', '組別', '性別', '班別']
                todays_template = df_template[['姓名', date_col]].copy()
                todays_template.columns = ['姓名', '預設區域']
                daily_data = pd.merge(todays_shifts, todays_template, on='姓名')
                
                # 🌟 每天開始前，全域掃描 A、B 組是否都已經去過 GB/GC
                ab_staff_all = [s for s in all_staff if get_team_of(s, df_shift) in ['A', 'B']]
                ab_needs_gbgc = any(sum(monthly_counts[s].get(x, 0) for x in ['GB', 'GC']) == 0 for s in ab_staff_all)

                for shift_type in ['D', 'E', 'N']:
                    shift_staff = daily_data[daily_data['班別'].astype(str).str.upper() == shift_type].copy()
                    if len(shift_staff) == 0: continue
                    
                    unassigned_staff = shift_staff['姓名'].tolist()
                    assignments = {}
                    pre_assigned_zones = []
                    
                    for _, row in shift_staff.iterrows():
                        name, preset = row['姓名'], str(row['預設區域']).strip().upper()
                        if preset not in ['X', 'NAN', 'NONE', ''] and name in unassigned_staff:
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

                    total_needed = len(unassigned_staff) + len(pre_assigned_zones)
                    base_zones = (master_zones * 2)[:total_needed] 
                    available_zones = base_zones.copy()
                    
                    for pz in pre_assigned_zones:
                        if pz in available_zones:
                            available_zones.remove(pz)
                        elif available_zones:
                            available_zones.pop()

                    # 連啟倫 MO系 霸王鎖定
                    if "N連啟倫" in unassigned_staff:
                        valid_lian = [z for z in ['MO', 'MO1', 'MO2'] if z in available_zones]
                        if valid_lian:
                            valid_lian.sort(key=lambda z: monthly_counts.get("N連啟倫", {}).get(z, 0))
                            chosen = valid_lian[0]
                            assignments["N連啟倫"] = chosen
                            unassigned_staff.remove("N連啟倫")
                            available_zones.remove(chosen)
                        else:
                            chosen = "MO" if monthly_counts.get("N連啟倫", {}).get("MO",0) <= monthly_counts.get("N連啟倫", {}).get("MO1",0) else "MO1"
                            assignments["N連啟倫"] = chosen
                            unassigned_staff.remove("N連啟倫")
                            if available_zones: available_zones.pop()

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
                                    is_team_G = (get_team_of(name, shift_staff) == 'G')
                                    if is_train or is_TP or is_team_G:
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
                    
                    # 建立今日大區組別追蹤
                    current_day_macro_teams = {}
                    for assigned_name, assigned_z in assignments.items():
                        t_val = get_team_of(assigned_name, shift_staff)
                        m_val = get_macro(assigned_z)
                        if m_val not in current_day_macro_teams: current_day_macro_teams[m_val] = set()
                        current_day_macro_teams[m_val].add(t_val)

                    # F/E 組優先配區
                    for name in [n for n in list(unassigned_staff) if get_team_of(n, shift_staff) == 'F']:
                        a_zones = [z for z in ["A1", "R2", "MO", "MO1", "S1"] if z in available_zones]
                        if a_zones:
                            z_scores = [(z, get_zone_score(z, name, 'F', day_idx, df_result, date_columns, monthly_counts, work_blocks, current_day_macro_teams, ab_needs_gbgc)) for z in a_zones]
                            z_scores.sort(key=lambda x: x[1])
                            best_z, best_s = z_scores[0]
                            if best_s < 20000: 
                                assignments[name] = best_z
                                unassigned_staff.remove(name)
                                available_zones.remove(best_z)
                                m_val = get_macro(best_z)
                                if m_val not in current_day_macro_teams: current_day_macro_teams[m_val] = set()
                                current_day_macro_teams[m_val].add('F')

                    for name in [n for n in list(unassigned_staff) if get_team_of(n, shift_staff) == 'E']:
                        if "A2" in available_zones:
                            if get_zone_score("A2", name, 'E', day_idx, df_result, date_columns, monthly_counts, work_blocks, current_day_macro_teams, ab_needs_gbgc) < 20000:
                                assignments[name] = "A2"
                                unassigned_staff.remove(name)
                                available_zones.remove("A2")
                                m_val = get_macro("A2")
                                if m_val not in current_day_macro_teams: current_day_macro_teams[m_val] = set()
                                current_day_macro_teams[m_val].add('E')

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

                        z_scores = [(z, get_zone_score(z, name, team, day_idx, df_result, date_columns, monthly_counts, work_blocks, current_day_macro_teams, ab_needs_gbgc)) for z in valid_cands]
                        z_scores.sort(key=lambda x: x[1])
                        
                        chosen_zone = z_scores[0][0]
                        assignments[name] = chosen_zone
                        unassigned_staff.remove(name)
                        if chosen_zone in available_zones: available_zones.remove(chosen_zone)
                        
                        m_val = get_macro(chosen_zone)
                        if m_val not in current_day_macro_teams: current_day_macro_teams[m_val] = set()
                        current_day_macro_teams[m_val].add(team)

                    for name, assigned_zone in assignments.items():
                        df_result.loc[df_result['姓名'] == name, date_col] = assigned_zone
                        monthly_counts[name][assigned_zone] = monthly_counts[name].get(assigned_zone, 0) + 1

                progress_bar.progress((day_idx + 1) / len(date_columns))

            name_col_index = df_result.columns.get_loc('姓名')
            majority_shift_dict = {name: max(set([x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']]), key=[x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']].count) if [x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']] else "" for name in all_staff}
            df_result.insert(name_col_index, '當月班別', df_result['姓名'].map(majority_shift_dict))
            
            zone_count_order = ["L", "L2", "T", "T2", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "S1", "S2", "S", "P", "P2", "MO", "MO1", "MO2", "GB", "GC", "GD", "職代"]
            for count_zone in zone_count_order:
                df_result[count_zone] = df_result[date_columns].apply(lambda row: (row == count_zone).sum(), axis=1)
            
            # 合併結算欄位
            df_result['A1+B2+C2計'] = df_result[['A1', 'B2', 'C2']].sum(axis=1) if all(x in df_result.columns for x in ['A1', 'B2', 'C2']) else 0
            df_result['MO系計'] = df_result[['MO', 'MO1', 'MO2']].sum(axis=1) if all(x in df_result.columns for x in ['MO', 'MO1', 'MO2']) else 0
            df_result['GB+GC計'] = df_result[['GB', 'GC']].sum(axis=1) if all(x in df_result.columns for x in ['GB', 'GC']) else 0

            df_result = df_result.fillna("")

            st.success("🎉 排班完成！已解鎖 GB/GC 讓路機制，並針對 B1, C1, R 實施絕對均分！")
            st.dataframe(df_result.head(10))

            styled_df = df_result.style.map(highlight_b1_r, subset=['B1', 'R'] if 'B1' in df_result.columns and 'R' in df_result.columns else [])

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                styled_df.to_excel(writer, index=False, sheet_name='排班結果')
                df_shift.to_excel(writer, index=False, sheet_name='原始班表')
            st.download_button("📥 下載最終排班表 (Excel 螢光黃標示版)", data=output.getvalue(), file_name="排班結果_動態平衡版.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"發生內部錯誤：{e}")
