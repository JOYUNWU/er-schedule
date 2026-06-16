import streamlit as st
import pandas as pd
import numpy as np
import random
import io
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

# ==========================================
# 0. 頂層小幫手函數
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
# 網頁 UI 初始化
# ==========================================
st.set_page_config(page_title="自動排班系統", layout="wide")
st.title("急診護理人員自動初步排班系統")
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

st.markdown("---")
if st.button("🚀 開始自動排班運算", disabled=not data_ready):
    with st.spinner("🚀 引擎啟動，正在載入動態 Excel 公式與缺乏組長統計..."):
        try:
            df_result = df_template.copy()
            monthly_counts = {name: {} for name in all_staff}
            progress_bar = st.progress(0)
            
            # 🌟 新增：用來記錄哪天哪班缺乏組長的儲存容器
            missing_l_records = []
            
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
                'A': ["L", "T", "T2", "B1", "C1", "B2", "C2", "R", "R1", "R3", "MO", "MO1", "MO2", "S2", "S", "GC", "GB", "GD"],
                'B': ["T", "T2", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2", "GC", "GB", "GD"],
                'C': ["A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "GC", "GB", "GD"],
                'D': ["A1", "B1", "C1", "A2", "B2", "C2", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "GC", "GB", "GD"],
                'E': ["A1", "C1", "A2", "B2", "R1", "R2", "P", "MO", "MO1", "MO2", "S1", "GB", "GC"],
                'F': ["A1", "A2", "R2", "MO", "MO1", "P", "S1"],
                'G': ["A1", "R2", "MO", "MO1", "S1"]
            }
            
            master_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "P", "MO", "MO1", "MO2", "S1", "S", "S2", "R3", "T2", "GB", "GC", "GD"]

            for day_idx, date_col in enumerate(date_columns):
                
                ab_staff = [n for n in all_staff if get_team_of(n, df_shift) in ['A', 'B']]
                ab_gb_gc_fulfilled = all(sum(monthly_counts[n].get(x, 0) for x in ['GB', 'GC']) >= 1 for n in ab_staff)

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
                        
                        # 🌟 核心防線：如果走完三個順位，該班別依然沒有組長(L)，將其記錄下來！
                        if "L" not in assignments.values():
                            missing_l_records.append({
                                '日期': f"2026/06/{str(date_col).zfill(2)}",
                                '缺乏組長的班別': f"{shift_type} 班",
                                '異常提示': "⚠️ 該班別設定的三個順位組長今天全部休假，缺少L，請手動指派"
                            })

                    total_needed = len(unassigned_staff) + len(pre_assigned_zones)
                    
                    base_zones = (master_zones * 2)[:total_needed] 
                    available_zones = base_zones.copy()
                    
                    for pz in pre_assigned_zones:
                        if pz in available_zones:
                            available_zones.remove(pz)
                        elif available_zones:
                            available_zones.pop() 

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
                    
                    current_day_macro_teams = {}
                    for assigned_name, assigned_z in assignments.items():
                        t_val = get_team_of(assigned_name, shift_staff)
                        m_val = get_macro(assigned_z)
                        if m_val not in current_day_macro_teams: current_day_macro_teams[m_val] = set()
                        current_day_macro_teams[m_val].add(t_val)

                    random.shuffle(unassigned_staff)
                    unassigned_staff.sort(key=lambda x: get_team_priority(x, shift_staff))

                    for name in list(unassigned_staff):
                        gender = shift_staff[shift_staff['姓名'] == name]['性別'].values[0] if len(shift_staff[shift_staff['姓名'] == name]) > 0 else ""
                        team = get_team_of(name, shift_staff)
                        team_allowed = team_allowed_zones.get(team, available_zones)

                        valid_cands = [z for z in available_zones if z in team_allowed and not (z == "S2" and str(gender).upper() == "M")]
                        
                        if not valid_cands: 
                            valid_cands = [z for z in available_zones if not (z == "S2" and str(gender).upper() == "M")]
                            if team == 'A':
                                valid_cands = [z for z in valid_cands if z not in ['R2', 'A2', 'S1']]
                            if not valid_cands:
                                valid_cands = available_zones.copy()

                        z_scores = [(z, get_zone_score(z, name, team, day_idx, df_result, date_columns, monthly_counts, work_blocks, current_day_macro_teams, ab_gb_gc_fulfilled)) for z in valid_cands]
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

            # 插入班別總結欄位
            name_col_index = df_result.columns.get_loc('姓名')
            majority_shift_dict = {name: max(set([x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']]), key=[x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']].count) if [x for x in df_shift[df_shift['姓名'] == name][date_columns].values.flatten() if x in ['D', 'E', 'N']] else "" for name in all_staff}
            df_result.insert(name_col_index, '當月班別', df_result['姓名'].map(majority_shift_dict))
            
            # 準備統計欄位順序
            zone_count_order = ["L", "L2", "T", "T2", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "S1", "S2", "S", "P", "P2", "MO", "MO1", "MO2", "GB", "GC", "GD", "職代"]
            summary_cols = ['A2+B2+C2計', 'MO系計', 'R1+R3計', 'GB+GC計']
            
            # 🌟 複製一份專門用來產生 Excel 的 DataFrame
            df_excel = df_result.copy()

            # (1) 計算 df_result 的靜態數字，用來在網頁介面預覽
            for count_zone in zone_count_order:
                df_result[count_zone] = df_result[date_columns].apply(lambda row: (row == count_zone).sum(), axis=1)
            
            df_result['A2+B2+C2計'] = df_result[['A2', 'B2', 'C2']].sum(axis=1) if all(x in df_result.columns for x in ['A2', 'B2', 'C2']) else 0
            df_result['MO系計'] = df_result[['MO', 'MO1', 'MO2']].sum(axis=1) if all(x in df_result.columns for x in ['MO', 'MO1', 'MO2']) else 0
            df_result['R1+R3計'] = df_result[['R1', 'R3']].sum(axis=1) if all(x in df_result.columns for x in ['R1', 'R3']) else 0
            df_result['GB+GC計'] = df_result[['GB', 'GC']].sum(axis=1) if all(x in df_result.columns for x in ['GB', 'GC']) else 0
            df_result = df_result.fillna("")

            # (2) 替 df_excel 注入原生的 Excel 公式 (=COUNTIF)
            first_date_col_idx = df_excel.columns.get_loc(date_columns[0]) + 1
            last_date_col_idx = df_excel.columns.get_loc(date_columns[-1]) + 1
            col_start = get_column_letter(first_date_col_idx)
            col_end = get_column_letter(last_date_col_idx)

            for col in zone_count_order + summary_cols:
                if col not in df_excel.columns:
                    df_excel[col] = ""

            for i in range(len(df_excel)):
                row_excel = i + 2 
                range_str = f"{col_start}{row_excel}:{col_end}{row_excel}"
                
                for z in zone_count_order:
                    df_excel.at[i, z] = f'=COUNTIF({range_str}, "{z}")'
                
                df_excel.at[i, 'A2+B2+C2計'] = f'=COUNTIF({range_str}, "A2")+COUNTIF({range_str}, "B2")+COUNTIF({range_str}, "C2")'
                df_excel.at[i, 'MO系計'] = f'=COUNTIF({range_str}, "MO")+COUNTIF({range_str}, "MO1")+COUNTIF({range_str}, "MO2")'
                df_excel.at[i, 'R1+R3計'] = f'=COUNTIF({range_str}, "R1")+COUNTIF({range_str}, "R3")'
                df_excel.at[i, 'GB+GC計'] = f'=COUNTIF({range_str}, "GB")+COUNTIF({range_str}, "GC")'

            # 🌟 轉換缺乏組長名單成 DataFrame，若無任何記錄，則建立防走失完美提示！
            if len(missing_l_records) == 0:
                df_missing_l = pd.DataFrame([{
                    '日期': '🎉 完美平衡！',
                    '缺乏組長的班別': '全月皆有組長',
                    '異常提示': '本月天天都有組長，無任何缺乏組長之班別，不需進行人工修改。'
                }])
            else:
                df_missing_l = pd.DataFrame(missing_l_records)

            st.success("🎉 排班完成！已全面復原『第三分頁：缺乏組長提示』與動梯函數！")
            
            # 網頁預覽
            def preview_highlight(val):
                try:
                    v = pd.to_numeric(val)
                    if v == 0: return 'background-color: #FFFF00; color: #000000'
                    elif v > 5: return 'background-color: #FFDAB9; color: #000000'
                except: pass
                return ''
            
            preview_df = df_result.head(10).copy()
            target_cols = [c for c in (zone_count_order + summary_cols) if c in preview_df.columns]
            st.dataframe(preview_df.style.map(preview_highlight, subset=target_cols))

            # 🌟 寫出包含真正 Excel 函數、條件格式與第三分頁的檔案
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False, sheet_name='排班結果')
                df_shift.to_excel(writer, index=False, sheet_name='原始班表')
                df_missing_l.to_excel(writer, index=False, sheet_name='缺乏組長提示') # 🌟 第三分頁隆重歸位
                
                ws = writer.sheets['排班結果']
                
                start_summary_col_idx = df_excel.columns.get_loc(zone_count_order[0]) + 1
                end_summary_col_idx = len(df_excel.columns)
                start_sum_col_letter = get_column_letter(start_summary_col_idx)
                end_sum_col_letter = get_column_letter(end_summary_col_idx)
                
                format_range = f"{start_sum_col_letter}2:{end_sum_col_letter}{len(df_excel)+1}"
                
                yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                peach_fill = PatternFill(start_color='FFDAB9', end_color='FFDAB9', fill_type='solid')
                
                ws.conditional_formatting.add(format_range, CellIsRule(operator='equal', formula=['0'], fill=yellow_fill))
                ws.conditional_formatting.add(format_range, CellIsRule(operator='greaterThan', formula=['5'], fill=peach_fill))

            st.download_button("📥 下載最終排班表", data=output.getvalue(), file_name="初步自動生成排班結果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"發生內部錯誤：{e}")
