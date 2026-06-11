import streamlit as st
import pandas as pd
import numpy as np
import random
import io

# ==========================================
# 0. 頂層小幫手函數 (拉到最外層，徹底解決縮排報錯問題)
# ==========================================
def get_team_of(staff_name, df_daily):
    ts = df_daily[df_daily['姓名'] == staff_name]['組別'].values
    return str(ts[0]).strip().upper() if len(ts) > 0 else ""

def get_team_priority(staff_name, df_daily):
    t = get_team_of(staff_name, df_daily)
    if t == 'D': return 1
    elif t == 'C': return 2
    elif t == 'B': return 3
    else: return 4

def get_zone_count(z, m_counts, s_name):
    if z in ['A2', 'B2', 'C2']: 
        return sum(m_counts[s_name].get(x, 0) for x in ['A2', 'B2', 'C2'])
    if z in ['GB', 'GC', 'GD']: 
        return sum(m_counts[s_name].get(x, 0) for x in ['GB', 'GC', 'GD'])
    return m_counts[s_name].get(z, 0)

# ==========================================
# 網頁 UI 初始化
# ==========================================
st.set_page_config(page_title="急診自動排班系統", layout="wide")
st.title("🏥 急診護理人員自動排班系統 (全功能整合版)")
st.markdown("---")

# ==========================================
# 1. 檔案上傳區塊
# ==========================================
col1, col2 = st.columns(2)
with col1:
    training_file = st.file_uploader("📂 1. 上傳班表 (training 檔案)", type=['xlsx', 'csv'])
with col2:
    template_file = st.file_uploader("📂 2. 上傳空白檔", type=['xlsx', 'csv'])

all_staff = []
df_shift = None
df_template = None
data_ready = False
date_columns = []

if training_file and template_file:
    try:
        df_shift = pd.read_csv(training_file) if 'csv' in training_file.name.lower() else pd.read_excel(training_file)
        df_template = pd.read_csv(template_file) if 'csv' in template_file.name.lower() else pd.read_excel(template_file)
            
        shift_cols = list(df_shift.columns)
        shift_cols[0] = '編號'
        shift_cols[1], shift_cols[2], shift_cols[3] = '組別', '性別', '姓名'
        date_length = len(shift_cols) - 4
        shift_cols[4:] = [str(i) for i in range(1, date_length + 1)]
        df_shift.columns = shift_cols
        
        template_cols = list(df_template.columns)
        template_cols[0] = '編號'
        template_cols[1], template_cols[2], template_cols[3] = '組別', '性別', '姓名'
        template_cols[4:4+date_length] = [str(i) for i in range(1, date_length + 1)]
        df_template.columns = template_cols
        
        df_shift = df_shift.dropna(subset=['姓名'])
        df_template = df_template.dropna(subset=['姓名'])
        df_shift['姓名'] = df_shift['姓名'].astype(str).str.strip()
        df_template['姓名'] = df_template['姓名'].astype(str).str.strip()
        
        all_staff = df_shift['姓名'].unique().tolist()
        date_columns = df_shift.columns[4:]
        data_ready = True
        
        st.success("✅ 檔案讀取成功！人員名單已載入，請於左側設定規則並開始排班。")
    except Exception as e:
        st.error(f"檔案讀取失敗，請檢查格式：{e}")

# ==========================================
# 2. 左側邊欄：動態規則設定
# ==========================================
st.sidebar.header("⚙️ 本月特殊排班規則設定")

if not data_ready:
    st.sidebar.warning("⚠️ 請先上傳檔案，系統才能載入人員名單供您選取。")

st.sidebar.subheader("🎓 訓練人員綁定 (整月固定區)")
train_s2 = st.sidebar.multiselect("S2 訓練名單", options=all_staff)
train_b1_r = st.sidebar.multiselect("B1/R/C2/S 訓練名單", options=all_staff)
train_t = st.sidebar.multiselect("檢傷(T) 訓練名單", options=all_staff)

st.sidebar.markdown("---")
st.sidebar.subheader("👑 各班組長 (L) 順位設定")

def safe_options(default_val):
    return list(dict.fromkeys([default_val] + all_staff))

st.sidebar.markdown("**☀️ 白班 (D) 組長**")
d_l_1 = st.sidebar.selectbox("D班 第一順位", safe_options("AHN黃麗婷"))
d_l_2 = st.sidebar.selectbox("D班 第二順位", safe_options("N3尤美惠"))
d_l_3 = st.sidebar.selectbox("D班 第三順位", safe_options("N3許嘉文"))

st.sidebar.markdown("**🌆 小夜班 (E) 組長**")
e_l_1 = st.sidebar.selectbox("E班 第一順位", safe_options("AHN蕭惠澤"))
e_l_2 = st.sidebar.selectbox("E班 第二順位", safe_options("N2李曉侖"))
e_l_3 = st.sidebar.selectbox("E班 第三順位", safe_options("N3黃義全"))

st.sidebar.markdown("**🌙 大夜班 (N) 組長**")
n_l_1 = st.sidebar.selectbox("N班 第一順位", safe_options("N3許慧芳"))
n_l_2 = st.sidebar.selectbox("N班 第二順位", safe_options("N2江品儒"))
n_l_3 = st.sidebar.selectbox("N班 第三順位", safe_options("N1許家瑄"))

# ==========================================
# 3. 排班核心運算區塊
# ==========================================
st.markdown("---")

if st.button("🚀 開始自動排班運算 (套用上述規則)", disabled=not data_ready):
    with st.spinner("🧠 啟動究極防撞與預排霸王機制..."):
        try:
            df_result = df_template.copy()
            tp_tracker = {name: None for name in all_staff}
            train_brcs_tracker = {name: None for name in all_staff}
            history_tracker = {name: [] for name in all_staff}
            monthly_counts = {name: {} for name in all_staff}
            progress_bar = st.progress(0)
            
            zone_groups = {
                "MO": "MO_G", "MO1": "MO_G", "MO2": "MO_G",
                "S": "S_G", "S1": "S_G", "S2": "S_G",
                "A1": "ABC_G", "B1": "ABC_G", "C1": "ABC_G", "A2": "ABC_G", "B2": "ABC_G", "C2": "ABC_G",
                "R": "R_G", "R1": "R_G", "R2": "R_G", "R3": "R_G",
                "GB": "G_G", "GC": "G_G", "GD": "G_G"
            }
            
            team_allowed_zones = {
                'A': ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2", "GC", "GB", "GD"],
                'B': ["A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "GC", "GB", "GD"],
                'C': ["A1", "C1", "A2", "B2", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2", "GC", "GB", "GD"],
                'D': ["A1", "R2", "MO", "MO1", "S1"]
            }
            
            MAX_DAYS = {
                "T": 3, "R": 3, "B1": 3, "S": 3, "P": 3, "C1": 3,
                "T2": 2, "GB": 2, "GC": 1, "GD": 1, "S2": 2
            }
            
            work_blocks = {name: [0]*len(date_columns) for name in all_staff}
            for name in all_staff:
                for day_idx in range(len(date_columns)):
                    count = 0
                    for future_idx in range(day_idx, len(date_columns)):
                        col = date_columns[future_idx]
                        val = str(df_shift.loc[df_shift['姓名'] == name, col].values[0]).strip().upper()
                        if val == 'OFF':
                            break
                        count += 1
                    work_blocks[name][day_idx] = count
            
            for day_idx, date_col in enumerate(date_columns):
                todays_shifts = df_shift[['姓名', '組別', '性別', date_col]].copy()
                todays_shifts.columns = ['姓名', '組別', '性別', '班別']
                todays_template = df_template[['姓名', date_col]].copy()
                todays_template.columns = ['姓名', '預設區域']
                daily_data = pd.merge(todays_shifts, todays_template, on='姓名')
                
                for shift_type in ['D', 'E', 'N']:
                    shift_staff = daily_data[daily_data['班別'].astype(str).str.upper() == shift_type].copy()
                    staff_count = len(shift_staff)
                    if staff_count == 0: continue
                    
                    unassigned_staff = shift_staff['姓名'].tolist()
                    assignments = {}
                    
                    # 1. 決定基礎區域 (基於當班總人數)
                    if staff_count <= 17:
                        base_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "P", "MO", "MO1", "MO2", "S1", "S"]
                    elif staff_count == 18:
                        base_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "P", "MO", "MO1", "MO2", "S1", "S2", "S"]
                    elif staff_count == 19:
                        base_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S"]
                    elif staff_count == 20:
                        base_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2"]
                    elif staff_count == 21:
                        base_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2", "GB"]
                    elif staff_count == 22:
                        base_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2", "GB", "GC"]
                    else:
                        base_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2", "GB", "GC", "GD"]

                    available_zones = base_zones.copy()
                    target_zone_count = staff_count - 1 # 扣除 L 的隱形名額
                    
                    while len(available_zones) < target_zone_count:
                        available_zones.append(f"支援{len(available_zones)+1}")
                    if len(available_zones) > target_zone_count:
                        available_zones = available_zones[:target_zone_count]
                    
                    # 🌟 2. 不動檔霸王條款 (最優先執行！)
                    for _, row in shift_staff.iterrows():
                        name, preset = row['姓名'], str(row['預設區域']).strip()
                        if preset.upper() not in ['X', 'NAN', 'NONE', ''] and name in unassigned_staff:
                            assignments[name] = preset
                            unassigned_staff.remove(name)
                            if preset == "L":
                                pass
                            else:
                                if preset in available_zones: 
                                    available_zones.remove(preset)
                                elif available_zones: 
                                    available_zones.pop()
                    
                    # 🌟 3. L 順位指派 (如果 L 還沒被手動佔走才派)
                    if "L" not in assignments.values():
                        if shift_type == 'D': l_chain = [d_l_1, d_l_2, d_l_3]
                        elif shift_type == 'E': l_chain = [e_l_1, e_l_2, e_l_3]
                        elif shift_type == 'N': l_chain = [n_l_1, n_l_2, n_l_3]
                        else: l_chain = []

                        for l_candidate in l_chain:
                            if l_candidate in unassigned_staff:
                                assignments[l_candidate] = "L"
                                unassigned_staff.remove(l_candidate)
                                break
                    
                    # 4. 訓練名單優先綁定
                    for name in list(unassigned_staff):
                        if name in train_s2:
                            assignments[name] = "S2"
                            unassigned_staff.remove(name)
                            if "S2" in available_zones: available_zones.remove("S2")
                            elif available_zones: available_zones.pop()
                        elif name in train_t:
                            assignments[name] = "T"
                            unassigned_staff.remove(name)
                            if "T" in available_zones: available_zones.remove("T")
                            elif available_zones: available_zones.pop()
                        elif name in train_b1_r:
                            prev_train_zone = train_brcs_tracker.get(name)
                            if prev_train_zone in ["B1", "R", "S", "C2"]:
                                assignments[name] = prev_train_zone
                                unassigned_staff.remove(name)
                                if prev_train_zone in available_zones: available_zones.remove(prev_train_zone)
                                elif available_zones: available_zones.pop()
                            else:
                                valid_train_zones = ["B1", "R", "S", "C2"]
                                valid_train_zones.sort(key=lambda z: (monthly_counts[name].get(z, 0), ["B1", "R", "S", "C2"].index(z)))
                                chosen_train = valid_train_zones[0]
                                assignments[name] = chosen_train
                                unassigned_staff.remove(name)
                                train_brcs_tracker[name] = chosen_train
                                if chosen_train in available_zones: available_zones.remove(chosen_train)
                                elif available_zones: available_zones.pop()
                            
                    # 5. T/P 連續狀態處理
                    for name in list(unassigned_staff):
                        prev_tp = tp_tracker.get(name)
                        if prev_tp in ["T", "P"]:
                            rem_days = work_blocks[name][day_idx]
                            if monthly_counts[name].get(prev_tp, 0) + rem_days <= MAX_DAYS.get(prev_tp, 3):
                                assignments[name] = prev_tp
                                unassigned_staff.remove(name)
                                if prev_tp in available_zones: available_zones.remove(prev_tp)
                                elif available_zones: available_zones.pop()
                            else:
                                tp_tracker[name] = None
                                
                    # 6. 連啟倫鎖定
                    if "N連啟倫" in unassigned_staff:
                        chosen = "MO" if monthly_counts["N連啟倫"].get("MO",0) <= monthly_counts["N連啟倫"].get("MO1",0) else "MO1"
                        assignments["N連啟倫"] = chosen
                        unassigned_staff.remove("N連啟倫")
                        if chosen in available_zones: available_zones.remove(chosen)
                        elif available_zones: available_zones.pop()
                    
                    # 7. 剩餘隨機與跨組均分分配
                    random.shuffle(unassigned_staff)
                    unassigned_staff.sort(key=lambda x: get_team_priority(x, shift_staff))

                    shift_cluster_teams = {}
                    for assigned_n, assigned_z in assignments.items():
                        t_val = get_team_of(assigned_n, shift_staff)
                        c_val = zone_groups.get(assigned_z, assigned_z)
                        if c_val not in shift_cluster_teams: shift_cluster_teams[c_val] = set()
                        shift_cluster_teams[c_val].add(t_val)

                    for name in list(unassigned_staff):
                        gender_series = shift_staff[shift_staff['姓名'] == name]['性別'].values
                        gender = gender_series[0] if len(gender_series) > 0 else ""
                        team = get_team_of(name, shift_staff)
                        recent_groups = history_tracker[name][-2:] if name in history_tracker else []

                        team_allowed = team_allowed_zones.get(team, available_zones)

                        candidate_zones = []
                        for zone in list(available_zones):
                            if zone not in team_allowed: continue
                            if zone == "S2" and str(gender).upper() == "M": continue
                            if zone in ["T", "P"]:
                                rem_days = work_blocks[name][day_idx]
                                if monthly_counts[name].get(zone, 0) + rem_days > MAX_DAYS.get(zone, 3): continue
                            elif zone in MAX_DAYS and monthly_counts[name].get(zone, 0) >= MAX_DAYS[zone]: continue
                            candidate_zones.append(zone)

                        if not candidate_zones:
                            for zone in list(available_zones):
                                if zone not in team_allowed: continue
                                if zone == "S2" and str(gender).upper() == "M": continue
                                candidate_zones.append(zone)

                        if not candidate_zones:
                            for zone in list(available_zones):
                                if zone == "S2" and str(gender).upper() == "M": continue
                                candidate_zones.append(zone)

                        if not candidate_zones: 
                            candidate_zones = available_zones.copy() if available_zones else ["支援"]

                        pref_candidate_zones = candidate_zones.copy()
                        
                        diverse_zones = []
                        for z in pref_candidate_zones:
                            c_val = zone_groups.get(z, z)
                            if team not in shift_cluster_teams.get(c_val, set()):
                                diverse_zones.append(z)
                        if diverse_zones:
                            pref_candidate_zones = diverse_zones

                        best_zones = [z for z in pref_candidate_zones if zone_groups.get(z, z) not in recent_groups]

                        if best_zones:
                            best_zones.sort(key=lambda z: get_zone_count(z, monthly_counts, name))
                            min_count = get_zone_count(best_zones[0], monthly_counts, name)
                            lowest_zones = [z for z in best_zones if get_zone_count(z, monthly_counts, name) == min_count]
                            chosen_zone = random.choice(lowest_zones)
                        else:
                            last_group = history_tracker[name][-1:] if name in history_tracker else []
                            better_zones = [z for z in pref_candidate_zones if zone_groups.get(z, z) not in last_group]
                            if better_zones:
                                better_zones.sort(key=lambda z: get_zone_count(z, monthly_counts, name))
                                min_count = get_zone_count(better_zones[0], monthly_counts, name)
                                lowest_zones = [z for z in better_zones if get_zone_count(z, monthly_counts, name) == min_count]
                                chosen_zone = random.choice(lowest_zones)
                            else:
                                pref_candidate_zones.sort(key=lambda z: get_zone_count(z, monthly_counts, name))
                                min_count = get_zone_count(pref_candidate_zones[0], monthly_counts, name)
                                lowest_zones = [z for z in pref_candidate_zones if get_zone_count(z, monthly_counts, name) == min_count]
                                chosen_zone = random.choice(lowest_zones)

                        assignments[name] = chosen_zone
                        unassigned_staff.remove(name)
                        available_zones.remove(chosen_zone)
                        
                        assigned_cluster = zone_groups.get(chosen_zone, chosen_zone)
                        if assigned_cluster not in shift_cluster_teams: shift_cluster_teams[assigned_cluster] = set()
                        shift_cluster_teams[assigned_cluster].add(team)

                    for name, assigned_zone in assignments.items():
                        df_result.loc[df_result['姓名'] == name, date_col] = assigned_zone
                        group = zone_groups.get(assigned_zone, assigned_zone)
                        history_tracker[name].append(group)
                        monthly_counts[name][assigned_zone] = monthly_counts[name].get(assigned_zone, 0) + 1

                        if assigned_zone in ["T", "P"]: tp_tracker[name] = assigned_zone
                        else: tp_tracker[name] = None

                todays_off_staff = df_shift[df_shift[date_col].astype(str).str.upper() == 'OFF']['姓名'].tolist()
                for off_name in todays_off_staff:
                    tp_tracker[off_name] = None
                    train_brcs_tracker[off_name] = None

                progress_bar.progress((day_idx + 1) / len(date_columns))

            majority_shift_dict = {}
            for name in all_staff:
                row_data = df_shift[df_shift['姓名'] == name][date_columns].values.flatten()
                valid_shifts = [str(x).strip().upper() for x in row_data if str(x).strip().upper() in ['D', 'E', 'N']]
                major_shift = max(set(valid_shifts), key=valid_shifts.count) if valid_shifts else ""
                majority_shift_dict[name] = major_shift

            name_col_index = df_result.columns.get_loc('姓名')
            df_result.insert(name_col_index, '當月班別', df_result['姓名'].map(majority_shift_dict))

            zone_count_order = ["L", "L2", "T", "T2", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "S1", "S2", "S", "P", "P2", "MO", "MO1", "MO2", "GB", "GC", "GD", "職代"]
            for count_zone in zone_count_order:
                df_result[count_zone] = df_result[date_columns].apply(lambda row: (row == count_zone).sum(), axis=1)

            summary_total = {'姓名': '各班獨立人數'}
            summary_d = {'姓名': 'D'}
            summary_e = {'姓名': 'E'}
            summary_n = {'姓名': 'N'}

            for date_col in date_columns:
                daily_shifts = df_shift[date_col].astype(str).str.upper()
                d_count = (daily_shifts == 'D').sum()
                e_count = (daily_shifts == 'E').sum()
                n_count = (daily_shifts == 'N').sum()

                summary_d[date_col] = d_count
                summary_e[date_col] = e_count
                summary_n[date_col] = n_count
                summary_total[date_col] = d_count + e_count + n_count

            df_summary = pd.DataFrame([summary_total, summary_d, summary_e, summary_n])
            df_result = pd.concat([df_result, df_summary], ignore_index=True)
            df_result = df_result.fillna("")

            st.success("🎉 排班運算完成！已啟動「絕對尊重不動檔」與「組長遞補機制」。")
            st.dataframe(df_result.head(10))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='排班結果')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 下載最終排班表 (Excel)", 
                data=excel_data, 
                file_name="排班結果_不動檔霸王條款版.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"執行時發生內部錯誤：{e}")
