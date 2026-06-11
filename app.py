import streamlit as st
import pandas as pd
import numpy as np
import random
import io

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
        shift_cols[1], shift_cols[2], shift_cols[3] = '組別', '性別', '姓名'
        df_shift.columns = shift_cols
        
        template_cols = list(df_template.columns)
        template_cols[1], template_cols[2], template_cols[3] = '組別', '性別', '姓名'
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
    with st.spinner("🧠 啟動未來預測機制與絕對連續邏輯運算中..."):
        try:
            df_result = df_template.copy()
            tp_tracker = {name: None for name in all_staff}
            history_tracker = {name: [] for name in all_staff}
            monthly_counts = {name: {} for name in all_staff}
            progress_bar = st.progress(0)
            
            zone_groups = {
                "MO": "MO_G", "MO1": "MO_G", "MO2": "MO_G",
                "S": "S_G", "S1": "S_G", "S2": "S_G",
                "A1": "ABC_G", "B1": "ABC_G", "C1": "ABC_G", "A2": "ABC_G", "B2": "ABC_G", "C2": "ABC_G",
                "R": "R_G", "R1": "R_G", "R2": "R_G", "R3": "R_G"
            }
            team_a_avoids = ["A1", "R2", "MO", "S1", "C1"]
            
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
                    
                    base_zones = ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S"]
                    if staff_count >= 19: base_zones.append("T2")
                    if staff_count >= 20: base_zones.append("GB")
                    if staff_count >= 21: base_zones.append("GC")
                    if staff_count >= 22: base_zones.append("GD")
                    
                    available_zones = base_zones.copy()
                    unassigned_staff = shift_staff['姓名'].tolist()
                    assignments = {}
                    
                    # 1. 不動檔優先
                    for _, row in shift_staff.iterrows():
                        name, preset = row['姓名'], str(row['預設區域']).strip()
                        if preset not in ['x', 'nan', 'None']:
                            assignments[name] = preset
                            unassigned_staff.remove(name)
                            if preset in available_zones: available_zones.remove(preset)
                    
                    # 2. 訓練名單優先綁定
                    for name in list(unassigned_staff):
                        if name in train_s2 and "S2" in available_zones:
                            assignments[name] = "S2"
                            unassigned_staff.remove(name)
                            available_zones.remove("S2")
                        elif name in train_t and "T" in available_zones:
                            assignments[name] = "T"
                            unassigned_staff.remove(name)
                            available_zones.remove("T")
                        elif name in train_b1_r:
                            for target_zone in ["B1", "R", "C2", "S"]:
                                if target_zone in available_zones:
                                    assignments[name] = target_zone
                                    unassigned_staff.remove(name)
                                    available_zones.remove(target_zone)
                                    break
                            
                    # 3. T/P 連續狀態處理
                    for name in list(unassigned_staff):
                        prev_tp = tp_tracker.get(name)
                        if prev_tp and prev_tp in available_zones:
                            assignments[name] = prev_tp
                            unassigned_staff.remove(name)
                            available_zones.remove(prev_tp)
                    
                    # 4. L 順位指派
                    if shift_type == 'D': l_chain = [d_l_1, d_l_2, d_l_3]
                    elif shift_type == 'E': l_chain = [e_l_1, e_l_2, e_l_3]
                    elif shift_type == 'N': l_chain = [n_l_1, n_l_2, n_l_3]
                    else: l_chain = []

                    for l_candidate in l_chain:
                        if l_candidate in unassigned_staff:
                            assignments[l_candidate] = "L"
                            unassigned_staff.remove(l_candidate)
                            break
                                
                    # 5. 連啟倫鎖定
                    if "N連啟倫" in unassigned_staff:
                        for mo_zone in ["MO", "MO1"]:
                            if mo_zone in available_zones:
                                assignments["N連啟倫"] = mo_zone
                                unassigned_staff.remove("N連啟倫")
                                available_zones.remove(mo_zone)
                                break
                    
                    # 6. 剩餘分配
                    random.shuffle(unassigned_staff) 
                    
                    for name in list(unassigned_staff):
                        gender_series = shift_staff[shift_staff['姓名'] == name]['性別'].values
                        gender = gender_series[0] if len(gender_series) > 0 else ""
                        
                        team_series = shift_staff[shift_staff['姓名'] == name]['組別'].values
                        team = str(team_series[0]).strip().upper() if len(team_series) > 0 else ""
                        
                        recent_groups = history_tracker[name][-2:] if name in history_tracker else []
                        
                        candidate_zones = []
                        for zone in list(available_zones):
                            if zone == "S2" and str(gender).upper() == "M": continue
                            if zone in ["T", "P"]:
                                rem_days = work_blocks[name][day_idx]
                                if monthly_counts[name].get(zone, 0) + rem_days > MAX_DAYS.get(zone, 3):
                                    continue
                            elif zone in MAX_DAYS and monthly_counts[name].get(zone, 0) >= MAX_DAYS[zone]:
                                continue
                            candidate_zones.append(zone)
                            
                        if not candidate_zones:
                            for zone in list(available_zones):
                                if zone == "S2" and str(gender).upper() == "M": continue
                                candidate_zones.append(zone)
                                
                        if not candidate_zones: continue 
                        
                        pref_candidate_zones = candidate_zones.copy()
                        if team == 'A':
                            filtered = [z for z in pref_candidate_zones if z not in team_a_avoids]
                            if len(filtered) > 0:
                                pref_candidate_zones = filtered
                        
                        def get_zone_count(z):
                            if z in ['A2', 'B2', 'C2']:
                                return sum(monthly_counts[name].get(x, 0) for x in ['A2', 'B2', 'C2'])
                            return monthly_counts[name].get(z, 0)
                        
                        best_zones = [z for z in pref_candidate_zones if zone_groups.get(z, z) not in recent_groups]
                        
                        if best_zones:
                            best_zones.sort(key=lambda z: get_zone_count(z))
                            min_count = get_zone_count(best_zones[0])
                            lowest_zones = [z for z in best_zones if get_zone_count(z) == min_count]
                            chosen_zone = random.choice(lowest_zones)
                        else:
                            last_group = history_tracker[name][-1:] if name in history_tracker else []
                            better_zones = [z for z in pref_candidate_zones if zone_groups.get(z, z) not in last_group]
                            if better_zones:
                                better_zones.sort(key=lambda z: get_zone_count(z))
                                min_count = get_zone_count(better_zones[0])
                                lowest_zones = [z for z in better_zones if get_zone_count(z) == min_count]
                                chosen_zone = random.choice(lowest_zones)
                            else:
                                pref_candidate_zones.sort(key=lambda z: get_zone_count(z))
                                min_count = get_zone_count(pref_candidate_zones[0])
                                lowest_zones = [z for z in pref_candidate_zones if get_zone_count(z) == min_count]
                                chosen_zone = random.choice(lowest_zones)

                        assignments[name] = chosen_zone
                        unassigned_staff.remove(name)
                        available_zones.remove(chosen_zone)
                        if chosen_zone in ["T", "P"]: tp_tracker[name] = chosen_zone

                    for name, assigned_zone in assignments.items():
                        df_result.loc[df_result['姓名'] == name, date_col] = assigned_zone
                        group = zone_groups.get(assigned_zone, assigned_zone)
                        history_tracker[name].append(group)
                        monthly_counts[name][assigned_zone] = monthly_counts[name].get(assigned_zone, 0) + 1

                todays_off_staff = df_shift[df_shift[date_col].astype(str).str.upper() == 'OFF']['姓名'].tolist()
                for off_name in todays_off_staff:
                    tp_tracker[off_name] = None
                    
                progress_bar.progress((day_idx + 1) / len(date_columns))
            
            # 插入當月班別
            majority_shift_dict = {}
            for name in all_staff:
                row_data = df_shift[df_shift['姓名'] == name][date_columns].values.flatten()
                valid_shifts = [str(x).strip().upper() for x in row_data if str(x).strip().upper() in ['D', 'E', 'N']]
                major_shift = max(set(valid_shifts), key=valid_shifts.count) if valid_shifts else ""
                majority_shift_dict[name] = major_shift
                
            name_col_index = df_result.columns.get_loc('姓名')
            df_result.insert(name_col_index, '當月班別', df_result['姓名'].map(majority_shift_dict))
            
            # 🌟 新增：在表格最下方加入每日班別獨立人數統計
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
            
            st.success("🎉 排班運算完成！已新增「當月班別」與「每日獨立人數統計」。")
            st.dataframe(df_result.head(10))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='排班結果')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 下載最終排班表 (Excel)", 
                data=excel_data, 
                file_name="排班結果_預測連續_含統計版.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"執行時發生內部錯誤：{e}")
