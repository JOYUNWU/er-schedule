import streamlit as st
import pandas as pd
import numpy as np
import random
import io

# ==========================================
# 0. 頂層小幫手函數 (含全新分區與重症定義)
# ==========================================
def get_team_of(staff_name, df_daily):
    ts = df_daily[df_daily['姓名'] == staff_name]['組別'].values
    return str(ts[0]).strip().upper() if len(ts) > 0 else ""

# 🌟 更新：適應 6 組的優先度排序 (限制越多的越優先排)
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
    if z in ['A2', 'B2', 'C2']: 
        return sum(m_counts[s_name].get(x, 0) for x in ['A2', 'B2', 'C2'])
    if z in ['GB', 'GC', 'GD']: 
        return sum(m_counts[s_name].get(x, 0) for x in ['GB', 'GC', 'GD'])
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

def get_forbidden_zones(name, day_idx, df_result, date_columns):
    past_2_zones = []
    for i in [1, 2]:
        if day_idx - i >= 0:
            pz = df_result.loc[df_result['姓名'] == name, date_columns[day_idx - i]].values[0]
            if pd.notna(pz) and str(pz).strip() not in ['OFF', '', 'X', 'NAN', 'L', 'L2', '職代']:
                past_2_zones.append(str(pz).strip())
    
    forbidden_macros = [get_macro(pz) for pz in past_2_zones]
    forbidden_severe = any(is_severe(pz) for pz in past_2_zones)
    
    return forbidden_macros, forbidden_severe

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
    with st.spinner("🧠 啟動究極防撞與連續截斷機制..."):
        try:
            df_result = df_template.copy()
            monthly_counts = {name: {} for name in all_staff}
            progress_bar = st.progress(0)
            
            # 🌟 更新：6組的允許區域設定
            team_allowed_zones = {
                'A': ["T", "A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "T2", "GC", "GB", "GD"],
                'B': ["A1", "B1", "C1", "A2", "B2", "C2", "R", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "S", "GC", "GB", "GD"],
                'C': ["A1", "C1", "A2", "B2", "R1", "R2", "R3", "P", "MO", "MO1", "MO2", "S1", "S2", "GC", "GB", "GD"],
                'D': ["A1", "C1", "A2", "R1", "R2", "P", "MO", "MO1", "S1"],
                'E': ["A1", "A2", "R2", "MO", "MO1", "P", "S1"],
                'F': ["A1", "R2", "MO", "MO1", "S1"]
            }
            
            MAX_DAYS = {
                "T": 3, "R": 3, "B1": 3, "S": 3, "P": 3, "C1": 3,
                "T2": 2, "GB": 2, "GC": 1, "GD": 1, "S2": 2
            }
            
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
                    target_zone_count = staff_count - 1 
                    
                    while len(available_zones) < target_zone_count:
                        available_zones.append(f"支援{len(available_zones)+1}")
                    if len(available_zones) > target_zone_count:
                        available_zones = available_zones[:target_zone_count]
                    
                    # 1. 處理手動預排
                    for _, row in shift_staff.iterrows():
                        name, preset = row['姓名'], str(row['預設區域']).strip()
                        if preset.upper() not in ['X', 'NAN', 'NONE', ''] and name in unassigned_staff:
                            assignments[name] = preset
                            unassigned_staff.remove(name)
                            if preset == "L": pass
                            else:
                                if preset in available_zones: available_zones.remove(preset)
                                elif available_zones: available_zones.pop()
                    
                    # 2. 處理 L 順位
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
                    
                    # =======================================================
                    # 🌟 3. 連續排班機制 (強制霸王保留區，遇 OFF 才會自動截斷)
                    # =======================================================
                    continuous_requests = []
                    for name in list(unassigned_staff):
                        if day_idx > 0:
                            prev_shift = str(df_shift.loc[df_shift['姓名'] == name, date_columns[day_idx - 1]].values[0]).strip().upper()
                            # 只有昨天沒休假，才符合「連上」資格
                            if prev_shift != 'OFF':
                                prev_zone = str(df_result.loc[df_result['姓名'] == name, date_columns[day_idx - 1]].values[0]).strip()
                                if prev_zone and prev_zone not in ['OFF', 'L', 'X', 'NAN', '']:
                                    is_trainee = (name in train_s2 and prev_zone == "S2") or \
                                                 (name in train_t and prev_zone == "T") or \
                                                 (name in train_b1_r and prev_zone in ["B1", "R", "S", "C2"])
                                    is_TP = prev_zone in ["T", "P"]
                                    # 🌟 更新：組別 F 也列入強迫連上的對象
                                    is_team_F = (get_team_of(name, shift_staff) == 'F')
                                    
                                    if is_trainee or is_TP or is_team_F:
                                        continuous_requests.append({'name': name, 'zone': prev_zone, 'is_trainee': is_trainee})
                    
                    # 讓訓練生優先搶走連續名額
                    continuous_requests.sort(key=lambda x: not x['is_trainee']) 
                    for req in continuous_requests:
                        name = req['name']
                        z = req['zone']
                        if name in unassigned_staff and z in available_zones:
                            assignments[name] = z
                            unassigned_staff.remove(name)
                            available_zones.remove(z)

                    # =======================================================
                    # 🌟 4. 新訓啟動與平均分配 (防撞與歸建機制)
                    # =======================================================
                    # (A) S2 訓練
                    s2_candidates = [n for n in list(unassigned_staff) if n in train_s2]
                    if s2_candidates:
                        s2_candidates.sort(key=lambda n: monthly_counts[n].get("S2", 0))
                        if "S2" in available_zones:
                            chosen = s2_candidates[0]
                            assignments[chosen] = "S2"
                            unassigned_staff.remove(chosen)
                            available_zones.remove("S2")

                    # (B) T 訓練
                    t_candidates = [n for n in list(unassigned_staff) if n in train_t]
                    if t_candidates:
                        t_candidates.sort(key=lambda n: monthly_counts[n].get("T", 0))
                        if "T" in available_zones:
                            chosen = t_candidates[0]
                            assignments[chosen] = "T"
                            unassigned_staff.remove(chosen)
                            available_zones.remove("T")

                    # (C) B1/R/C2/S 訓練
                    brcs_candidates = [n for n in list(unassigned_staff) if n in train_b1_r]
                    for name in brcs_candidates:
                        if name not in unassigned_staff: continue
                        valid_train_zones = [z for z in ["B1", "R", "S", "C2"] if z in available_zones]
                        if valid_train_zones:
                            valid_train_zones.sort(key=lambda z: (monthly_counts[name].get(z, 0), ["B1", "R", "S", "C2"].index(z)))
                            chosen = valid_train_zones[0]
                            assignments[name] = chosen
                            unassigned_staff.remove(name)
                            available_zones.remove(chosen)

                    # =======================================================
                    # 5. 連啟倫獨立鎖定
                    if "N連啟倫" in unassigned_staff:
                        chosen = "MO" if monthly_counts["N連啟倫"].get("MO",0) <= monthly_counts["N連啟倫"].get("MO1",0) else "MO1"
                        if chosen in available_zones:
                            assignments["N連啟倫"] = chosen
                            unassigned_staff.remove("N連啟倫")
                            available_zones.remove(chosen)
                    
                    # =======================================================
                    # 🌟 6. F組與E組的優先選區 (適用於剛放完 OFF 回來，要重新分配戰區的人)
                    
                    # (A) F 組優先拿 A1, R2, MO, MO1, S1
                    team_f_staff = [n for n in list(unassigned_staff) if get_team_of(n, shift_staff) == 'F']
                    team_f_priority_zones = ["A1", "R2", "MO", "MO1", "S1"]
                    for name in team_f_staff:
                        f_macros, f_severe = get_forbidden_zones(name, day_idx, df_result, date_columns)
                        avail_for_f = []
                        for z in team_f_priority_zones:
                            if z in available_zones:
                                # 確認這個區域不違反「跳過2天大區」和「重症不連上」的規範
                                if get_macro(z) not in f_macros and not (f_severe and is_severe(z)):
                                    avail_for_f.append(z)
                        
                        if not avail_for_f:
                            avail_for_f = [z for z in team_f_priority_zones if z in available_zones]

                        if avail_for_f:
                            avail_for_f.sort(key=lambda z: get_zone_count(z, monthly_counts, name))
                            min_count = get_zone_count(avail_for_f[0], monthly_counts, name)
                            lowest_zones = [z for z in avail_for_f if get_zone_count(z, monthly_counts, name) == min_count]
                            chosen_zone = random.choice(lowest_zones)
                            assignments[name] = chosen_zone
                            unassigned_staff.remove(name)
                            available_zones.remove(chosen_zone)

                    # (B) E 組優先拿 A2
                    team_e_staff = [n for n in list(unassigned_staff) if get_team_of(n, shift_staff) == 'E']
                    for name in team_e_staff:
                        if "A2" in available_zones:
                            f_macros, f_severe = get_forbidden_zones(name, day_idx, df_result, date_columns)
                            if get_macro("A2") not in f_macros:
                                assignments[name] = "A2"
                                unassigned_staff.remove(name)
                                available_zones.remove("A2")

                    # =======================================================
                    # 🌟 7. 一般歸建人員的終極防撞分發
                    # =======================================================
                    random.shuffle(unassigned_staff)
                    unassigned_staff.sort(key=lambda x: get_team_priority(x, shift_staff))

                    for name in list(unassigned_staff):
                        gender_series = shift_staff[shift_staff['姓名'] == name]['性別'].values
                        gender = gender_series[0] if len(gender_series) > 0 else ""
                        team = get_team_of(name, shift_staff)
                        
                        f_macros, f_severe = get_forbidden_zones(name, day_idx, df_result, date_columns)
                        team_allowed = team_allowed_zones.get(team, available_zones)

                        candidate_zones = []
                        for zone in list(available_zones):
                            if zone not in team_allowed: continue
                            if zone == "S2" and str(gender).upper() == "M": continue # 男性不上S2
                            if get_macro(zone) in f_macros: continue # 防撞：跳過兩天大區
                            if f_severe and is_severe(zone): continue # 防撞：重症不上兩天
                            candidate_zones.append(zone)

                        # 如果防撞規則太嚴格導致無位可去，則階段性放寬條件
                        if not candidate_zones:
                            for zone in list(available_zones):
                                if zone not in team_allowed: continue
                                if zone == "S2" and str(gender).upper() == "M": continue
                                if get_macro(zone) in f_macros: continue 
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

                        # 篩選 Max Days 上限
                        valid_candidates = []
                        for zone in candidate_zones:
                            if zone in MAX_DAYS and monthly_counts[name].get(zone, 0) >= MAX_DAYS[zone]:
                                continue
                            valid_candidates.append(zone)
                        if valid_candidates:
                            candidate_zones = valid_candidates

                        # 挑選本月站過最少次的區域以達到絕對平均
                        candidate_zones.sort(key=lambda z: get_zone_count(z, monthly_counts, name))
                        min_count = get_zone_count(candidate_zones[0], monthly_counts, name)
                        lowest_zones = [z for z in candidate_zones if get_zone_count(z, monthly_counts, name) == min_count]
                        chosen_zone = random.choice(lowest_zones)

                        assignments[name] = chosen_zone
                        unassigned_staff.remove(name)
                        if chosen_zone in available_zones:
                            available_zones.remove(chosen_zone)

                    # 將今日排定結果寫入歷史追蹤表
                    for name, assigned_zone in assignments.items():
                        df_result.loc[df_result['姓名'] == name, date_col] = assigned_zone
                        monthly_counts[name][assigned_zone] = monthly_counts[name].get(assigned_zone, 0) + 1

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

            st.success("🎉 終極排班運算完成！已啟動「6組戰區部署」、「F組連續霸王條款」與「重症防過勞」機制。")
            st.dataframe(df_result.head(10))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='排班結果')
                df_shift.to_excel(writer, index=False, sheet_name='原始班表')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 下載最終排班表 (Excel)", 
                data=excel_data, 
                file_name="排班結果_6組陣型版.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"執行時發生內部錯誤：{e}")
