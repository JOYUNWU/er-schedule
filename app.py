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
    with st.spinner("🧠 啟動各組別專屬守備區域與 D 組優先排序邏輯..."):
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
            
            # 🌟 新增：各組別專屬可上班區域 (白名單)
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
                    
                    # 1. L 順位指派
                    if shift_type == 'D': l_chain = [d_l_1, d_l_2, d_l_3]
                    elif shift_type == 'E': l_chain = [e_l_1, e_l_2, e_l_3]
                    elif shift_type == 'N': l_chain = [n_l_1, n_l_2, n_l_3]
                    else: l_chain = []

                    for l_candidate in l_chain:
                        if l_candidate in unassigned_staff:
                            assignments[l_candidate] = "L"
                            unassigned_staff.remove(l_candidate)
                            break
                    
                    # 動態滿編邏輯
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
                    
                    while len(available_zones) < len(unassigned_staff):
                        available_zones.append(f"支援{len(available_zones)+1}")
                    if len(available_zones) > len(unassigned_staff):
                        available_zones = available_zones[:len(unassigned_staff)]
                    
                    # 2. 不動檔優先
                    for _, row in shift_staff.iterrows():
                        name, preset = row['姓名'], str(row['預設區域']).strip()
                        if preset.upper() not in ['X', 'NAN', 'NONE', ''] and name in unassigned_staff:
                            assignments[name] = preset
                            unassigned_staff.remove(name)
                            if preset in available_zones: available_zones.remove(preset)
                            elif available_zones: available_zones.pop()
                    
                    # 3. 訓練名單優先綁定
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
                            
                    # 4. T/P 連續狀態處理
                    for name in list(unassigned_staff):
                        prev_tp = tp_tracker.get(name)
