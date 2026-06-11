import streamlit as st
import pandas as pd
import numpy as np
import random
import io  # 新增：用來將資料轉換成 Excel 檔案格式的套件

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

# 初始化變數
all_staff = []
df_shift = None
df_template = None
data_ready = False
date_columns = []

if training_file and template_file:
    try:
        df_shift = pd.read_csv(training_file) if 'csv' in training_file.name.lower() else pd.read_excel(training_file)
        df_template = pd.read_csv(template_file) if 'csv' in template_file.name.lower() else pd.read_excel(template_file)
            
        # 強制將 df_shift 的第 2, 3, 4 欄重新命名
        shift_cols = list(df_shift.columns)
        shift_cols[1], shift_cols[2], shift_cols[3] = '組別', '性別', '姓名'
        df_shift.columns = shift_cols
        
        # 強制將 df_template 的第 2, 3, 4 欄重新命名
        template_cols = list(df_template.columns)
        template_cols[1], template_cols[2], template_cols[3] = '組別', '性別', '姓名'
        df_template.columns = template_cols
        
        # 資料清洗
        df_shift = df_shift.dropna(subset=['姓名'])
        df_template = df_template.dropna(subset=['姓名'])
        df_shift['姓名'] = df_shift['姓名'].astype(str).str.strip()
        df_template['姓名'] = df_template['姓名'].astype(str).str.strip()
        
        # 取得人員名單與日期
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
    with st.spinner("🧠 正在執行邏輯運算與區域輪替中..."):
        try:
            df_result = df_template.copy()
            tp_tracker = {name: None for name in all_staff}
            history_tracker = {name: [] for name in all_staff}
            progress_bar = st.progress(0)
            
            # 定義區域群組
            zone_groups = {
                "MO": "MO_G", "MO1": "MO_G", "MO2": "MO_G",
                "S": "S_G", "S1": "S_G", "S2": "S_G",
                "A1": "ABC_G", "B1": "ABC_G", "C1": "ABC_G", "A2": "ABC_G", "B2": "ABC_G", "C2": "ABC_G",
                "R": "R_G", "R1": "R_G", "R2": "R_G", "R3": "R_G"
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
                    if shift_type == 'D':
                        l_chain = [d_l_1, d_l_2, d_l_3]
                    elif shift_type == 'E':
                        l_chain = [e_l_1, e_l_2, e_l_3]
                    elif shift_type == 'N':
                        l_chain = [n_l_1, n_l_2, n_l_3]
                    else:
                        l_chain = []

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
                    
                    # 6. 剩餘分配 (加入歷史輪替與隨機洗牌機制)
                    random.shuffle(unassigned_staff) 
                    
                    for name in list(unassigned_staff):
                        gender_series = shift_staff[shift_staff['姓名'] == name]['性別'].values
                        gender = gender_series[0] if len(gender_series) > 0 else ""
                        
                        recent_groups = history_tracker[name][-2:] if name in history_tracker else []
                        
                        candidate_zones = []
                        for zone in list(available_zones):
                            if zone == "S2" and str(gender).upper() == "M": continue
                            candidate_zones.append(zone)
                            
                        if not candidate_zones: continue 
                        
                        best_zones = [z for z in candidate_zones if zone_groups.get(z, z) not in recent_groups]
                        
                        if best_zones:
                            chosen_zone = random.choice(best_zones)
                        else:
                            last_group = history_tracker[name][-1:] if name in history_tracker else []
                            better_zones = [z for z in candidate_zones if zone_groups.get(z, z) not in last_group]
                            if better_zones:
                                chosen_zone = random.choice(better_zones)
                            else:
                                chosen_zone = random.choice(candidate_zones)

                        assignments[name] = chosen_zone
                        unassigned_staff.remove(name)
                        available_zones.remove(chosen_zone)
                        if chosen_zone in ["T", "P"]: tp_tracker[name] = chosen_zone

                    # 寫回結果與歷史記憶
                    for name, assigned_zone in assignments.items():
                        df_result.loc[df_result['姓名'] == name, date_col] = assigned_zone
                        group = zone_groups.get(assigned_zone, assigned_zone)
                        history_tracker[name].append(group)

                # 解除 OFF 人員的 TP 狀態
                todays_off_staff = df_shift[df_shift[date_col].astype(str).str.upper() == 'OFF']['姓名'].tolist()
                for off_name in todays_off_staff:
                    tp_tracker[off_name] = None
                    
                progress_bar.progress((day_idx + 1) / len(date_columns))
            
            st.success("🎉 排班運算完成！區域已根據記憶進行輪替。")
            st.dataframe(df_result.head(10))
            
            # 🌟 新增：將產出的結果轉為 Excel 檔案供下載
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='排班結果')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 下載最終排班表 (Excel)", 
                data=excel_data, 
                file_name="排班結果_輪替版.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"執行時發生內部錯誤：{e}")
