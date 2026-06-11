import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="急診自動排班系統", layout="wide")
st.title("🏥 急診護理人員自動排班系統 (全功能整合版)")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    training_file = st.file_uploader("📂 1. 上傳班表 (training 檔案)", type=['xlsx', 'csv'])
with col2:
    template_file = st.file_uploader("📂 2. 上傳空白檔", type=['xlsx', 'csv'])

if training_file and template_file:
    try:
        # 讀取檔案
        df_shift = pd.read_csv(training_file) if 'csv' in training_file.name.lower() else pd.read_excel(training_file)
        df_template = pd.read_csv(template_file) if 'csv' in template_file.name.lower() else pd.read_excel(template_file)
            
        # 重新命名基礎欄位
        base_columns = {df_shift.columns[1]: '組別', df_shift.columns[2]: '性別', df_shift.columns[3]: '姓名'}
        df_shift = df_shift.rename(columns=base_columns)
        df_template = df_template.rename(columns=base_columns)
        
        # 🧹 資料清洗：過濾空白姓名、去除姓名前後的隱形空白鍵
        df_shift = df_shift.dropna(subset=['姓名'])
        df_template = df_template.dropna(subset=['姓名'])
        df_shift['姓名'] = df_shift['姓名'].astype(str).str.strip()
        df_template['姓名'] = df_template['姓名'].astype(str).str.strip()
        
        all_staff = df_shift['姓名'].unique().tolist()
        date_columns = df_shift.columns[4:]
        
        st.success("✅ 檔案讀取成功！請至左側邊欄設定本月訓練與組長規則。")
        
        # ===== 左側邊欄動態設定 =====
        st.sidebar.header("⚙️ 本月特殊排班規則設定")
        st.sidebar.subheader("🎓 訓練人員綁定 (整月固定區)")
        train_s2 = st.sidebar.multiselect("S2 訓練名單", all_staff)
        train_b1_r = st.sidebar.multiselect("B1/R/C2/S 訓練名單", all_staff)
        train_t = st.sidebar.multiselect("檢傷(T) 訓練名單", all_staff)
        
        st.sidebar.markdown("---")
        st.sidebar.subheader("👑 當月組長 (L) 第一順位設定")
        l_chain_1 = st.sidebar.selectbox("第一組首選", ["AHN黃麗婷"] + all_staff)
        l_chain_2 = st.sidebar.selectbox("第二組首選", ["AHN蕭惠澤"] + all_staff)
        l_chain_3 = st.sidebar.selectbox("第三組首選", ["N3許慧芳"] + all_staff)
        
        # ===== 排班核心運算 =====
        st.markdown("---")
        if st.button("🚀 開始自動排班運算 (套用上述規則)"):
            with st.spinner("🧠 正在執行邏輯運算中..."):
                df_result = df_template.copy()
                tp_tracker = {name: None for name in all_staff}
                progress_bar = st.progress(0)
                
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
                        
                        # 規則 2-6: 開區邏輯
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
                            # 補上 B1/R/C2/S 訓練名單邏輯
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
                        
                        # 4. L 順位指派 (讀取 UI 設定)
                        l_chains = [
                            [l_chain_1, "N3尤美惠", "N3許嘉文"],
                            [l_chain_2, "N2李曉侖", "N3黃義全"],
                            [l_chain_3, "N2江品儒", "N1許家瑄"]
                        ]
                        for chain in l_chains:
                            for l_candidate in chain:
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
                        
                        # 6. 剩餘分配 (避開男性去 S2，並加入安全保護)
                        for name in list(unassigned_staff):
                            # 加上防呆，確保能安全抓到性別
                            gender_series = shift_staff[shift_staff['姓名'] == name]['性別'].values
                            gender = gender_series[0] if len(gender_series) > 0 else ""
                            
                            for zone in list(available_zones):
                                if zone == "S2" and str(gender).upper() == "M": continue
                                assignments[name] = zone
                                unassigned_staff.remove(name)
                                available_zones.remove(zone)
                                if zone in ["T", "P"]: tp_tracker[name] = zone
                                break

                        # 寫回結果
                        for name, assigned_zone in assignments.items():
                            df_result.loc[df_result['姓名'] == name, date_col] = assigned_zone

                    # 解除 OFF 人員的 TP 狀態
                    todays_off_staff = df_shift[df_shift[date_col].astype(str).str.upper() == 'OFF']['姓名'].tolist()
                    for off_name in todays_off_staff:
                        tp_tracker[off_name] = None
                        
                    progress_bar.progress((day_idx + 1) / len(date_columns))
                
                st.success("🎉 排班運算完成！")
                st.dataframe(df_result.head(10))
                csv = df_result.to_csv(index=False).encode('utf-8-sig')
                st.download_button(label="📥 下載最終排班表 (CSV)", data=csv, file_name="排班結果_測試版.csv", mime="text/csv")
                
    except Exception as e:
        st.error(f"執行時發生錯誤：{e}")
else:
    st.info("💡 請上傳檔案以解鎖規則設定介面。")
