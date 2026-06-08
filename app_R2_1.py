import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import glob

# ==============================================================================
# [릴리즈 버전 - app_R2.1] 현장 장비 인터페이스 스펙 맞춤 반영 버전
# ==============================================================================

st.set_page_config(
    page_title="PBA 조립 공정 관리 시스템 (V2.1)",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 💾 3대 데이터베이스 파일 경로 정의
MAIN_DB = "pba_main_process.csv"        # 인두기(1-25번), RS-232 상온방치 로그
MASK_DB = "pba_metal_mask.csv"         # 메탈마스크 모델 직접 입력 및 텐션 5포인트
VISC_DB = "pba_solder_viscosity.csv"   # 솔더크림 Lot별 5회 측정 및 평균

# ------------------------------------------------------------------------------
# 📂 데이터베이스 파일 초기화 및 로드 함수들
# ------------------------------------------------------------------------------
def load_main_db():
    if not os.path.exists(MAIN_DB):
        # 데모 초기 이력 생성
        dates = [(datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(15)][::-1]
        df = pd.DataFrame({
            'Date': dates,
            'Model_Name': ['SMT-PRIME-01', 'SMT-NEXUS-A', 'SMT-PRIME-02']*5,
            'Solder_Start_Time': ['08:30']*15,
            'Solder_End_Time': ['10:40']*15,
            'Solder_Leave_Min': [130]*15, # 자동 계산 결과 (분)
            'Iron_No': [f"{i}번" for i in np.random.randint(1, 26, size=15)], # 1~25번 인두기 지정
            'Iron_Temp': np.random.uniform(345.0, 362.0, size=15).round(1),
            'Iron_Leak_Volt': np.random.uniform(0.5, 1.8, size=15).round(2),
            'Iron_Resistance': np.random.uniform(1.0, 4.2, size=15).round(2)
        })
        df.to_csv(MAIN_DB, index=False, encoding='utf-8-sig')
    return pd.read_csv(MAIN_DB)

def load_mask_db():
    if not os.path.exists(MASK_DB):
        dates = [(datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(10)][::-1]
        df = pd.DataFrame({
            'Date': dates,
            'Model_Name': ['SMT-PRIME-01', 'SMT-NEXUS-A', 'SMT-PRIME-01', 'SMT-CORE-X', 'SMT-NEXUS-A']*2,
            'Daily_Count': np.random.randint(1000, 2500, size=10),
            'P1_Center': np.random.uniform(0.50, 0.85, size=10).round(2),
            'P2_TopLeft': np.random.uniform(0.50, 0.85, size=10).round(2),
            'P3_TopRight': np.random.uniform(0.50, 0.85, size=10).round(2),
            'P4_BotLeft': np.random.uniform(0.50, 0.85, size=10).round(2),
            'P5_BotRight': np.random.uniform(0.50, 0.85, size=10).round(2)
        })
        df.to_csv(MASK_DB, index=False, encoding='utf-8-sig')
    return pd.read_csv(MASK_DB)

def load_visc_db():
    if not os.path.exists(VISC_DB):
        df = pd.DataFrame({
            'Date': [(datetime.now() - timedelta(days=3)).strftime('%Y-%m-%d'), (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')],
            'Lot_No': ['LOT260601', 'LOT260605'],
            'M1': [180, 210], 'M2': [195, 205], 'M3': [210, 220], 'M4': [185, 215], 'M5': [200, 230],
            'Average': [194.0, 216.0]
        })
        df.to_csv(VISC_DB, index=False, encoding='utf-8-sig')
    return pd.read_csv(VISC_DB)

# 데이터베이스 가동
df_main = load_main_db()
df_mask = load_mask_db()
df_visc = load_visc_db()

# ------------------------------------------------------------------------------
# 🧊 냉장고 온도 설비 PC 일별 파일 매핑 파싱 함수
# ------------------------------------------------------------------------------
def load_daily_fridge_data():
    network_folder = r"\\192.168.1.10\fridge_data"
    local_backup_folder = r"C:\fridge_data"
    target_folder = network_folder if os.path.exists(network_folder) else local_backup_folder
    
    today_str = datetime.now().strftime('%Y-%m-%d')
    
    if not os.path.exists(target_folder):
        times = [datetime.now() - timedelta(minutes=30*i) for i in range(24)]
        df_demo = pd.DataFrame({'DateTime': times, 'Temp': np.random.uniform(4.5, 6.8, size=24).round(1)})
        df_demo['Date'] = df_demo['DateTime'].dt.strftime('%H:%M')
        return df_demo.sort_values('DateTime'), "⚠️ [데모 가동] 설비 PC 네트워크 공유 폴더를 찾을 수 없습니다."
    
    try:
        today_file = os.path.join(target_folder, f"{today_str}.xlsx")
        if not os.path.exists(today_file):
            today_file = os.path.join(target_folder, f"{today_str}.xls")
            
        if not os.path.exists(today_file):
            excel_files = glob.glob(os.path.join(target_folder, "*.xls*"))
            if not excel_files:
                return None, "⚠️ 공유 폴더 내에 읽을 수 있는 엑셀 파일이 없습니다."
            chosen_file = max(excel_files, key=os.path.getmtime)
        else:
            chosen_file = today_file
            
        df = pd.read_excel(chosen_file)
        df.columns = ['DateTime', 'Temp']
        df['DateTime'] = pd.to_datetime(df['DateTime'])
        df = df.sort_values('DateTime')
        df['Date'] = df['DateTime'].dt.strftime('%H:%M')
        return df, f"🔄 냉장고 설비 파일 연동 성공 ({os.path.basename(chosen_file)})"
    except Exception as e:
        return None, f"❌ 냉장고 데이터 파싱 실패: {str(e)}"

# ==============================================================================
# 🏢 대시보드 상단 타이틀 및 KPI 현황판
# ==============================================================================
st.title("🏭 PBA 조립 공정 품질 모니터링 시스템 (Release app_R2.1)")
st.markdown("설비 RS-232/파일 이력 추적 및 공정 요소별 영구 저장·수정 통합 대시보드")
st.divider()

# 최신 KPI 값 수집
df_fridge, fridge_msg = load_daily_fridge_data()
current_fridge_temp = df_fridge['Temp'].iloc[-1] if df_fridge is not None and len(df_fridge)>0 else 0.0
latest_iron_no = df_main['Iron_No'].iloc[-1] if len(df_main)>0 else "N/A"
latest_iron_temp = df_main['Iron_Temp'].iloc[-1] if len(df_main)>0 else 0.0
latest_mask_model = df_mask['Model_Name'].iloc[-1] if len(df_mask)>0 else "없음"

# 전체 모델별 누적 마스크 타수 계산
mask_totals = df_mask.groupby('Model_Name')['Daily_Count'].sum().to_dict()
latest_mask_total = mask_totals.get(latest_mask_model, 0)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
with kpi1:
    f_status = "정상" if 3.0 <= current_fridge_temp <= 10.0 else "🚨 이상"
    st.metric("🧊 냉장고 현재 온도", f"{current_fridge_temp} ℃", delta=f_status, delta_color="normal" if f_status=="정상" else "inverse")
with kpi2:
    st.metric(f"🔥 최근 인두기 온도 ({latest_iron_no})", f"{latest_iron_temp} ℃", "관리기준: 350±10℃")
with kpi3:
    st.metric(f"📐 마스크 [{latest_mask_model}] 누적타수", f"{latest_mask_total:,} 회", "현장 직접 입력 모델")
with kpi4:
    latest_visc_avg = df_visc['Average'].iloc[-1] if len(df_visc)>0 else 0.0
    st.metric("🧪 솔더크림 최근 평균점도", f"{latest_visc_avg} Pa·s", "M8 기준 (70~300)")

st.divider()

# ==============================================================================
# 🎛️ 사이드바 - 현장 스펙 맞춤 데이터 입력 / 수정 센터
# ==============================================================================
st.sidebar.header("🎛️ 데이터 입력 및 이력 수정 센터")
input_category = st.sidebar.radio("작업 선택", ["1. 인두기 & 상온방치 관리", "2. 메탈마스크 & 텐션 관리", "3. 솔더 Lot별 점도 관리"])

# --- 사이드바 기능 1: 인두기 (1-25번) 및 RS-232 상온방치 로그 관리 ---
if input_category == "1. 인두기 & 상온방치 관리":
    st.sidebar.subheader("✍️ 인두기 온도 및 상온방치 등록")
    
    target_date = st.sidebar.date_input("점검/방치 일자 선택", datetime.now())
    date_str = target_date.strftime('%Y-%m-%d')
    
    # 1-25번 중 관리할 인두기 번호 지정
    sb_iron_no = st.sidebar.selectbox("🎯 인두기 번호 지정", [f"{i}번" for i in range(1, 26)])
    
    # 수정 유도 기능: 기존 등록 데이터 매핑 확인 (날짜 + 인두기 번호 기준)
    existing_row = df_main[(df_main['Date'] == date_str) & (df_main['Iron_No'] == sb_iron_no)]
    is_edit = not existing_row.empty
    
    default_model = existing_row['Model_Name'].iloc[0] if is_edit else ""
    default_start = existing_row['Solder_Start_Time'].iloc[0] if is_edit else "08:00"
    default_end = existing_row['Solder_End_Time'].iloc[0] if is_edit else "10:15"
    default_temp = float(existing_row['Iron_Temp'].iloc[0]) if is_edit else 350.0
    default_volt = float(existing_row['Iron_Leak_Volt'].iloc[0]) if is_edit else 1.2
    default_res = float(existing_row['Iron_Resistance'].iloc[0]) if is_edit else 2.1
    
    if is_edit:
        st.sidebar.info(f"🔄 [수정 모드] {date_str} / {sb_iron_no}의 기존 데이터가 교체됩니다.")
    else:
        st.sidebar.success(f"🆕 [신규 등록] {date_str} / {sb_iron_no}의 데이터를 새로 생성합니다.")

    with st.sidebar.form("main_form"):
        # 모델명 직접 입력 요구사항 반영
        sb_model = st.text_input("생산 모델명 직접 입력", value=default_model, placeholder="예: SMT-NEXUS-01").strip().upper()
        
        st.markdown("**🔌 RS-232 통신 설비 데이터**")
        sb_start = st.text_input("방치 시작 시간 (HH:MM)", value=default_start)
        sb_end = st.text_input("방치 종료 시간 (HH:MM)", value=default_end)
        
        st.markdown(f"**🔥 {sb_iron_no} 인두기 측정 데이터**")
        sb_itemp = st.number_input("인두기 팁 온도 (℃)", min_value=300.0, max_value=450.0, value=default_temp, step=1.0)
        sb_ivolt = st.number_input("누설 전압 (mV)", min_value=0.0, value=default_volt, step=0.1)
        sb_ires = st.number_input("접지 저항 (Ω)", min_value=0.0, value=default_res, step=0.1)
        
        submit_main = st.form_submit_button("💾 데이터 저장/수정 반영")
        
    if submit_main:
        if not sb_model:
            st.sidebar.error("❌ 생산 모델명을 입력해 주세요.")
        else:
            # RS-232 시작/종료 시간에 따른 방치 분(Minute) 자동 계산
            try:
                t1 = datetime.strptime(sb_start, "%H:%M")
                t2 = datetime.strptime(sb_end, "%H:%M")
                leave_duration = int((t2 - t1).total_seconds() / 60)
                if leave_duration < 0:
                    leave_duration += 1440 # 익일 생산 케이스 보정
            except:
                st.sidebar.error("❌ 시간 입력 포맷(HH:MM)이 오동작했습니다. 기본값 120분으로 임시 대체 처리합니다.")
                leave_duration = 120

            if is_edit:
                # 기존 해당 행 탈락 처리 후 재생성(수정)
                df_main = df_main.drop(existing_row.index)
            
            new_data = pd.DataFrame([{
                'Date': date_str, 'Model_Name': sb_model, 
                'Solder_Start_Time': sb_start, 'Solder_End_Time': sb_end, 'Solder_Leave_Min': leave_duration,
                'Iron_No': sb_iron_no, 'Iron_Temp': round(sb_itemp, 1), 
                'Iron_Leak_Volt': round(sb_ivolt, 2), 'Iron_Resistance': round(sb_ires, 2)
            }])
            df_main = pd.concat([df_main, new_data], ignore_index=True).sort_values(['Date', 'Iron_No'])
            df_main.to_csv(MAIN_DB, index=False, encoding='utf-8-sig')
            st.sidebar.success("💾 인두기 및 RS-232 방치 이력이 보존되었습니다.")
            st.rerun()

# --- 사이드바 기능 2: 메탈마스크 모델 직접 입력 및 5포인트 텐션 관리 ---
elif input_category == "2. 메탈마스크 & 텐션 관리":
    st.sidebar.subheader("📐 마스크 모델 및 텐션 입력")
    target_date = st.sidebar.date_input("점검 일자 선택", datetime.now())
    date_str = target_date.strftime('%Y-%m-%d')
    
    # 모델명 자유 텍스트 직접 입력 방식 적용
    sm_model = st.sidebar.text_input("가동 메탈마스크 모델명 직접 입력", value="SMT-PRIME-01").strip().upper()
    
    existing_mask = df_mask[(df_mask['Date'] == date_str) & (df_mask['Model_Name'] == sm_model)]
    is_mask_edit = not existing_mask.empty
    
    d_count = int(existing_mask['Daily_Count'].iloc[0]) if is_mask_edit else 1500
    d_p1 = float(existing_mask['P1_Center'].iloc[0]) if is_mask_edit else 0.65
    d_p2 = float(existing_mask['P2_TopLeft'].iloc[0]) if is_mask_edit else 0.65
    d_p3 = float(existing_mask['P3_TopRight'].iloc[0]) if is_mask_edit else 0.65
    d_p4 = float(existing_mask['P4_BotLeft'].iloc[0]) if is_mask_edit else 0.65
    d_p5 = float(existing_mask['P5_BotRight'].iloc[0]) if is_mask_edit else 0.65
    
    if is_mask_edit:
        st.sidebar.info(f"🔄 [수정 모드] {date_str} / {sm_model} 마스크 정보를 수정합니다.")
        
    with st.sidebar.form("mask_form"):
        sm_count = st.number_input("당일 가동 타수 (회)", min_value=0, value=d_count, step=100)
        
        st.markdown("**📌 텐션 측정치 (스펙: 0.45 ~ 0.90 mm)**")
        mp1 = st.number_input("Point 1 (Center)", min_value=0.0, max_value=2.0, value=d_p1, step=0.01)
        mp2 = st.number_input("Point 2 (Top-Left)", min_value=0.0, max_value=2.0, value=d_p2, step=0.01)
        mp3 = st.number_input("Point 3 (Top-Right)", min_value=0.0, max_value=2.0, value=d_p3, step=0.01)
        mp4 = st.number_input("Point 4 (Bottom-Left)", min_value=0.0, max_value=2.0, value=d_p4, step=0.01)
        mp5 = st.number_input("Point 5 (Bottom-Right)", min_value=0.0, max_value=2.0, value=d_p5, step=0.01)
        
        submit_mask = st.form_submit_button("📐 마스크 정보 저장/수정")
        
    if submit_mask:
        if not sm_model:
            st.sidebar.error("❌ 마스크 모델명을 입력해 주세요.")
        else:
            if is_mask_edit:
                df_mask = df_mask.drop(existing_mask.index)
                
            new_mask_data = pd.DataFrame([{
                'Date': date_str, 'Model_Name': sm_model, 'Daily_Count': sm_count,
                'P1_Center': round(mp1, 2), 'P2_TopLeft': round(mp2, 2), 'P3_TopRight': round(mp3, 2),
                'P4_BotLeft': round(mp4, 2), 'P5_BotRight': round(mp5, 2)
            }])
            df_mask = pd.concat([df_mask, new_mask_data], ignore_index=True).sort_values('Date')
            df_mask.to_csv(MASK_DB, index=False, encoding='utf-8-sig')
            st.sidebar.success(f"💾 {sm_model} 모델 텐션/타수 데이터가 기록되었습니다.")
            st.rerun()

# --- 사이드바 기능 3: 솔더 Lot별 점도 관리 ---
elif input_category == "3. 솔더 Lot별 점도 관리":
    st.sidebar.subheader("🧪 AIM M8 솔더 점도 5회 측정")
    target_date = st.sidebar.date_input("측정 일자 선택", datetime.now())
    date_str = target_date.strftime('%Y-%m-%d')
    v_lot = st.sidebar.text_input("Solder 크림 Lot 번호 입력", value="LOT260601").strip().upper()
    
    existing_visc = df_visc[(df_visc['Date'] == date_str) & (df_visc['Lot_No'] == v_lot)]
    is_v_edit = not existing_visc.empty
    
    dv1 = int(existing_visc['M1'].iloc[0]) if is_v_edit else 190
    dv2 = int(existing_visc['M2'].iloc[0]) if is_v_edit else 200
    dv3 = int(existing_visc['M3'].iloc[0]) if is_v_edit else 205
    dv4 = int(existing_visc['M4'].iloc[0]) if is_v_edit else 195
    dv5 = int(existing_visc['M5'].iloc[0]) if is_v_edit else 210
    
    if is_v_edit:
        st.sidebar.info(f"🔄 [수정 모드] {v_lot}의 기존 측정치를 수정합니다.")
        
    with st.sidebar.form("visc_form"):
        st.markdown("500g Jar 기준 Lot별 5회차 데이터")
        m1 = st.number_input("1회차 측정값 (Pa·s)", min_value=0, value=dv1)
        m2 = st.number_input("2회차 측정값 (Pa·s)", min_value=0, value=dv2)
        m3 = st.number_input("3회차 측정값 (Pa·s)", min_value=0, value=dv3)
        m4 = st.number_input("4회차 측정값 (Pa·s)", min_value=0, value=dv4)
        m5 = st.number_input("5회차 측정값 (Pa·s)", min_value=0, value=dv5)
        
        submit_visc = st.form_submit_button("🧪 점도 데이터 저장/수정")
        
    if submit_visc and v_lot:
        if is_v_edit:
            df_visc = df_visc.drop(existing_visc.index)
            
        v_avg = round(np.mean([m1, m2, m3, m4, m5]), 1)
        new_visc_data = pd.DataFrame([{
            'Date': date_str, 'Lot_No': v_lot,
            'M1': m1, 'M2': m2, 'M3': m3, 'M4': m4, 'M5': m5, 'Average': v_avg
        }])
        df_visc = pd.concat([df_visc, new_visc_data], ignore_index=True).sort_values('Date')
        df_visc.to_csv(VISC_DB, index=False, encoding='utf-8-sig')
        st.sidebar.success(f"💾 평균 {v_avg} Pa·s 연산 완료 및 로컬 저장 성공.")
        st.rerun()

# ==============================================================================
# 📊 메인 화면 : 공정 탭별 세부 현황 대시보드
# ==============================================================================
tab_fridge, tab_solder, tab_mask, tab_iron, tab_printer = st.tabs([
    "🧊 1. 냉장고 온도 제어", 
    "⏳ 2. 솔더크림 (RS-232 상온방치/점도)", 
    "📐 3. 메탈 마스크 (직접입력 모델/5P 텐션)", 
    "⚡ 4. 인두기 검사 (1번~25번 관리)",
    "⚙️ 5. 스크린 프린터 (설비 예정)"
])

# ------------------------------------------------------------------------------
# [TAB 1] 냉장고 온도 제어 섹션
# ------------------------------------------------------------------------------
with tab_fridge:
    st.subheader("🧊 보관 냉장고 일별 생성 로그 실시간 추적")
    st.info(fridge_msg)
    
    if df_fridge is not None and len(df_fridge) > 0:
        fig_f = px.line(df_fridge, x='Date', y='Temp', title="당일 생성된 파일 내 실시간 온도 로깅 추이", markers=True)
        fig_f.add_hrect(y0=3.0, y1=10.0, fillcolor="green", opacity=0.08, line_width=0, annotation_text="정상 범위 (3~10℃)")
        fig_f.update_layout(xaxis_title="추출 수집 시간", yaxis_title="온도 (℃)")
        st.plotly_chart(fig_f, width="stretch")
        st.dataframe(df_fridge.tail(6), width="stretch")

# ------------------------------------------------------------------------------
# [TAB 2] 솔더크림 상온방치(RS-232) 및 점도 관리 섹션
# ------------------------------------------------------------------------------
with tab_solder:
    col_s1, col_s2 = st.columns(2)
    
    with col_s1:
        st.subheader("⏳ RS-232 통신 이력 기반 상온방치 이력 관리")
        st.caption("통신 수집 데이터 스펙 [방치 일자, 시작 시간, 종료 시간] -> 총 방치 시간 도출 (2시간 스펙 검증)")
        
        # 준수 상태 조건 매핑
        df_main['Status'] = df_main['Solder_Leave_Min'].apply(lambda x: 'OK (2시간이상)' if x >= 120 else 'NG (시간미달)')
        
        fig_s1 = px.bar(df_main, x='Date', y='Solder_Leave_Min', color='Status',
                        hover_data=['Model_Name', 'Solder_Start_Time', 'Solder_End_Time'],
                        color_discrete_map={'OK (2시간이상)': '#2ECC71', 'NG (시간미달)': '#E74C3C'},
                        text_auto=True, title="일자별 설비 연동 방치 수행 시간 (분)")
        fig_s1.add_hline(y=120, line_dash="dash", line_color="black", annotation_text="스펙 한계선 (120분)")
        st.plotly_chart(fig_s1, width="stretch")
        
    with col_s2:
        st.subheader("🧪 AIM M8 제품 Lot별 점도 관리 (70-300 Pa·s)")
        st.caption("500g Jar 기준 개별 5회 측정 분포 및 연산된 평균 트렌드")
        
        if len(df_visc) > 0:
            fig_s2 = go.Figure()
            fig_s2.add_trace(go.Scatter(x=df_visc['Lot_No'], y=df_visc['Average'], mode='lines+markers', name='Lot별 평균값', line=dict(color='blue', width=3)))
            
            # 5회차 측정 산포 데이터 바인딩
            for idx, r in df_visc.iterrows():
                fig_s2.add_trace(go.Scatter(x=[r['Lot_No']]*5, y=[r['M1'], r['M2'], r['M3'], r['M4'], r['M5']], mode='markers', marker=dict(color='gray', opacity=0.6), showlegend=False))
            
            fig_s2.add_hrect(y0=70, y1=300, fillcolor="blue", opacity=0.05, line_width=0, annotation_text="M8 합격 스펙 (70~300 Pa·s)")
            fig_s2.update_layout(title="Lot별 5회 측정 산포 및 평균 트렌드", xaxis_title="Lot 번호", yaxis_title="점도 (Pa·s)")
            st.plotly_chart(fig_s2, width="stretch")
            
    st.markdown("#### 📋 솔더크림 품질 기록 데이터베이스")
    st.dataframe(df_visc, width="stretch")

# ------------------------------------------------------------------------------
# [TAB 3] 메탈 마스크 타수 및 텐션 관리 섹션 (직접 입력 모델 대응)
# ------------------------------------------------------------------------------
with tab_mask:
    st.subheader("📐 메탈 마스크 직접 입력 모델별 당일 사용량 및 총 누적 가동 타수")
    
    # 자유롭게 입력된 모델명 기반 동적 그룹 집계
    df_mask_agg = df_mask.groupby('Model_Name').agg(
        최근작업일=('Date', 'max'),
        최근작업타수=('Daily_Count', 'last')
    ).reset_index()
    df_mask_agg['역대_총_누적_타수'] = df_mask_agg['Model_Name'].map(mask_totals)
    
    col_m1, col_m2 = st.columns([1, 2])
    with col_m1:
        st.markdown("**📊 직접 입력 마스크 자산별 집계 스냅샷**")
        st.dataframe(df_mask_agg, width="stretch")
    with col_m2:
        fig_m1 = px.bar(df_mask, x='Date', y='Daily_Count', color='Model_Name', title="일자별/모델별 마스크 실적 분포", text_auto=True)
        st.plotly_chart(fig_m1, width="stretch")
        
    st.divider()
    st.subheader("📌 메탈마스크 5개 포인트(Center/외곽) 텐션 모니터링")
    st.caption("품질 관리 스펙 한계 범위: 0.45mm - 0.90mm")
    
    fig_m2 = go.Figure()
    x_labels = df_mask['Date'] + " (" + df_mask['Model_Name'] + ")"
    fig_m2.add_trace(go.Scatter(x=x_labels, y=df_mask['P1_Center'], name='P1 Center', mode='lines+markers'))
    fig_m2.add_trace(go.Scatter(x=x_labels, y=df_mask['P2_TopLeft'], name='P2 Top-Left', mode='lines+markers'))
    fig_m2.add_trace(go.Scatter(x=x_labels, y=df_mask['P3_TopRight'], name='P3 Top-Right', mode='lines+markers'))
    fig_m2.add_trace(go.Scatter(x=x_labels, y=df_mask['P4_BotLeft'], name='P4 Bottom-Left', mode='lines+markers'))
    fig_m2.add_trace(go.Scatter(x=x_labels, y=df_mask['P5_BotRight'], name='P5 Bottom-Right', mode='lines+markers'))
    
    fig_m2.add_hrect(y0=0.45, y1=0.90, fillcolor="orange", opacity=0.07, line_width=0, annotation_text="텐션 스펙 범위 (0.45 ~ 0.90mm)")
    fig_m2.update_layout(title="마스크 측정 이력별 5개 포인트 텐션 변동 트렌드", xaxis_title="측정일 (모델명)", yaxis_title="텐션 데이터 (mm)")
    st.plotly_chart(fig_m2, width="stretch")

# ------------------------------------------------------------------------------
# [TAB 4] 인두기 정기 검사 섹션 (1~25번 선택적 시각화 및 이력 관리)
# ------------------------------------------------------------------------------
with tab_iron:
    st.subheader("⚡ 인두기 호기별 온도 및 정전기 방지 Factor 관리 (1번~25번)")
    
    # 사용 편의를 위한 메인 화면 내 필터 추가
    selected_iron_view = st.selectbox("🔍 시각화 필터링할 인두기 호기 선택", [f"{i}번" for i in range(1, 26)])
    df_iron_filtered = df_main[df_main['Iron_No'] == selected_iron_view]
    
    if len(df_iron_filtered) > 0:
        col_i1, col_i2, col_i3 = st.columns(3)
        with col_i1:
            fig_i1 = px.line(df_iron_filtered, x='Date', y='Iron_Temp', title=f"🌡️ {selected_iron_view} 팁 온도 트렌드 (350±10℃)", markers=True, color_discrete_sequence=['#E67E22'])
            fig_i1.add_hrect(y0=340.0, y1=360.0, fillcolor="orange", opacity=0.1, annotation_text="정상범위")
            st.plotly_chart(fig_i1, width="stretch")
        with col_i2:
            fig_i2 = px.bar(df_iron_filtered, x='Date', y='Iron_Leak_Volt', title=f"⚡ {selected_iron_view} 누설 전압 (Spec: 2.0mV 이하)", color_discrete_sequence=['#9B59B6'])
            fig_i2.add_hline(y=2.0, line_dash="dash", line_color="red")
            st.plotly_chart(fig_i2, width="stretch")
        with col_i3:
            fig_i3 = px.line(df_iron_filtered, x='Date', y='Iron_Resistance', title=f"🔌 {selected_iron_view} 접지 저항 (Spec: 5.0Ω 이하)", markers=True, color_discrete_sequence=['#34495E'])
            fig_i3.add_hline(y=5.0, line_dash="dash", line_color="red")
            st.plotly_chart(fig_i3, width="stretch")
    else:
        st.warning(f"선택하신 {selected_iron_view} 인두기는 최근 등록된 이력이 데이터베이스에 없습니다. 사이드바에서 측정값을 입력해 주세요.")
        
    st.markdown("#### 📋 인두기 및 공정 종합 마스터 데이터 로그 (전체 호기 이력 통합뷰)")
    st.dataframe(df_main.sort_values('Date', ascending=False), width="stretch")

# ------------------------------------------------------------------------------
# [TAB 5] 스크린 프린터 공정 섹션 (유예 처리)
# ------------------------------------------------------------------------------
with tab_printer:
    st.subheader("⚙️ 스크린 프린터 챔버 환경 제어")
    st.warning("⚠️ 신규 설비 입고 후 온/습도 자동 인터페이스 데이터 연동 연계 진행 예정 (현재 셋업 준비 중 항목입니다)")