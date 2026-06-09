import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import glob
import threading
import time

# ==============================================================================
# [안전 장치] pyserial 라이브러리가 없을 경우 에러로 멈추지 않고 데모 모드로 전환
# ==============================================================================
SERIAL_AVAILABLE = False
try:
    import serial
    SERIAL_AVAILABLE = True
except ImportError:
    SERIAL_AVAILABLE = False

# ==============================================================================
# [릴리즈 버전 - app_R5 개정판] 지정 폴더(PBA_MonitoringDB) 파일 저장 버전
# ==============================================================================

st.set_page_config(
    page_title="PBA 조립 공정 관리 시스템 (V5-폴더지정형)",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 📂 데이터베이스 전용 저장 폴더 정의 및 자동 생성
DB_FOLDER = "PBA_MonitoringDB"
if not os.path.exists(DB_FOLDER):
    os.makedirs(DB_FOLDER)

# 💾 특정 폴더(PBA_MonitoringDB) 내부의 4대 독립 데이터베이스 파일 경로 바인딩
SOLDER_DB = os.path.join(DB_FOLDER, "pba_solder_process.csv")      # RS-232 상온방치 및 채널 로그
IRON_DB = os.path.join(DB_FOLDER, "pba_iron_check.csv")            # 인두기(1-25번) Factor 검사
MASK_DB = os.path.join(DB_FOLDER, "pba_metal_mask.csv")           # 메탈마스크 모델 및 텐션 5포인트
VISC_DB = os.path.join(DB_FOLDER, "pba_solder_viscosity.csv")     # 솔더크림 Lot별 5회 측정 및 평균

# ------------------------------------------------------------------------------
# 📂 데이터베이스 파일 초기화 및 로드 함수들
# ------------------------------------------------------------------------------
def load_solder_db():
    if not os.path.exists(SOLDER_DB):
        dates = [(datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(15)][::-1]
        df = pd.DataFrame({
            'Date': dates,
            'Solder_Lot_No': ['LOT260601', 'LOT260602', 'LOT260603']*5,
            'Solder_Start_Time': ['08:30']*15,
            'Solder_End_Time': ['10:40']*15,
            'Solder_Leave_Min': [130]*15,
            'Solder_Channel': ['CH1', 'CH2', 'CH3']*5
        })
        df.to_csv(SOLDER_DB, index=False, encoding='utf-8-sig')
    return pd.read_csv(SOLDER_DB)

def load_iron_db():
    if not os.path.exists(IRON_DB):
        dates = [(datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(15)][::-1]
        df = pd.DataFrame({
            'Date': dates,
            'Iron_No': [f"{i}번" for i in np.random.randint(1, 26, size=15)],
            'Iron_Temp': np.random.uniform(345.0, 362.0, size=15).round(1),
            'Iron_Leak_Volt': np.random.uniform(0.5, 1.8, size=15).round(2),
            'Iron_Resistance': np.random.uniform(1.0, 4.2, size=15).round(2)
        })
        df.to_csv(IRON_DB, index=False, encoding='utf-8-sig')
    return pd.read_csv(IRON_DB)

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

# 초기 폴더 및 데이터 데이터 프레임 로드
df_solder = load_solder_db()
df_iron = load_iron_db()
df_mask = load_mask_db()
df_visc = load_visc_db()

def load_daily_fridge_data():
    target_folder = r"C:\fridge_data"
    if not os.path.exists(target_folder):
        times = [datetime.now() - timedelta(minutes=30*i) for i in range(24)]
        df_demo = pd.DataFrame({'DateTime': times, 'Temp': np.random.uniform(4.5, 6.8, size=24).round(1)})
        df_demo['Date'] = df_demo['DateTime'].dt.strftime('%H:%M')
        return df_demo.sort_values('DateTime'), "⚠️ [데모 모드] C:\\fridge_data 폴더 가상 데이터 연동 중"
    try:
        excel_files = glob.glob(os.path.join(target_folder, "*.xls*"))
        if not excel_files: return None, "⚠️ 폴더 내 엑셀 파일 없음"
        chosen_file = max(excel_files, key=os.path.getmtime)
        df = pd.read_excel(chosen_file)
        df.columns = ['DateTime', 'Temp']
        df['DateTime'] = pd.to_datetime(df['DateTime'])
        df = df.sort_values('DateTime')
        df['Date'] = df['DateTime'].dt.strftime('%H:%M')
        return df, f"🔄 연동 성공 ({os.path.basename(chosen_file)})"
    except Exception as e:
        return None, f"❌ 에러: {str(e)}"

# ------------------------------------------------------------------------------
# 🔌 RS-232 실시간 설비 수신 스레드 (지정 폴더 내 SOLDER_DB에 누적 저장)
# ------------------------------------------------------------------------------
SERIAL_PORT = "COM3"
BAUD_RATE = 9600

def rs232_listener():
    if not SERIAL_AVAILABLE:
        while True:
            time.sleep(60)
            continue

    while True:
        try:
            ser = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=1)
            while True:
                if ser.in_waiting > 0:
                    raw_data = ser.readline().decode('utf-8', errors='ignore').strip()
                    if raw_data:
                        parts = raw_data.split(',')
                        if len(parts) >= 4:
                            rx_lot = parts[0].strip().upper()
                            rx_start = parts[1].strip()
                            rx_end = parts[2].strip()
                            rx_ch = parts[3].strip().upper()
                            rx_date = datetime.now().strftime('%Y-%m-%d')
                            
                            try:
                                t1 = datetime.strptime(rx_start, "%H:%M")
                                t2 = datetime.strptime(rx_end, "%H:%M")
                                diff = int((t2 - t1).total_seconds() / 60)
                                if diff < 0: diff += 1440
                            except:
                                diff = 0
                            
                            df_current = pd.read_csv(SOLDER_DB)
                            new_row = pd.DataFrame([{
                                'Date': rx_date, 'Solder_Lot_No': rx_lot, 
                                'Solder_Start_Time': rx_start, 'Solder_End_Time': rx_end, 
                                'Solder_Leave_Min': diff, 'Solder_Channel': rx_ch
                            }])
                            df_updated = pd.concat([df_current, new_row], ignore_index=True)
                            df_updated.to_csv(SOLDER_DB, index=False, encoding='utf-8-sig')
                time.sleep(0.5)
        except Exception:
            time.sleep(10)

if "rs232_thread" not in st.session_state:
    st.session_state.rs232_thread = True
    t = threading.Thread(target=rs232_listener, daemon=True)
    t.start()

# ==============================================================================
# 🏢 대시보드 상단 타이틀 및 KPI 현황판
# ==============================================================================
st.title("🏭 PBA 조립 공정 품질 모니터링 시스템 (V5 - 저장폴더 지정형)")
st.info(f"📂 현재 데이터 파일 저장 폴더: [ {os.path.abspath(DB_FOLDER)} ]")

if SERIAL_AVAILABLE:
    st.success(f"🔌 설비 RS-232 실시간 직렬 통신 모듈 정상 가동 중 (포트: {SERIAL_PORT} / 속도: {BAUD_RATE})")
else:
    st.warning("⚠️ 환경 안내: 현재 구동 서버에 `pyserial` 패키지가 없거나 수동 복사 전이므로 RS-232 통신은 수신 대기 상태입니다.")

st.divider()

df_fridge, fridge_msg = load_daily_fridge_data()
current_fridge_temp = df_fridge['Temp'].iloc[-1] if df_fridge is not None and len(df_fridge)>0 else 0.0
latest_iron_no = df_iron['Iron_No'].iloc[-1] if len(df_iron)>0 else "N/A"
latest_iron_temp = df_iron['Iron_Temp'].iloc[-1] if len(df_iron)>0 else 0.0
latest_mask_model = df_mask['Model_Name'].iloc[-1] if len(df_mask)>0 else "없음"
mask_totals = df_mask.groupby('Model_Name')['Daily_Count'].sum().to_dict()
latest_mask_total = mask_totals.get(latest_mask_model, 0)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
with kpi1:
    f_status = "정상" if 3.0 <= current_fridge_temp <= 10.0 else "🚨 이상"
    st.metric("🧊 냉장고 현재 온도", f"{current_fridge_temp} ℃", delta=f_status)
with kpi2:
    st.metric(f"🔥 최근 인두기 온도 ({latest_iron_no})", f"{latest_iron_temp} ℃", "관리기준: 380±15℃")
with kpi3:
    st.metric(f"📐 메탈 마스크 [{latest_mask_model}] 누적사용횟수", f"{latest_mask_total:,} 회", "직접 입력 모델 기준")
with kpi4:
    latest_visc_avg = df_visc['Average'].iloc[-1] if len(df_visc)>0 else 0.0
    st.metric("🧪 솔더크림 최근 평균점도", f"{latest_visc_avg} Pa·s", "합격 기준 (70~300)")

st.divider()

# ==============================================================================
# 🎛️ 사이드바 - 관리 항목별 입력 폼
# ==============================================================================
st.sidebar.header("🎛️ 데이터 입력 및 이력 수정")
input_category = st.sidebar.radio("작업 선택", ["1. 솔더크림 상온방치 수동입력", "2. 메탈마스크 & 텐션 관리", "3. 솔더 Lot별 점도 관리"])

# --- 사이드바 1: 솔더크림 상온방치 수동 등록폼 ---
if input_category == "1. 솔더크림 상온방치 수동입력":
    st.sidebar.subheader("⏳ 상온방치 수동 관리 입력폼")
    target_date = st.sidebar.date_input("방치 일자 선택", datetime.now(), key="solder_date")
    date_str = target_date.strftime('%Y-%m-%d')
    
    with st.sidebar.form("solder_lot_form"):
        sb_lot = st.text_input("Solder 크림 Lot 번호 입력", value="LOT260610").strip().upper()
        sb_start = st.text_input("방치 시작 시간 (HH:MM)", value="08:30")
        sb_end = st.text_input("방치 종료 시간 (HH:MM)", value="10:45")
        sb_channel = st.selectbox("사용 채널 선택", ["CH1", "CH2", "CH3", "CH4", "CH5"])
        
        submit_solder = st.form_submit_button("⏳ 방치 이력 수동 저장")
        
    if submit_solder and sb_lot:
        try:
            t1 = datetime.strptime(sb_start, "%H:%M")
            t2 = datetime.strptime(sb_end, "%H:%M")
            leave_duration = int((t2 - t1).total_seconds() / 60)
            if leave_duration < 0: leave_duration += 1440
        except:
            leave_duration = 0

        new_data = pd.DataFrame([{
            'Date': date_str, 'Solder_Lot_No': sb_lot, 
            'Solder_Start_Time': sb_start, 'Solder_End_Time': sb_end, 
            'Solder_Leave_Min': leave_duration, 'Solder_Channel': sb_channel
        }])
        df_solder = pd.concat([df_solder, new_data], ignore_index=True).sort_values('Date')
        df_solder.to_csv(SOLDER_DB, index=False, encoding='utf-8-sig')
        st.sidebar.success(f"💾 {sb_lot} 자재 방치 데이터 추가 완료!")
        st.rerun()

# --- 사이드바 2: 메탈마스크 모델 직접 입력 및 텐션 관리 ---
elif input_category == "2. 메탈 마스크 & 텐션 관리":
    st.sidebar.subheader("📐 메탈 마스크 직접입력 및 텐션 제어")
    target_date = st.sidebar.date_input("점검 일자 선택", datetime.now(), key="mask_date")
    date_str = target_date.strftime('%Y-%m-%d')
    sm_model = st.sidebar.text_input("가동 메탈 마스크 모델명 직접 입력", value="SMT-PRIME-01").strip().upper()
    
    with st.sidebar.form("mask_form"):
        sm_count = st.number_input("당일 가동 타수 (회)", min_value=0, value=1500)
        st.markdown("**📌 텐션 측정치 (스펙: 0.45 ~ 0.90 mm)**")
        mp1 = st.number_input("Point 1 (Center)", value=0.65, step=0.01)
        mp2 = st.number_input("Point 2 (Top-Left)", value=0.68, step=0.01)
        mp3 = st.number_input("Point 3 (Top-Right)", value=0.64, step=0.01)
        mp4 = st.number_input("Point 4 (Bottom-Left)", value=0.67, step=0.01)
        mp5 = st.number_input("Point 5 (Bottom-Right)", value=0.66, step=0.01)
        
        submit_mask = st.form_submit_button("📐 메탈 마스크 정보 저장")
        
    if submit_mask and sm_model:
        new_mask_data = pd.DataFrame([{
            'Date': date_str, 'Model_Name': sm_model, 'Daily_Count': sm_count,
            'P1_Center': round(mp1, 2), 'P2_TopLeft': round(mp2, 2), 'P3_TopRight': round(mp3, 2),
            'P4_BotLeft': round(mp4, 2), 'P5_BotRight': round(mp5, 2)
        }])
        df_mask = pd.concat([df_mask, new_mask_data], ignore_index=True).sort_values('Date')
        df_mask.to_csv(MASK_DB, index=False, encoding='utf-8-sig')
        st.sidebar.success(f"💾 {sm_model} 모델 데이터 저장 완료.")
        st.rerun()

# --- 사이드바 3: 솔더 Lot별 점도 관리 ---
elif input_category == "3. 솔더 Lot별 점도 관리":
    st.sidebar.subheader("🧪 솔더 점도 5회 측정 입력")
    target_date = st.sidebar.date_input("측정 일자 선택", datetime.now(), key="visc_date")
    date_str = target_date.strftime('%Y-%m-%d')
    v_lot = st.sidebar.text_input("Solder 크림 Lot 번호 입력", value="LOT260601").strip().upper()
    
    with st.sidebar.form("visc_form"):
        m1 = st.number_input("1회차 측정값 (Pa·s)", value=190)
        m2 = st.number_input("2회차 측정값 (Pa·s)", value=200)
        m3 = st.number_input("3회차 측정값 (Pa·s)", value=205)
        m4 = st.number_input("4회차 측정값 (Pa·s)", value=195)
        m5 = st.number_input("5회차 측정값 (Pa·s)", value=210)
        submit_visc = st.form_submit_button("🧪 점도 데이터 저장")
        
    if submit_visc and v_lot:
        v_avg = round(np.mean([m1, m2, m3, m4, m5]), 1)
        new_visc_data = pd.DataFrame([{
            'Date': date_str, 'Lot_No': v_lot, 'M1': m1, 'M2': m2, 'M3': m3, 'M4': m4, 'M5': m5, 'Average': v_avg
        }])
        df_visc = pd.concat([df_visc, new_visc_data], ignore_index=True).sort_values('Date')
        df_visc.to_csv(VISC_DB, index=False, encoding='utf-8-sig')
        st.sidebar.success("💾 점도 데이터 세트 보존 성공.")
        st.rerun()

# ==============================================================================
# 📊 메인 화면 : 공정 탭별 세부 현황 대시보드
# ==============================================================================
tab_fridge, tab_solder, tab_mask, tab_iron = st.tabs([
    "🧊 1. 냉장고 온도 제어", 
    "⏳ 2. 솔더크림 (RS-232 상온방치/점도)", 
    "📐 3. 메탈 마스크 (직접입력 모델/5P 텐션)", 
    "⚡ 4. 인두기 검사 (1번~25번 관리)"
])

# --- [TAB 1] 냉장고 온도 제어 섹션 ---
with tab_fridge:
    st.subheader("🧊 보관 냉장고 일별 생성 로그 실시간 추적")
    st.info(fridge_msg)
    if df_fridge is not None and len(df_fridge) > 0:
        fig_f = px.line(df_fridge, x='Date', y='Temp', title="당일 생성된 파일 내 실시간 온도 로깅 추이", markers=True)
        fig_f.add_hrect(y0=3.0, y1=10.0, fillcolor="green", opacity=0.08, line_width=0, annotation_text="정상 범위 (3~10℃)")
        st.plotly_chart(fig_f, width="stretch")

# --- [TAB 2] 솔더크림 상온방치 및 점도 관리 섹션 ---
with tab_solder:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.subheader("⏳ Lot별 상온방치 이력 현황")
        df_solder['Status'] = df_solder['Solder_Leave_Min'].apply(lambda x: 'OK (2시간이상)' if x >= 120 else 'NG (시간미달)')
        
        fig_s1 = px.bar(df_solder, x='Date', y='Solder_Leave_Min', color='Status',
                        hover_data=['Solder_Lot_No', 'Solder_Start_Time', 'Solder_End_Time', 'Solder_Channel'],
                        color_discrete_map={'OK (2시간이상)': '#2ECC71', 'NG (시간미달)': '#E74C3C'},
                        text_auto=True, title="일자별 방치 경과 시간 및 합격 여부")
        fig_s1.add_hline(y=120, line_dash="dash", line_color="black")
        st.plotly_chart(fig_s1, width="stretch")
    with col_s2:
        st.subheader("🧪 AIM M8 제품 Lot별 점도 관리 (70-300 Pa·s)")
        if len(df_visc) > 0:
            fig_s2 = go.Figure()
            fig_s2.add_trace(go.Scatter(x=df_visc['Lot_No'], y=df_visc['Average'], mode='lines+markers', name='Lot별 평균값', line=dict(color='blue', width=3)))
            fig_s2.add_hrect(y0=70, y1=300, fillcolor="blue", opacity=0.05, line_width=0)
            st.plotly_chart(fig_s2, width="stretch")
    st.dataframe(df_solder[['Date', 'Solder_Lot_No', 'Solder_Start_Time', 'Solder_End_Time', 'Solder_Leave_Min', 'Solder_Channel', 'Status']], width="stretch")

# --- [TAB 3] 메탈 마스크 사용횟수 및 텐션 관리 섹션 ---
with tab_mask:
    st.subheader("📐 메탈 마스크 직접 입력 모델별 누적 사용 횟수")
    df_mask_agg = df_mask.groupby('Model_Name').agg(최근작업일=('Date', 'max'), 최근작업타수=('Daily_Count', 'last')).reset_index()
    df_mask_agg['역대_총_누적_타수'] = df_mask_agg['Model_Name'].map(mask_totals)
    
    col_m1, col_m2 = st.columns([1, 2])
    with col_m1: st.dataframe(df_mask_agg, width="stretch")
    with col_m2:
        fig_m1 = px.bar(df_mask, x='Date', y='Daily_Count', color='Model_Name', title="일자별/모델별 메탈 마스크 실적 분포", text_auto=True)
        st.plotly_chart(fig_m1, width="stretch")
        
    st.divider()
    st.subheader("📌 메탈 마스크 5개 포인트 텐션 모니터링 (0.45mm - 0.90mm)")
    fig_m2 = go.Figure()
    x_labels = df_mask['Date'] + " (" + df_mask['Model_Name'] + ")"
    for col in ['P1_Center', 'P2_TopLeft', 'P3_TopRight', 'P4_BotLeft', 'P5_BotRight']:
        fig_m2.add_trace(go.Scatter(x=x_labels, y=df_mask[col], name=col, mode='lines+markers'))
    fig_m2.add_hrect(y0=0.45, y1=0.90, fillcolor="orange", opacity=0.07, line_width=0)
    st.plotly_chart(fig_m2, width="stretch")

# --- [TAB 4] 인두기 정기 검사 섹션 ---
with tab_iron:
    st.subheader("⚡ 인두기 호기별 관리 센터 (1번 ~ 25번)")
    
    with st.expander("➕ ⚙️ 현장 실시간 인두기 측정 데이터 등록/수정 양식 열기", expanded=False):
        st.markdown("### 📝 호기별 인두기 Factor 관리 규격 입력")
        col_in1, col_in2, col_in3 = st.columns(3)
        with col_in1:
            iron_date = st.date_input("측정 일자 지정", datetime.now(), key="iron_input_date")
            iron_no_select = st.selectbox("🎯 인두기 호기 선택", [f"{i}번" for i in range(1, 26)], key="iron_input_no")
        with col_in2:
            iron_temp_val = st.number_input("인두기 팁 온도 (℃)", min_value=300.0, max_value=450.0, value=350.0, step=1.0)
            iron_volt_val = st.number_input("누설 전압 (mV)", min_value=0.0, value=1.2, step=0.1)
        with col_in3:
            iron_res_val = st.number_input("접지 저항 (Ω)", min_value=0.0, value=2.0, step=0.1)
            
        submit_iron_data = st.button("⚡ 인두기 측정값 데이터베이스 저장/변경")
        
        if submit_iron_data:
            i_date_str = iron_date.strftime('%Y-%m-%d')
            idx_exist = df_iron[(df_iron['Date'] == i_date_str) & (df_iron['Iron_No'] == iron_no_select)].index
            if not idx_exist.empty:
                df_iron = df_iron.drop(idx_exist)
                
            new_iron = pd.DataFrame([{
                'Date': i_date_str, 
                'Iron_No': iron_no_select, 
                'Iron_Temp': round(iron_temp_val, 1), 
                'Iron_Leak_Volt': round(iron_volt_val, 2), 
                'Iron_Resistance': round(iron_res_val, 2)
            }])
            df_iron = pd.concat([df_iron, new_iron], ignore_index=True).sort_values(['Date', 'Iron_No'])
            df_iron.to_csv(IRON_DB, index=False, encoding='utf-8-sig')
            st.success(f"💾 데이터가 {DB_FOLDER} 폴더 내부 파일에 안전하게 기록되었습니다.")
            st.rerun()

    st.markdown("---")
    selected_iron_view = st.selectbox("🔍 시각화 필터링할 인두기 호기 선택", [f"{i}번" for i in range(1, 26)], key="iron_view_select")
    df_iron_filtered = df_iron[df_iron['Iron_No'] == selected_iron_view]
    
    if len(df_iron_filtered) > 0:
        col_i1, col_i2, col_i3 = st.columns(3)
        with col_i1:
            fig_i1 = px.line(df_iron_filtered, x='Date', y='Iron_Temp', title=f"🌡️ {selected_iron_view} 팁 온도 트렌드 (380±15℃)", markers=True, color_discrete_sequence=['#E67E22'])
            fig_i1.add_hrect(y0=365.0, y1=395.0, fillcolor="orange", opacity=0.1, annotation_text="정상범위")
            st.plotly_chart(fig_i1, width="stretch")
        with col_i2:
            fig_i2 = px.bar(df_iron_filtered, x='Date', y='Iron_Leak_Volt', title=f"⚡ {selected_iron_view} 누설 전압 (Spec: 2.0mV 이하)", color_discrete_sequence=['#9B59B6'])
            fig_i2.add_hline(y=2.0, line_dash="dash", line_color="red")
            st.plotly_chart(fig_i2, width="stretch")
        with col_i3:
            fig_i3 = px.line(df_iron_filtered, x='Date', y='Iron_Resistance', title=f"🔌 {selected_iron_view} 접지 저항 (Spec: 5.0Ω 이하)", markers=True, color_discrete_sequence=['#34495E'])
            fig_i3.add_hline(y=5.0, line_dash="dash", line_color="red")
            st.plotly_chart(fig_i3, width="stretch")
            
        st.dataframe(df_iron_filtered[['Date', 'Iron_No', 'Iron_Temp', 'Iron_Leak_Volt', 'Iron_Resistance']], width="stretch")
    else:
        st.warning(f"선택하신 {selected_iron_view} 인두기는 등록된 이력이 없습니다.")
