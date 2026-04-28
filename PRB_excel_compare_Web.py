# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io
import os

# [1] 페이지 설정
st.set_page_config(page_title="PRB 인건비 통합 검토", layout="wide")

# --- 법인별 마스터 파일명 설정 (수정 완료) ---
# ※ 주의: GitHub에 업로드된 파일명과 대소문자/공백까지 완벽히 일치해야 합니다.
MASTER_FILES = {
    "Metanet DL": "MDL 통합 SC 인원.xlsx",
    "Metanet Fintech": "MF 통합 SC 인원.xlsx",
    "Metanet Digital": "MD 통합 SC 인원.xlsx",
    "Skelter Labs": "SKL 통합 SC 인원.xlsx"
}

# --- 로고 및 사이드바 설정 ---
try:
    st.sidebar.image("Metanet Fullcolor.png", use_container_width=True)
except:
    pass

st.sidebar.title("🛠️ 검토 설정")
st.title("📊 PRB Data Validation")

# --- 1단계: 법인 및 버전 선택 ---
main_category = st.sidebar.selectbox("법인 그룹을 선택하세요", list(MASTER_FILES.keys()))

if main_category == "Metanet DL":
    sub_version = st.sidebar.radio("세부 버전을 선택하세요", ["(DL) 2026년 3월 이전 버전", "(DL) 2026년 3월 이후 버전"])
elif main_category == "Metanet Fintech":
    sub_version = st.sidebar.radio("세부 버전을 선택하세요", ["(MF) 2026년 3월 이전 버전", "(MF) 2026년 3월 이후 버전"])
elif main_category == "Metanet Digital":
    sub_version = st.sidebar.radio("세부 버전을 선택하세요", ["(MD) ver1.1"])
else:
    sub_version = st.sidebar.radio("세부 버전을 선택하세요", ["(SKL) ver.1.1"])

# 현재 선택된 마스터 파일 경로 및 시트 이름 결정
CURRENT_MASTER_PATH = MASTER_FILES[main_category]

if "이전 버전" in sub_version:
    CURRENT_MASTER_SHEET = "(old)SC인원 현황_B15"
else:
    CURRENT_MASTER_SHEET = "(new)SC인원 현황_B20"

st.sidebar.success(f"현재 모드: {sub_version}")
st.sidebar.info(f"검증 기준 파일: {CURRENT_MASTER_PATH}\n기준 시트: {CURRENT_MASTER_SHEET}")

# 검증 대상 파일(업로드된 파일)의 데이터 시작 위치 및 컬럼 인덱스
# B=6(사번), C=7(성명), D=8(직급) 이라고 가정 (필요시 조정)
start_row, id_col, name_col, grade_col = 6, 6, 7, 8

# --- 유틸리티 함수 ---
def clean_id(val):
    if val is None or pd.isna(val): return ""
    s = str(val).split('.')[0].strip()
    return s.zfill(7) if s.isdigit() else s

def normalize_grade(val):
    if val is None or pd.isna(val): return "EMPTY"
    s = str(val).strip().upper().replace(" ", "").replace("-", "")
    if "PJ(B)계약" in s: s = s.replace("PJ(B)계약", "계약B")
    if "PJ(C)계약" in s: s = s.replace("PJ(C)계약", "계약C")
    return s

def convert_to_target_format(master_grade):
    if master_grade is None or pd.isna(master_grade): return ""
    s = str(master_grade).strip()
    s = s.replace("PJ(B)-계약", "계약B").replace("PJ(C)-계약", "계약C")
    return s

def style_p1_results(df):
    def apply_style(row):
        color_map = {
            '사번 업데이트': 'background-color: #CCE5FF; font-weight: bold;',
            '사번 보정': 'background-color: #D5E8D4; font-weight: bold;',
            '동명이인': 'background-color: #FFCCCC; font-weight: bold;'
        }
        style = color_map.get(row['비고'], '')
        return [style if col == '변경 사번' else '' for col in df.columns]
    return df.style.apply(apply_style, axis=1)

# --- 메인 실행 로직 ---
if 'integrated_results' not in st.session_state:
    st.session_state.integrated_results = None

st.divider()
target_file = st.file_uploader(f"검증할 대상 {sub_version} 파일을 업로드하세요", type=['xlsx'])

if st.sidebar.button("🧹 데이터 초기화"):
    st.session_state.integrated_results = None
    st.rerun()

if target_file:
    if st.button("🚀 데이터 검토 시작", use_container_width=True):
        # 파일 존재 여부 확인
        if not os.path.exists(CURRENT_MASTER_PATH):
            st.error(f"❌ 기준 파일('{CURRENT_MASTER_PATH}')을 찾을 수 없습니다. GitHub 저장소에 파일이 있는지 확인하세요.")
        else:
            try:
                with st.spinner(f'마스터 파일 [{CURRENT_MASTER_SHEET}] 분석 중...'):
                    # 엑셀의 6행부터 데이터가 있다면 skiprows=5 적용
                    df_master = pd.read_excel(CURRENT_MASTER_PATH, sheet_name=CURRENT_MASTER_SHEET, skiprows=5, engine='openpyxl')
                    
                    master_resources = {}
                    id_to_grade_map = {} 

                    for _, row in df_master.iterrows():
                        # skiprows 적용 후: 사번(0번 인덱스), 성명(2번 인덱스), 등급(3번 인덱스) 확인 필요
                        try:
                            m_id = clean_id(row.iloc[0])   
                            name = str(row.iloc[2]).strip() 
                            m_grade = row.iloc[3]          
                            
                            if name == 'nan' or not name: continue
                            
                            if name not in master_resources: master_resources[name] = []
                            master_resources[name].append({'id': m_id, 'grade': m_grade, 'used': False})
                            id_to_grade_map[m_id] = m_grade
                        except:
                            continue

                    # 대상 파일 처리 (업로드한 파일)
                    target_bytes = target_file.getvalue()
                    wb = load_workbook(io.BytesIO(target_bytes))
                    ws = wb['A3.자사인건비'] if 'A3.자사인건비' in wb.sheetnames else wb.active
                    
                    fills = {
                        "blue": PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid"),
                        "green": PatternFill(start_color="D5E8D4", end_color="D5E8D4", fill_type="solid"),
                        "red": PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                    }

                    p1_updates, p2_updates = [], []
                    
                    for r_idx in range(start_row, ws.max_row + 1):
                        name_val = ws.cell(r_idx, name_col).value
                        if not name_val or str(name_val).strip() in ['None', '성명', 'NAN']: continue
                        
                        name = str(name_val).strip()
                        original_id = clean_id(ws.cell(r_idx, id_col).value)
                        original_grade = ws.cell(r_idx, grade_col).value
                        final_id = original_id
                        
                        # [검증 1] 사번 체크 및 보정
                        if name in master_resources:
                            m_list = master_resources[name]
                            match = next((m for m in m_list if not m['used']), m_list[0])
                            final_id = match['id']
                            match['used'] = True
                            
                            if original_id != final_id:
                                ws.cell(r_idx, id_col).value = final_id
                                ws.cell(r_idx, id_col).fill = fills["blue"] if not original_id else fills["green"]
                                p1_updates.append({
                                    "행번호": r_idx, "성명": name, 
                                    "기존 사번": original_id if original_id else "공란", 
                                    "변경 사번": final_id, 
                                    "비고": "사번 업데이트" if not original_id else "사번 보정"
                                })

                        # [검증 2] 등급 체크
                        if final_id in id_to_grade_map:
                            m_grade = id_to_grade_map[final_id]
                            if normalize_grade(original_grade) != normalize_grade(m_grade):
                                fixed_grade = convert_to_target_format(m_grade)
                                ws.cell(r_idx, grade_col).value = fixed_grade
                                ws.cell(r_idx, grade_col).fill = fills["red"]
                                p2_updates.append({
                                    "행번호": r_idx, "사번": final_id, "성명": name, 
                                    "기존 등급": original_grade, "변경 등급": fixed_grade
                                })

                    output = io.BytesIO()
                    wb.save(output)
                    st.session_state.integrated_results = {
                        'p1_df': pd.DataFrame(p1_updates),
                        'p2_df': pd.DataFrame(p2_updates),
                        'file_content': output.getvalue(),
                        'file_name': f"PRB_검토결과_{sub_version.replace(' ', '_')}.xlsx"
                    }
                    st.balloons()

            except Exception as e:
                st.error(f"⚠️ 실행 중 오류가 발생했습니다: {e}")
                import traceback
                st.expander("상세 에러 로그 확인").code(traceback.format_exc())

# 결과 화면 출력부
if st.session_state.integrated_results:
    res = st.session_state.integrated_results
    st.divider()
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("🚩 사번 보정 내역")
        if not res['p1_df'].empty: 
            st.dataframe(style_p1_results(res['p1_df']), use_container_width=True)
        else: 
            st.info("✅ 사번 불일치 내역이 없습니다.")
    with col_b:
        st.subheader("🚩 등급 수정 내역")
        if not res['p2_df'].empty: 
            st.dataframe(res['p2_df'], use_container_width=True)
        else: 
            st.info("✅ 등급 불일치 내역이 없습니다.")