# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io
import os

# [1] 페이지 설정
st.set_page_config(page_title="PRB 인건비 통합 검토", layout="wide")

# --- 로고 및 사이드바 설정 ---
try:
    st.sidebar.image("Metanet Fullcolor.png", use_container_width=True)
except:
    pass

st.sidebar.title("🛠️ 검토 설정")
st.title("📊 PRB Data Validation")

# --- [중요] 깃허브에 있는 실제 엑셀 파일명과 일치시켜야 함 ---
MASTER_FILE_NAME = "MDL 통합 SC 인원.xlsx"

# --- 서버 파일 진단 기능 ---
st.sidebar.divider()
if not os.path.exists(MASTER_FILE_NAME):
    st.sidebar.error(f"❌ 파일을 찾을 수 없습니다: {MASTER_FILE_NAME}")
    st.sidebar.write("📂 **현재 서버 파일 목록:**")
    st.sidebar.write(os.listdir(".")) 
else:
    st.sidebar.success(f"✅ 기준 파일 로드 완료")

# --- 법인 및 시트 선택 로직 ---
main_category = st.sidebar.selectbox(
    "법인 그룹을 선택하세요",
    ["Metanet DL", "Metanet Fintech", "Metanet Digital", "Skelter Labs"]
)

target_sheet_name = ""
if main_category == "Metanet DL":
    sub_version = st.sidebar.radio(
        "세부 버전을 선택하세요",
        ["(old) 2026년 3월 이전 버전", "(new) 2026년 3월 이후 버전"]
    )
    if "old" in sub_version:
        target_sheet_name = "(old)SC인원 현황_B15"
    else:
        target_sheet_name = "(new)SC인원 현황_B20"
else:
    sub_version = st.sidebar.radio("세부 버전을 선택하세요", ["기본 양식"])
    target_sheet_name = "SC 인원현황"

# --- 열 설정 (B, D, E) ---
start_row = 2   
id_col = 2      # B열
name_col = 4    # D열
grade_col = 5   # E열

# --- 유틸리티 함수 ---
def clean_id(val):
    if val is None or pd.isna(val): return ""
    s = str(val).split('.')[0].strip()
    return s.zfill(7) if s.isdigit() else s

def normalize_grade(val):
    if val is None or pd.isna(val): return "EMPTY"
    return str(val).strip().upper().replace(" ", "").replace("-", "")

if 'integrated_results' not in st.session_state:
    st.session_state.integrated_results = None

# --- UI: 검토 대상 파일 업로드 (업로드 칸 1개) ---
st.divider()
target_file = st.file_uploader(f"검토할 {sub_version} 파일을 업로드하세요", type=['xlsx'])

if st.sidebar.button("🧹 결과 초기화"):
    st.session_state.integrated_results = None
    st.rerun()

# --- 실행 로직 ---
if target_file:
    if st.button("🚀 데이터 검토 시작", use_container_width=True):
        if not os.path.exists(MASTER_FILE_NAME):
            st.error(f"⚠️ 서버에 '{MASTER_FILE_NAME}' 파일이 없습니다.")
        else:
            try:
                with st.spinner('마스터 데이터 분석 중...'):
                    df_master = pd.read_excel(MASTER_FILE_NAME, sheet_name=target_sheet_name)
                    master_resources = {}
                    id_to_grade_map = {}
                    for _, row in df_master.iterrows():
                        m_id = clean_id(row.iloc[1])
                        name = str(row.iloc[3]).strip()
                        m_grade = row.iloc[4]
                        if name not in master_resources: master_resources[name] = []
                        master_resources[name].append({'id': m_id, 'grade': m_grade})
                        id_to_grade_map[m_id] = m_grade

                with st.spinner('검증 진행 중...'):
                    target_bytes = target_file.getvalue()
                    wb = load_workbook(io.BytesIO(target_bytes))
                    ws = wb.active
                    fill_blue = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
                    fill_red = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                    p1_updates, p2_updates = [], []

                    for r_idx in range(start_row, ws.max_row + 1):
                        name_val = ws.cell(r_idx, name_col).value
                        if not name_val or str(name_val).strip() == 'None': continue
                        name = str(name_val).strip()
                        original_id = clean_id(ws.cell(r_idx, id_col).value)
                        original_grade = ws.cell(r_idx, grade_col).value
                        
                        if name in master_resources:
                            m_id = master_resources[name][0]['id']
                            if original_id != m_id:
                                ws.cell(r_idx, id_col).value = m_id
                                ws.cell(r_idx, id_col).fill = fill_blue
                                p1_updates.append({"행번호": r_idx, "성명": name, "기존 사번": original_id, "변경 사번": m_id, "비고": "사번 보정"})
                                original_id = m_id

                        if original_id in id_to_grade_map:
                            m_grade = id_to_grade_map[original_id]
                            if normalize_grade(original_grade) != normalize_grade(m_grade):
                                ws.cell(r_idx, grade_col).value = m_grade
                                ws.cell(r_idx, grade_col).fill = fill_red
                                p2_updates.append({"행번호": r_idx, "사번": original_id, "성명": name, "기존 등급": original_grade, "변경 등급": m_grade})

                    output = io.BytesIO()
                    wb.save(output)
                    st.session_state.integrated_results = {
                        'p1_df': pd.DataFrame(p1_updates),
                        'p2_df': pd.DataFrame(p2_updates),
                        'file_content': output.getvalue(),
                        'file_name': f"검토결과_{target_sheet_name}.xlsx"
                    }
                    st.balloons()
            except Exception as e:
                st.error(f"⚠️ 오류 발생: {e}")

# --- 결과 표출 ---
if st.session_state.integrated_results:
    res = st.session_state.integrated_results
    st.divider()
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("🚩 사번 보정 내역")
        if not res['p1_df'].empty: st.dataframe(res['p1_df'], use_container_width=True)
        else: st.info("보정 내역 없음")
    with col_b:
        st.subheader("🚩 등급 수정 내역")
        if not res['p2_df'].empty: st.dataframe(res['p2_df'], use_container_width=True)
        else: st.info("수정 내역 없음")
    st.download_button("💾 결과 엑셀 다운로드", res['file_content'], res['file_name'], use_container_width=True, type="primary")