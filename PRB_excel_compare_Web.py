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
    # 사이드바 상단에 로고 배치
    st.sidebar.image("Metanet Fullcolor.png", use_container_width=True)
except:
    pass

st.sidebar.title("🛠️ 검토 설정")
st.title("📊 PRB Data Validation")

# --- [설정] 서버에 저장된 마스터 파일 이름 ---
# 깃허브에 올린 파일명과 대소문자/띄어쓰기가 일치해야 합니다.
MASTER_FILE_NAME = "MDL 통합 SC 인원.xlsx"

# --- 사이드바 메뉴 구성 ---
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

st.sidebar.success(f"현재 모드: {sub_version}")
st.sidebar.info(f"기준 시트: {target_sheet_name}")

# --- 검토 대상 파일(PRB)의 열 설정 (B, D, E열 기준) ---
start_row = 2   # 데이터 시작 행
id_col = 2      # B열: 사번
name_col = 4    # D열: 성명
grade_col = 5   # E열: 등급

# --- 유틸리티 함수 ---
def clean_id(val):
    if val is None or pd.isna(val): return ""
    s = str(val).split('.')[0].strip()
    return s.zfill(7) if s.isdigit() else s

def normalize_grade(val):
    if val is None or pd.isna(val): return "EMPTY"
    return str(val).strip().upper().replace(" ", "").replace("-", "")

def style_p1_results(df):
    def apply_style(row):
        if row['비고'] in ['사번 업데이트', '사번 보정']:
            return ['background-color: #CCE5FF; font-weight: bold;' if col == '변경 사번' else '' for col in df.columns]
        return ['' for _ in df.columns]
    return df.style.apply(apply_style, axis=1)

if 'integrated_results' not in st.session_state:
    st.session_state.integrated_results = None

# --- [UI 수정] 검토 대상 파일 하나만 업로드 ---
st.divider()
target_file = st.file_uploader(f"검토할 {sub_version} 파일을 업로드하세요", type=['xlsx'])

if st.sidebar.button("🧹 결과 데이터 초기화"):
    st.session_state.integrated_results = None
    st.rerun()

# --- 메인 실행 로직 ---
if target_file:
    if st.button("🚀 데이터 검토 시작", use_container_width=True):
        # 1. 서버 내 마스터 파일 존재 여부 확인
        if not os.path.exists(MASTER_FILE_NAME):
            st.error(f"⚠️ 서버에 '{MASTER_FILE_NAME}' 파일이 없습니다. 깃허브 업로드 상태를 확인하세요.")
        else:
            try:
                with st.spinner(f'서버 마스터 데이터({target_sheet_name}) 로드 중...'):
                    # 마스터 파일 읽기
                    df_master = pd.read_excel(MASTER_FILE_NAME, sheet_name=target_sheet_name)
                    
                    master_resources = {}
                    id_to_grade_map = {}

                    for _, row in df_master.iterrows():
                        # 마스터 파일 인덱스: B(1)=사번, D(3)=성명, E(4)=등급
                        m_id = clean_id(row.iloc[1])
                        name = str(row.iloc[3]).strip()
                        m_grade = row.iloc[4]
                        
                        if name not in master_resources: master_resources[name] = []
                        master_resources[name].append({'id': m_id, 'grade': m_grade})
                        id_to_grade_map[m_id] = m_grade

                with st.spinner('업로드 파일 검증 중...'):
                    target_bytes = target_file.getvalue()
                    wb = load_workbook(io.BytesIO(target_bytes))
                    ws = wb.active # 첫 번째 시트 대상
                    
                    fill_blue = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
                    fill_red = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

                    p1_updates, p2_updates = [], []

                    for r_idx in range(start_row, ws.max_row + 1):
                        name_val = ws.cell(r_idx, name_col).value
                        if not name_val or str(name_val).strip() == 'None': continue
                        name = str(name_val).strip()
                        
                        original_id = clean_id(ws.cell(r_idx, id_col).value)
                        original_grade = ws.cell(r_idx, grade_col).value
                        
                        # 사번 보정 로직
                        if name in master_resources:
                            m_id = master_resources[name][0]['id']
                            if original_id != m_id:
                                ws.cell(r_idx, id_col).value = m_id
                                ws.cell(r_idx, id_col).fill = fill_blue
                                p1_updates.append({"행번호": r_idx, "성명": name, "기존 사번": original_id, "변경 사번": m_id, "비고": "사번 보정"})
                                original_id = m_id # 등급 비교를 위해 보정된 사번 사용

                        # 등급 검증 로직
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
        if not res['p1_df'].empty: st.dataframe(style_p1_results(res['p1_df']), use_container_width=True)
        else: st.info("보정 내역 없음")
    with col_b:
        st.subheader("🚩 등급 수정 내역")
        if not res['p2_df'].empty: st.dataframe(res['p2_df'], use_container_width=True)
        else: st.info("수정 내역 없음")

    st.download_button("💾 결과 엑셀 다운로드", res['file_content'], res['file_name'], use_container_width=True, type="primary")