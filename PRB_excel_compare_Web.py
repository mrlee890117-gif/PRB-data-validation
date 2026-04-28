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

# --- [수정 포인트] 사이드바 메뉴 구성 (4가지 카테고리) ---

# 1단계: 법인 그룹 선택
main_category = st.sidebar.selectbox(
    "법인 그룹을 선택하세요",
    ["Metanet DL", "Metanet Fintech", "Metanet Digital", "Skelter Labs"]
)

# 2단계: 법인별 맞춤형 세부 버전 설정
if main_category == "Metanet DL":
    sub_version = st.sidebar.radio(
        "세부 버전을 선택하세요",
        ["(old) 2026년 3월 이전 버전", "(new) 2026년 3월 이후 버전"]
    )
elif main_category == "Metanet Fintech":
    sub_version = st.sidebar.radio(
        "세부 버전을 선택하세요",
        ["(old) 2026년 3월 이전 버전", "(new) 2026년 3월 이후 버전"]
    )
elif main_category == "Metanet Digital":
    sub_version = st.sidebar.radio(
        "세부 버전을 선택하세요",
        ["(MD) ver.1.1 "]
    )
else: # Skelter Labs
    sub_version = st.sidebar.radio(
        "세부 버전을 선택하세요",
        ["(SKL) ver.1.1 "]
    )

st.sidebar.success(f"현재 모드: {sub_version}")

# --- 버전별 엑셀 설정값 (양식에 맞춰 숫자 수정) ---
# Tip: sub_version의 이름이 정확히 일치해야 합니다.
if "(DL) 2026년 3월 이전 버전" in sub_version:
    start_row, id_col, name_col, grade_col = 6, 7, 8, 9
elif "(DL) 2026년 3월 이후 버전" in sub_version:
    start_row, id_col, name_col, grade_col = 6, 7, 8, 9
else:
    # 나머지 법인들의 기본 설정값
    start_row, id_col, name_col, grade_col = 6, 7, 8, 9

st.markdown(f"### **검토 대상: {sub_version}**")

# --- 표준화 및 유틸리티 로직 ---
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
        if row['비고'] == '사번 업데이트':
            return ['background-color: #CCE5FF; font-weight: bold;' if col == '변경 사번' else '' for col in df.columns]
        elif row['비고'] == '사번 보정':
            return ['background-color: #D5E8D4; font-weight: bold;' if col == '변경 사번' else '' for col in df.columns]
        elif row['비고'] == '동명이인':
            return ['background-color: #FFCCCC; font-weight: bold;' if col == '변경 사번' else '' for col in df.columns]
        return ['' for _ in df.columns]
    return df.style.apply(apply_style, axis=1)

if 'integrated_results' not in st.session_state:
    st.session_state.integrated_results = None

# --- 파일 업로드 UI ---
st.divider()
col_u1, col_u2 = st.columns(2)
with col_u1:
    master_file = st.file_uploader("기준 등급 파일 (2행 시작)", type=['xlsx'])
with col_u2:
    target_file = st.file_uploader(f"대상 {sub_version} 파일", type=['xlsx'])

if st.sidebar.button("🧹 데이터 초기화"):
    st.session_state.integrated_results = None
    st.rerun()

# --- 메인 실행 로직 ---
if master_file and target_file:
    if st.button("🚀 데이터 검토 시작", use_container_width=True):
        try:
            with st.spinner(f'{sub_version} 데이터 분석 중...'):
                master_bytes = master_file.getvalue()
                try:
                    df_master = pd.read_excel(io.BytesIO(master_bytes), sheet_name='SC 인원현황', header=0)
                except:
                    df_master = pd.read_excel(io.BytesIO(master_bytes), header=0)
                
                master_resources = {}
                id_to_grade_map = {} 

                for _, row in df_master.iterrows():
                    m_id = clean_id(row.iloc[0])    
                    name = str(row.iloc[1]).strip() 
                    m_grade = row.iloc[12]          
                    
                    if name not in master_resources: master_resources[name] = []
                    master_resources[name].append({'id': m_id, 'grade': m_grade, 'used': False})
                    id_to_grade_map[m_id] = m_grade

                target_bytes = target_file.getvalue()
                wb = load_workbook(io.BytesIO(target_bytes))
                ws = wb['A3.자사인건비'] if 'A3.자사인건비' in wb.sheetnames else wb.active
                
                fill_blue = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
                fill_green = PatternFill(start_color="D5E8D4", end_color="D5E8D4", fill_type="solid")
                fill_red = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                fill_grade_err = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

                p1_updates, p2_updates = [], []
                target_pair_counts = {}

                for r in range(start_row, ws.max_row + 1):
                    t_id = clean_id(ws.cell(r, id_col).value)
                    t_name = str(ws.cell(r, name_col).value).strip()
                    if t_name and t_name != 'None':
                        pair = (t_name, t_id)
                        target_pair_counts[pair] = target_pair_counts.get(pair, 0) + 1

                for r_idx in range(start_row, ws.max_row + 1):
                    name_val = ws.cell(r_idx, name_col).value
                    if not name_val or str(name_val).strip() == 'None': continue
                    name = str(name_val).strip()
                    
                    original_id = clean_id(ws.cell(r_idx, id_col).value)
                    original_grade = ws.cell(r_idx, grade_col).value
                    final_id = original_id
                    
                    if name in master_resources:
                        m_list = master_resources[name]
                        if len(m_list) > 1:
                            match = next((m for m in m_list if not m['used']), m_list[0])
                            final_id = match['id']
                            match['used'] = True
                            ws.cell(r_idx, id_col).value = final_id
                            ws.cell(r_idx, id_col).fill = fill_red
                            p1_updates.append({"행번호": r_idx, "성명": name, "기존 사번": original_id if original_id else "공란", "변경 사번": final_id, "비고": "동명이인"})
                        else:
                            match = m_list[0]
                            final_id = match['id']
                            if original_id != final_id:
                                ws.cell(r_idx, id_col).value = final_id
                                ws.cell(r_idx, id_col).fill = fill_blue if not original_id else fill_green
                                p1_updates.append({"행번호": r_idx, "성명": name, "기존 사번": original_id if original_id else "공란", "변경 사번": final_id, "비고": "사번 업데이트" if not original_id else "사번 보정"})
                            elif target_pair_counts.get((name, original_id), 0) > 1:
                                p1_updates.append({"행번호": r_idx, "성명": name, "기존 사번": original_id, "변경 사번": final_id, "비고": "동일인 중복"})

                    if final_id in id_to_grade_map:
                        m_grade = id_to_grade_map[final_id]
                        if normalize_grade(original_grade) != normalize_grade(m_grade):
                            fixed_grade = convert_to_target_format(m_grade)
                            ws.cell(r_idx, grade_col).value = fixed_grade
                            ws.cell(r_idx, grade_col).fill = fill_grade_err
                            p2_updates.append({"행번호": r_idx, "사번": final_id, "성명": name, "기존 등급": original_grade, "변경 등급": fixed_grade})

                output = io.BytesIO()
                wb.save(output)
                st.session_state.integrated_results = {
                    'p1_df': pd.DataFrame(p1_updates),
                    'p2_df': pd.DataFrame(p2_updates),
                    'file_content': output.getvalue(),
                    'file_name': f"PRB_검토결과_{sub_version.split(')')[0]}.xlsx"
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
        st.subheader("🚩 사번 확인 필요 내역")
        if not res['p1_df'].empty: st.dataframe(style_p1_results(res['p1_df']), use_container_width=True)
        else: st.info("특이사항 없음")

    with col_b:
        st.subheader("🚩 등급 수정 내역")
        if not res['p2_df'].empty: st.dataframe(res['p2_df'], use_container_width=True)
        else: st.info("특이사항 없음")

    st.download_button("💾 검토 결과 엑셀 다운로드", res['file_content'], res['file_name'], use_container_width=True, type="primary")