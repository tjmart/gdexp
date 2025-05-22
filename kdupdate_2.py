import streamlit as st
import pandas as pd
import mysql.connector
import io
import datetime
import re
import numpy as np

# 페이지 설정: 반드시 가장 첫 번째 Streamlit 호출
st.set_page_config(page_title="엑셀 DB 관리 시스템", layout="wide")

# MySQL 연결 설정
# 각 호출 시 새 커넥션을 생성하도록 변경
def init_connection():
    return mysql.connector.connect(
        host="tjmserver.iptime.org",
        user="exceluser",
        password="!crown2320A",
        database="excel_db"
    )

# DB 초기 설정 (히스토리 테이블 생성)
def setup_database():
    conn = init_connection()
    cursor = conn.cursor()
    # 백틱으로 식별자 감싸기
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS `history_branch` (
            `id` INT AUTO_INCREMENT PRIMARY KEY,
            `row_id` INT,
            `column_name` VARCHAR(255),
            `old_value` TEXT,
            `new_value` TEXT,
            `timestamp` DATETIME
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS `history_zone` (
            `id` INT AUTO_INCREMENT PRIMARY KEY,
            `row_id` INT,
            `column_name` VARCHAR(255),
            `old_value` TEXT,
            `new_value` TEXT,
            `timestamp` DATETIME
        )
        """
    )
    conn.commit()
    cursor.close()
    conn.close()

# 히스토리 테이블 생성 호출
setup_database()

# 앱 타이틀
st.title("엑셀 DB 관리 시스템")

# 사이드바 메뉴
menu = st.sidebar.radio("메뉴 선택", ["데이터 업로드", "영업소 관리", "배송구역 관리"] )

# 공통: 엑셀 업로드 후 MySQL 테이블로 저장
def load_and_save(file, table_name):
    conn = init_connection()
    cursor = conn.cursor()
    df = pd.read_excel(file, engine='openpyxl') if file.name.endswith('xlsx') else pd.read_excel(file)
    df = df.fillna('')
    # No 컬럼 정수화
    if 'No' in df.columns:
        df['No'] = pd.to_numeric(df['No'], errors='coerce').fillna(0).astype(int)
    # 우편번호 .0 제거
    if '우편번호' in df.columns:
        df['우편번호'] = df['우편번호'].astype(str).str.replace(r'\.0$', '', regex=True)
    cols = df.columns.tolist()
    # 기존 테이블 삭제 및 생성
    cursor.execute(f"DROP TABLE IF EXISTS `{table_name}`")
    defs = [f"`{c}` INT" if c=='No' else f"`{c}` TEXT" for c in cols]
    cursor.execute(f"CREATE TABLE `{table_name}` (id INT AUTO_INCREMENT PRIMARY KEY, {', '.join(defs)})")
    # 데이터 삽입
    for _, row in df.iterrows():
        vals = [int(row[c]) if c=='No' else str(row[c]) for c in cols]
        cols_sql = ','.join(f"`{c}`" for c in cols)
        placeholders = ','.join(['%s']*len(cols))
        cursor.execute(f"INSERT INTO `{table_name}` ({cols_sql}) VALUES ({placeholders})", vals)
    conn.commit()
    cursor.close()
    conn.close()
    # 데이터를 DataFrame으로 반환
    conn = init_connection()
    df_sql = pd.read_sql(f"SELECT * FROM `{table_name}`", conn, index_col='id')
    conn.close()
    return df_sql

# 페이지: 데이터 업로드
if menu == "데이터 업로드":
    st.header("데이터 업로드")
    bfile = st.file_uploader("영업소 엑셀 업로드", type=['xlsx','xls'], key='upload_branch')
    if bfile:
        df_b = load_and_save(bfile, 'branch_table')
        st.success("영업소 데이터 저장 완료")
    zfile = st.file_uploader("배송구역 엑셀 업로드", type=['xlsx','xls'], key='upload_zone')
    if zfile:
        df_z = load_and_save(zfile, 'delivery_zone_table')
        st.success("배송구역 데이터 저장 완료")

# 페이지: 영업소 관리
elif menu == "영업소 관리":
    st.header("영업소 데이터 관리")
    try:
        conn = init_connection()
        df_b = pd.read_sql("SELECT * FROM `branch_table`", conn, index_col='id')
        conn.close()
    except Exception:
        df_b = pd.DataFrame()
    if df_b.empty:
        st.warning("먼저 데이터를 업로드해주세요.")
    else:
        # 다운로드 버튼
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df_b.to_excel(w, index=False, sheet_name='영업소')
        buf.seek(0)
        st.download_button("영업소 엑셀 다운로드", buf, file_name=f"branch_{datetime.datetime.now():%Y%m%d}.xlsx")
        # 조회/수정 탭
        tab1, tab2 = st.tabs(["조회/수정", "신규 추가"])
        with tab1:
            term = st.text_input("검색어", key='b_search')
            sort = st.selectbox("정렬 컬럼", options=df_b.columns, key='b_sort')
            asc = st.checkbox("오름차순", value=True, key='b_asc')
            df_disp = df_b.copy()
            if term:
                df_disp = df_disp[df_disp.apply(lambda r: r.astype(str).str.contains(term, False).any(), axis=1)]
            # 숫자+문자 복합 정렬
            try:
                df_disp['_key'] = pd.to_numeric(df_disp[sort], errors='coerce')
                df_disp = df_disp.sort_values(by=['_key', sort], ascending=[asc,asc]).drop(columns=['_key'])
            except Exception:
                df_disp = df_disp.sort_values(by=sort, ascending=asc)
            edited = st.data_editor(df_disp, use_container_width=True)
            if st.button("변경사항 저장", key='save_branch'):
                conn = init_connection()
                cursor = conn.cursor()
                changes = 0
                for idx, row in edited.iterrows():
                    orig = df_b.loc[idx]
                    for c in df_b.columns:
                        if str(orig[c]) != str(row[c]):
                            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            cursor.execute("INSERT INTO `history_branch` (row_id, column_name, old_value, new_value, timestamp) VALUES (%s,%s,%s,%s,%s)",
                                           (int(idx), c, str(orig[c]), str(row[c]), ts))
                            cursor.execute(f"UPDATE `branch_table` SET `{c}`=%s WHERE id=%s", (row[c], idx))
                            changes += 1
                conn.commit()
                cursor.close()
                conn.close()
                st.success(f"{changes}개 셀 업데이트 완료")
        with tab2:
            st.subheader("신규 데이터 추가")
            cols = df_b.columns.tolist()
            new = {}
            for c in cols:
                if c=='No':
                    new[c] = st.number_input(c, value=int(df_b['No'].max()+1), step=1)
                else:
                    new[c] = st.text_input(c)
            if st.button("추가", key='add_branch'):
                conn = init_connection()
                cursor = conn.cursor()
                cols_sql = ','.join(f"`{c}`" for c in cols)
                placeholders = ','.join(['%s']*len(cols))
                vals = [int(new[c]) if c=='No' else new[c] for c in cols]
                cursor.execute(f"INSERT INTO `branch_table` ({cols_sql}) VALUES ({placeholders})", vals)
                new_id = cursor.lastrowid
                ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                for c in cols:
                    cursor.execute("INSERT INTO `history_branch` (row_id, column_name, old_value, new_value, timestamp) VALUES (%s,%s,%s,%s,%s)",
                                   (new_id, c, '[NEW]', str(new[c]), ts))
                conn.commit()
                cursor.close()
                conn.close()
                st.success("신규 데이터 추가 완료")

# 페이지: 배송구역 관리
elif menu == "배송구역 관리":
    st.header("배송구역 데이터 관리")
    try:
        conn = init_connection()
        df_z = pd.read_sql("SELECT * FROM `delivery_zone_table`", conn, index_col='id')
        conn.close()
    except Exception:
        df_z = pd.DataFrame()
    if df_z.empty:
        st.warning("먼저 데이터를 업로드해주세요.")
    else:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df_z.to_excel(w, index=False, sheet_name='배송구역')
        buf.seek(0)
        st.download_button("배송구역 엑셀 다운로드", buf, file_name=f"zone_{datetime.datetime.now():%Y%m%d}.xlsx")
        tab1, tab2 = st.tabs(["조회/수정", "신규 추가"])
        with tab1:
            term = st.text_input("검색어", key='z_search')
            sort = st.selectbox("정렬 컬럼", options=df_z.columns, key='z_sort')
            asc = st.checkbox("오름차순", value=True, key='z_asc')
            df_disp = df_z.copy()
            if term:
                df_disp = df_disp[df_disp.apply(lambda r: r.astype(str).str.contains(term, False).any(), axis=1)]
            try:
                df_disp['_key'] = pd.to_numeric(df_disp[sort], errors='coerce')
                df_disp = df_disp.sort_values(by=['_key', sort], ascending=[asc,asc]).drop(columns=['_key'])
            except Exception:
                df_disp = df_disp.sort_values(by=sort, ascending=asc)
            edited = st.data_editor(df_disp, use_container_width=True)
            if st.button("변경사항 저장", key='save_zone'):
                conn = init_connection()
                cursor = conn.cursor()
                changes = 0
                for idx, row in edited.iterrows():
                    orig = df_z.loc[idx]
                    for c in df_z.columns:
                        if str(orig[c]) != str(row[c]):
                            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            cursor.execute("INSERT INTO `history_zone` (row_id, column_name, old_value, new_value, timestamp) VALUES (%s,%s,%s,%s,%s)",
                                           (int(idx), c, str(orig[c]), str(row[c]), ts))
                            cursor.execute(f"UPDATE `delivery_zone_table` SET `{c}`=%s WHERE id=%s", (row[c], idx))
                            changes += 1
                conn.commit()
                cursor.close()
                conn.close()
                st.success(f"{changes}개 셀 업데이트 완료")
        with tab2:
            st.subheader("신규 데이터 추가")
            cols = df_z.columns.tolist()
            new = {}
            for c in cols:
                if c=='No':
                    new[c] = st.number_input(c, value=int(df_z['No'].max()+1), step=1)
                else:
                    new[c] = st.text_input(c)
            if st.button("추가", key='add_zone'):
                conn = init_connection()
                cursor = conn.cursor()
                cols_sql = ','.join(f"`{c}`" for c in cols)
                placeholders = ','.join(['%s']*len(cols))
                vals = [int(new[c]) if c=='No' else new[c] for c in cols]
                cursor.execute(f"INSERT INTO `delivery_zone_table` ({cols_sql}) VALUES ({placeholders})", vals)
                new_id = cursor.lastrowid
                ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                for c in cols:
                    cursor.execute("INSERT INTO `history_zone` (row_id, column_name, old_value, new_value, timestamp) VALUES (%s,%s,%s,%s,%s)",
                                   (new_id, c, '[NEW]', str(new[c]), ts))
                conn.commit()
                cursor.close()
                conn.close()
                st.success("신규 데이터 추가 완료")
