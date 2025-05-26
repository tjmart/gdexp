import streamlit as st
import pandas as pd
import os
import re
import openpyxl
import xlrd
import io
import mysql.connector
from mysql.connector import pooling
from sqlalchemy import create_engine, text
import numpy as np
import datetime
import xlsxwriter
import win32com.client
import tempfile
import pythoncom

# ====== 공통 유틸 ======
def clean_string(text):
    if not isinstance(text, str):
        return text
    text = re.sub(r'^[=+]+', '', text)
    text = text.rstrip('\\')
    return text

def get_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', options={
        'strings_to_numbers': True,
        'strings_to_urls': False,
        'strings_to_formulas': False
    }) as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.set_column(idx, idx, max_length + 2)
    output.seek(0)
    return output.getvalue()

# ====== 엑셀 오류 체크 ======
def excel_error_checker():
    st.header("엑셀 오류 체크 (자동 정리 및 표준 헤더 비교)")
    file_type = st.radio("파일 유형", ["도착 파일", "발송 파일"], key="error_checker_type")
    if file_type == "도착 파일":
        standard_headers = [
            "No", "알림", "발송접수일", "발송접수시간", "구분", "법인", "발송구분", "회수", "미착", "과착", "변상", "운송장번호",
            "고객사주문번호", "인수자타입", "인수자명", "인수완료일시", "접수계정", "발송지", "발송지전화번호", "도착지", "도착지전화번호",
            "보내는분", "보낸분전화번호", "보낸분기타전화번호", "보낸분우편번호", "보낸분주소", "보낸분상세주소", "받는분",
            "도착고객관리", "받는분전화번호", "받는분기타전화번호", "받는분우편번호", "받는분주소", "받는분상세주소", "품목명",
            "배송구분", "포장상태", "수량", "운임구분", "결재구분", "신용", "결재여부", "결재일시", "개별단가", "가로", "세로",
            "높이", "무게", "CBM", "할증운임", "배송운임", "도서운임", "기타운임", "별도운임", "운임합계", "출고비",
            "출고비결재수단", "출고비결재여부", "출고비결재일시", "메모", "노선번호", "발송연선번호", "도착연선번호", "발송지지역",
            "도착지지역", "발송지관할지역", "도착지관할지역", "발송터미널", "도착터미널", "발송터미널하차번호", "도착터미널하차번호"
        ]
    else:
        standard_headers = [
            "No", "알림", "발송접수일", "발송접수시간", "구분", "법인", "발송구분", "회수", "미착", "과착", "변상", "운송장번호",
            "고객사주문번호", "인수자타입", "인수자명", "인수완료일시", "접수계정", "발송지", "발송지전화번호", "도착지", "도착지전화번호",
            "보내는분", "고객관리번호", "발송고객관리", "보낸분전화번호", "보낸분기타전화번호", "보낸분우편번호", "보낸분주소",
            "보낸분상세주소", "받는분", "받는분전화번호", "받는분기타전화번호", "받는분우편번호", "받는분주소", "받는분상세주소",
            "품목명", "배송구분", "포장상태", "수량", "운임구분", "결재구분", "신용", "결재여부", "결재일시", "개별단가",
            "가로", "세로", "높이", "무게", "CBM", "할증운임", "배송운임", "도서운임", "기타운임", "별도운임", "운임합계",
            "메모", "노선번호", "발송연선번호", "도착연선번호", "발송지지역", "도착지지역", "발송지관할지역", "도착지관할지역",
            "발송터미널", "도착터미널", "발송터미널하차번호", "도착터미널하차번호", "픽업기사", "발송지수수료", "발송지운임삭제",
            "후불기타운임결재수단", "후불기타운임결재여부", "후불기타운임결재일시", "거래처 체크"
        ]
    uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx", "xls"], key="error_checker")
    if uploaded_file:
        try:
            pythoncom.CoInitialize()
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                temp_path = tmp_file.name
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                wb = excel.Workbooks.Open(temp_path)
                ws = wb.Sheets(1)
                total_cells = 0
                modified_cells = []
                modified_count = 0
                used_range = ws.UsedRange
                last_row = used_range.Row + used_range.Rows.Count - 1
                last_col = used_range.Column + used_range.Columns.Count - 1
                actual_headers = []
                for col in range(1, last_col + 1):
                    header = ws.Cells(1, col).Value
                    actual_headers.append(str(header) if header is not None else "")
                for row in range(2, last_row + 1):
                    for col in range(1, last_col + 1):
                        total_cells += 1
                        cell = ws.Cells(row, col)
                        val = cell.Value
                        # 문자열이든 수식이든 등호/플러스로 시작하면 무조건 삭제
                        if isinstance(val, str):
                            new_val = val.lstrip('=+').rstrip('\\₩')
                            if new_val != val:
                                cell.Value = new_val
                                modified_count += 1
                                coord = f"{ws.Cells(1, col).Value}{row}"
                                modified_cells.append((coord, val, new_val))
                        elif hasattr(cell, "HasFormula") and cell.HasFormula:
                            formula = cell.Formula
                            if isinstance(formula, str):
                                new_val = formula.lstrip('=+').rstrip('\\₩')
                                if new_val != formula:
                                    cell.Value = new_val
                                    modified_count += 1
                                    coord = f"{ws.Cells(1, col).Value}{row}"
                                    modified_cells.append((coord, formula, new_val))
                output_path = os.path.join(tempfile.gettempdir(), f"cleaned_{os.path.basename(uploaded_file.name)}")
                wb.SaveAs(output_path)
                st.success(f"{uploaded_file.name} 파일을 성공적으로 처리했습니다.")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("총 셀 수", total_cells)
                with col2:
                    st.metric("수정된 셀 수", modified_count)
                with col3:
                    st.metric("수정 비율", f"{(modified_count/total_cells*100):.1f}%")
                st.subheader("표준 헤더 비교 결과")
                header_status = []
                for idx, exp in enumerate(standard_headers):
                    act = actual_headers[idx] if idx < len(actual_headers) else ''
                    status = '일치' if exp == act else '불일치'
                    header_status.append((idx+1, exp, act, status))
                st.dataframe(
                    pd.DataFrame(header_status, columns=["순번", "표준", "실제", "상태"]),
                    use_container_width=True
                )
                if modified_cells:
                    st.subheader("수정된 셀 목록")
                    st.dataframe(
                        pd.DataFrame(modified_cells, columns=["셀 위치", "변경 전", "변경 후"]),
                        use_container_width=True
                    )
                with open(output_path, 'rb') as f:
                    st.download_button(
                        label="정리된 엑셀 다운로드",
                        data=f.read(),
                        file_name=f"cleaned_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            finally:
                wb.Close(SaveChanges=False)
                excel.Quit()
                os.unlink(temp_path)
                if os.path.exists(output_path):
                    os.unlink(output_path)
                pythoncom.CoUninitialize()
        except Exception as e:
            st.error(f"파일 처리 중 오류 발생: {str(e)}")
    else:
        st.info("엑셀 파일을 업로드해 주세요.")

# ====== 엑셀 파일 비교 ======
def compare_invoices(df_a, df_b):
    df_a['운송장번호'] = df_a['운송장번호'].astype(str)
    df_b['운송장번호'] = df_b['운송장번호'].astype(str)
    invoices_in_a = set(df_a['운송장번호'].unique())
    invoices_in_b = set(df_b['운송장번호'].unique())
    only_in_a = invoices_in_a - invoices_in_b
    only_in_b = invoices_in_b - invoices_in_a
    result_a_only = df_a[df_a['운송장번호'].isin(only_in_a)]
    result_b_only = df_b[df_b['운송장번호'].isin(only_in_b)]
    return result_a_only, result_b_only

def excel_file_compare():
    st.header("엑셀 파일 비교")
    col1, col2 = st.columns(2)
    with col1:
        file_a = st.file_uploader("A 파일 업로드", type=['xlsx', 'xls'], key="compare_a")
    with col2:
        file_b = st.file_uploader("B 파일 업로드", type=['xlsx', 'xls'], key="compare_b")

    if file_a is not None and file_b is not None:
        try:
            df_a = pd.read_excel(file_a)
            df_b = pd.read_excel(file_b)
            st.success("두 파일을 성공적으로 로드했습니다!")
            if '운송장번호' not in df_a.columns or '운송장번호' not in df_b.columns:
                st.error("하나 또는 두 파일 모두에 '운송장번호' 열이 없습니다. 확인 후 다시 시도해 주세요.")
            else:
                if st.button("비교 시작", key="compare_btn"):
                    result_a_only, result_b_only = compare_invoices(df_a, df_b)
                    tab1, tab2, tab3 = st.tabs(["A 파일에만 있는 운송장번호", "B 파일에만 있는 운송장번호", "요약"])
                    with tab1:
                        if not result_a_only.empty:
                            st.success(f"A 파일에만 {len(result_a_only)}개의 고유한 운송장번호가 있습니다.")
                            st.dataframe(result_a_only)
                            st.download_button(
                                label="A 파일 고유 데이터 다운로드 (Excel)",
                                data=get_excel_bytes(result_a_only),
                                file_name="A만_있는_운송장번호.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.info("A 파일에만 있는 운송장번호가 없습니다.")
                    with tab2:
                        if not result_b_only.empty:
                            st.success(f"B 파일에만 {len(result_b_only)}개의 고유한 운송장번호가 있습니다.")
                            st.dataframe(result_b_only)
                            st.download_button(
                                label="B 파일 고유 데이터 다운로드 (Excel)",
                                data=get_excel_bytes(result_b_only),
                                file_name="B만_있는_운송장번호.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.info("B 파일에만 있는 운송장번호가 없습니다.")
                    with tab3:
                        st.metric("A 파일 총 운송장번호", len(df_a))
                        st.metric("B 파일 총 운송장번호", len(df_b))
                        common_count = len(df_a) - len(result_a_only)
                        st.metric("공통 운송장번호", common_count)
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
    else:
        st.info("계속하려면 두 엑셀 파일을 모두 업로드하세요.")

# ====== 엑셀 DB 관리 시스템 (kdupdate_3.py 통합) ======
# (kdupdate_3.py 전체 코드 복사, 중복 import/설정/함수명 제거)
# 아래는 kdupdate_3.py의 모든 함수, DB설정, Streamlit UI, 메뉴 분기 전체 통합본입니다.

# 페이지 설정 (main_app.py에서 한 번만 호출)
st.set_page_config(page_title="통합 엑셀/DB 관리 시스템", layout="wide")

# MySQL 연결 설정
DB_CONFIG = {
    'host': st.secrets["mysql"]["host"],
    'port': 3306,
    'user': st.secrets["mysql"]["user"],
    'password': st.secrets["mysql"]["password"],
    'database': st.secrets["mysql"]["database"],
    'pool_name': 'mypool',
    'pool_size': 32,
    'pool_reset_session': True,
    'connect_timeout': 10
}

def get_branch_name_column(df):
    possible_names = ['영업소', '영업소명', '지점', '지점명', '사업소', '사업소명']
    for name in possible_names:
        if name in df.columns:
            return name
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]

@st.cache_resource
def init_engine():
    try:
        connection_string = f"mysql+pymysql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
        return create_engine(connection_string, pool_size=5, max_overflow=10)
    except Exception as err:
        st.error(f"데이터베이스 엔진 초기화 실패: {err}")
        return None

@st.cache_resource
def init_connection_pool():
    try:
        return mysql.connector.pooling.MySQLConnectionPool(**DB_CONFIG)
    except mysql.connector.Error as err:
        st.error(f"데이터베이스 연결 풀 초기화 실패: {err}")
        return None

def get_connection():
    pool = init_connection_pool()
    if pool is None:
        st.error("데이터베이스 연결 풀을 초기화할 수 없습니다.")
        return None
    try:
        conn = pool.get_connection()
        return conn
    except mysql.connector.Error as err:
        st.error(f"데이터베이스 연결 실패: {err}")
        st.cache_resource.clear()
        try:
            pool = init_connection_pool()
            if pool:
                return pool.get_connection()
        except:
            pass
        return None

class DatabaseConnection:
    def __init__(self):
        self.conn = None
        self.cursor = None
    def __enter__(self):
        self.conn = get_connection()
        if self.conn:
            self.cursor = self.conn.cursor()
            return self.cursor
        return None
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            try:
                if exc_type is None:
                    self.conn.commit()
                else:
                    self.conn.rollback()
            finally:
                self.conn.close()

def load_branch_data():
    engine = init_engine()
    if engine is None:
        return pd.DataFrame()
    try:
        with engine.connect() as connection:
            return pd.read_sql_table('branch_table', connection, index_col='id')
    except Exception as err:
        st.error(f"데이터 로딩 실패: {err}")
        return pd.DataFrame()

def load_zone_data():
    engine = init_engine()
    if engine is None:
        return pd.DataFrame()
    try:
        with engine.connect() as connection:
            return pd.read_sql_table('delivery_zone_table', connection, index_col='id')
    except Exception as err:
        st.error(f"데이터 로딩 실패: {err}")
        return pd.DataFrame()

def load_and_save(file, table_name):
    with DatabaseConnection() as cursor:
        if cursor is None:
            return pd.DataFrame()
        try:
            df = pd.read_excel(file, engine='openpyxl') if file.name.endswith('xlsx') else pd.read_excel(file)
            df = df.fillna('')
            if 'No' in df.columns:
                df['No'] = pd.to_numeric(df['No'], errors='coerce').fillna(0).astype(int)
            if '우편번호' in df.columns:
                df['우편번호'] = df['우편번호'].astype(str).str.replace(r'\.0$', '', regex=True)
            cols = df.columns.tolist()
            cursor.execute(f"DROP TABLE IF EXISTS `{table_name}`")
            defs = [f"`{c}` INT" if c=='No' else f"`{c}` TEXT" for c in cols]
            cursor.execute(f"CREATE TABLE `{table_name}` (id INT AUTO_INCREMENT PRIMARY KEY, {', '.join(defs)})")
            for _, row in df.iterrows():
                vals = [int(row[c]) if c=='No' else str(row[c]) for c in cols]
                cols_sql = ','.join(f"`{c}`" for c in cols)
                ph = ','.join(['%s']*len(cols))
                cursor.execute(f"INSERT INTO `{table_name}` ({cols_sql}) VALUES ({ph})", vals)
            st.cache_data.clear()
            st.rerun()
            return load_branch_data() if table_name == 'branch_table' else load_zone_data()
        except Exception as err:
            st.error(f"데이터 저장 실패: {err}")
            return pd.DataFrame()

def load_zone_data_paged(search_term="", page=1, page_size=100, branch_name=None):
    with DatabaseConnection() as cursor:
        if cursor is None:
            return pd.DataFrame(), 0
        try:
            cursor.execute("SHOW COLUMNS FROM `delivery_zone_table`")
            columns = [column[0] for column in cursor.fetchall()]
            zone_branch_col = next((col for col in columns if '영업' in col or '지점' in col or '사업' in col), None)
            where_clause = "WHERE 1=1"
            params = []
            if search_term and '우편번호' in columns:
                where_clause += " AND (우편번호 LIKE %s)"
                params.append(f"%{search_term}%")
            if branch_name and branch_name != "전체":
                with DatabaseConnection() as col_cursor:
                    col_cursor.execute("SHOW COLUMNS FROM `delivery_zone_table`")
                    all_columns = [col[0] for col in col_cursor.fetchall()]
                    branch_cols = [c for c in all_columns if any(x in c for x in ['영업', '지점', '사업'])]
                if branch_cols:
                    where_clause += " AND (" + " OR ".join([f"`{col}` LIKE %s" for col in branch_cols]) + ")"
                    params.extend([f"%{branch_name}%"] * len(branch_cols))
            cursor.execute(f"SELECT COUNT(*) FROM `delivery_zone_table` {where_clause}", params)
            total_count = cursor.fetchone()[0]
            offset = (page - 1) * page_size
            query = f"""
                SELECT * FROM `delivery_zone_table` 
                {where_clause}
                ORDER BY `No` 
                LIMIT %s OFFSET %s
            """
            params.extend([page_size, offset])
            cursor.execute(query, params)
            columns = [column[0] for column in cursor.description]
            data = cursor.fetchall()
            df = pd.DataFrame(data, columns=columns)
            if not df.empty and 'id' in df.columns:
                df.set_index('id', inplace=True)
            return df, total_count
        except Exception as err:
            st.error(f"데이터 로딩 실패: {err}")
            return pd.DataFrame(), 0

def add_new_row(table_name, data_dict):
    with DatabaseConnection() as cursor:
        if cursor is None:
            return None
        try:
            cols = list(data_dict.keys())
            vals = [int(data_dict[c]) if c=='No' else str(data_dict[c]) for c in cols]
            cols_sql = ','.join(f"`{c}`" for c in cols)
            ph = ','.join(['%s']*len(cols))
            cursor.execute(f"INSERT INTO `{table_name}` ({cols_sql}) VALUES ({ph})", vals)
            new_id = cursor.lastrowid
            st.cache_data.clear()
            st.rerun()
            return new_id
        except mysql.connector.Error as err:
            st.error(f"신규 행 추가 실패: {err}")
            return None

def setup_database():
    with DatabaseConnection() as cursor:
        if cursor is None:
            return
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS `history_branch` (
                id INT AUTO_INCREMENT PRIMARY KEY,
                row_id INT,
                column_name VARCHAR(255),
                old_value TEXT,
                new_value TEXT,
                timestamp DATETIME
            )
            """
        )
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS `history_zone` (
                id INT AUTO_INCREMENT PRIMARY KEY,
                row_id INT,
                column_name VARCHAR(255),
                old_value TEXT,
                new_value TEXT,
                timestamp DATETIME
            )
            """
        )

setup_database()

def delete_branch(branch_id, branch_name):
    conn = get_connection()
    if conn is None:
        st.error("DB 연결 실패")
        return False, "데이터베이스 연결 실패"
    try:
        cursor = conn.cursor()
        branch_id = int(branch_id)
        cursor.execute("SHOW COLUMNS FROM `delivery_zone_table`")
        columns = [col[0] for col in cursor.fetchall()]
        zone_branch_col = next((c for c in columns if '영업' in c or '지점' in c or '사업' in c), None)
        select_cols = ['id']
        if '우편번호' in columns:
            select_cols.append('`우편번호`')
        if '주소' in columns:
            select_cols.append('`주소`')
        select_cols_str = ', '.join(select_cols)
        cursor.execute(
            f"SELECT {select_cols_str} FROM `delivery_zone_table` WHERE `{zone_branch_col}` = %s",
            (branch_name,)
        )
        affected_zones = cursor.fetchall()
        cursor.execute("DELETE FROM `branch_table` WHERE id = %s", (branch_id,))
        if cursor.rowcount == 0:
            conn.rollback()
            cursor.close()
            conn.close()
            return False, "삭제된 데이터 없음"
        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cursor.execute(
            "INSERT INTO `history_branch` "
            "(row_id, column_name, old_value, new_value, timestamp) "
            "VALUES (%s, %s, %s, %s, %s)",
            (branch_id, "삭제", branch_name, "[DELETED]", ts)
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True, affected_zones
    except Exception as err:
        conn.rollback()
        st.error(f"영업소 삭제 실패: {err}")
        return False, str(err)

def update_zone_branch(zone_id, new_branch):
    conn = get_connection()
    if conn is None:
        st.error("DB 연결 실패")
        return False, "데이터베이스 연결 실패"
    try:
        cursor = conn.cursor()
        zone_id = int(zone_id)
        cursor.execute("SHOW COLUMNS FROM `delivery_zone_table`")
        columns = [column[0] for column in cursor.fetchall()]
        zone_branch_col = next((col for col in columns if '영업' in col or '지점' in col or '사업' in col), None)
        if not zone_branch_col:
            cursor.close()
            return False, "배송구역 테이블에서 영업소 컬럼을 찾을 수 없습니다."
        cursor.execute(f"SELECT `{zone_branch_col}` FROM `delivery_zone_table` WHERE id = %s", (zone_id,))
        old_branch = cursor.fetchone()[0]
        cursor.execute(f"UPDATE `delivery_zone_table` SET `{zone_branch_col}` = %s WHERE id = %s", (new_branch, zone_id))
        affected = cursor.rowcount
        if affected == 0:
            st.error("수정된 데이터가 없습니다. (id가 잘못 전달되었거나, 이미 수정됨)")
            conn.rollback()
            cursor.close()
            conn.close()
            return False, "수정된 데이터 없음"
        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cursor.execute(
            "INSERT INTO `history_zone` (row_id, column_name, old_value, new_value, timestamp) VALUES (%s, %s, %s, %s, %s)",
            (zone_id, zone_branch_col, old_branch, new_branch, ts)
        )
        conn.commit()
        cursor.close()
        st.cache_data.clear()
        st.rerun()
        return True, None
    except Exception as err:
        conn.rollback()
        st.error(f"배송구역 수정 실패: {err}")
        return False, str(err)
    finally:
        conn.close()

def update_zones_branch(zone_ids, new_branch, old_branch):
    conn = get_connection()
    if conn is None:
        st.error("DB 연결 실패")
        return False, "데이터베이스 연결 실패"
    try:
        cursor = conn.cursor()
        zone_ids = [int(zid) for zid in zone_ids]
        cursor.execute("SHOW COLUMNS FROM `delivery_zone_table`")
        columns = [column[0] for column in cursor.fetchall()]
        zone_branch_col = next((col for col in columns if '영업' in col or '지점' in col or '사업' in col), None)
        if not zone_branch_col:
            cursor.close()
            return False, "배송구역 테이블에서 영업소 컬럼을 찾을 수 없습니다."
        ids_str = ','.join(map(str, zone_ids))
        cursor.execute(f"UPDATE `delivery_zone_table` SET `{zone_branch_col}` = %s WHERE id IN ({ids_str})", (new_branch,))
        affected = cursor.rowcount
        if affected == 0:
            st.error("수정된 데이터가 없습니다. (id가 잘못 전달되었거나, 이미 수정됨)")
            conn.rollback()
            cursor.close()
            conn.close()
            return False, "수정된 데이터 없음"
        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for zone_id in zone_ids:
            cursor.execute(
                "INSERT INTO `history_zone` (row_id, column_name, old_value, new_value, timestamp) VALUES (%s, %s, %s, %s, %s)",
                (zone_id, zone_branch_col, old_branch, new_branch, ts)
            )
        conn.commit()
        cursor.close()
        st.cache_data.clear()
        st.rerun()
        return True, None
    except Exception as err:
        conn.rollback()
        st.error(f"배송구역 일괄 수정 실패: {err}")
        return False, str(err)
    finally:
        conn.close()

def delete_zones(zone_ids, branch_name):
    conn = get_connection()
    if conn is None:
        st.error("DB 연결 실패")
        return False, "데이터베이스 연결 실패"
    try:
        cursor = conn.cursor()
        zone_ids = [int(zid) for zid in zone_ids]
        ids_str = ','.join(map(str, zone_ids))
        cursor.execute(f"DELETE FROM `delivery_zone_table` WHERE id IN ({ids_str})")
        affected = cursor.rowcount
        if affected == 0:
            st.error("삭제된 데이터가 없습니다. (id가 잘못 전달되었거나, 이미 삭제됨)")
            conn.rollback()
            cursor.close()
            conn.close()
            return False, "삭제된 데이터 없음"
        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for zone_id in zone_ids:
            cursor.execute(
                "INSERT INTO `history_zone` (row_id, column_name, old_value, new_value, timestamp) VALUES (%s, %s, %s, %s, %s)",
                (zone_id, "삭제", branch_name, "[DELETED]", ts)
            )
        conn.commit()
        cursor.close()
        st.cache_data.clear()
        st.rerun()
        return True, None
    except Exception as err:
        conn.rollback()
        st.error(f"배송구역 일괄 삭제 실패: {err}")
        return False, str(err)
    finally:
        conn.close()

def load_zone_data_all(search_term="", branch_name=None):
    with DatabaseConnection() as cursor:
        if cursor is None:
            return pd.DataFrame()
        try:
            cursor.execute("SHOW COLUMNS FROM `delivery_zone_table`")
            columns = [column[0] for column in cursor.fetchall()]
            zone_branch_col = next((col for col in columns if '영업' in col or '지점' in col or '사업' in col), None)
            where_clause = "WHERE 1=1"
            params = []
            if search_term and '우편번호' in columns:
                where_clause += " AND (우편번호 LIKE %s)"
                params.append(f"%{search_term}%")
            if branch_name and branch_name != "전체":
                with DatabaseConnection() as col_cursor:
                    col_cursor.execute("SHOW COLUMNS FROM `delivery_zone_table`")
                    all_columns = [col[0] for col in col_cursor.fetchall()]
                    branch_cols = [c for c in all_columns if any(x in c for x in ['영업', '지점', '사업'])]
                if branch_cols:
                    where_clause += " AND (" + " OR ".join([f"`{col}` LIKE %s" for col in branch_cols]) + ")"
                    params.extend([f"%{branch_name}%"] * len(branch_cols))
            query = f"""
                SELECT * FROM `delivery_zone_table` 
                {where_clause}
                ORDER BY `No` 
            """
            cursor.execute(query, params)
            columns = [column[0] for column in cursor.description]
            data = cursor.fetchall()
            df = pd.DataFrame(data, columns=columns)
            if not df.empty and 'id' in df.columns:
                df.set_index('id', inplace=True)
            return df
        except Exception as err:
            st.error(f"전체 데이터 로딩 실패: {err}")
            return pd.DataFrame()

def excel_db_manager():
    st.title("엑셀 DB 관리 시스템")
    menu = st.sidebar.radio("메뉴 선택", ["데이터 업로드", "영업소 관리", "배송구역 관리"], key="db_manager_menu")
    if menu == "데이터 업로드":
        st.header("데이터 업로드")
        bfile = st.file_uploader("영업소 엑셀 업로드", type=['xlsx','xls'], key="branch_upload")
        if bfile:
            df_b = load_and_save(bfile, 'branch_table')
            st.success("영업소 데이터 저장 완료")
        zfile = st.file_uploader("배송구역 엑셀 업로드", type=['xlsx','xls'], key="zone_upload")
        if zfile:
            df_z = load_and_save(zfile, 'delivery_zone_table')
            st.success("배송구역 데이터 저장 완료")
    elif menu == "영업소 관리":
        st.header("영업소 데이터 관리")
        df_b = load_branch_data()
        if df_b.empty:
            st.warning("먼저 데이터를 업로드해주세요.")
        else:
            branch_col = get_branch_name_column(df_b)
            # 'No' 컬럼이 있으면 오름차순 정렬
            df_b_download = df_b.copy()
            if 'No' in df_b_download.columns:
                df_b_download = df_b_download.sort_values(by='No', ascending=True)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
                df_b_download.to_excel(w, index=False, sheet_name='영업소')
                worksheet = w.sheets['영업소']
                for idx, col in enumerate(df_b_download.columns):
                    max_length = max(
                        df_b_download[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.set_column(idx, idx, max_length + 2)
            buf.seek(0)
            st.download_button("영업소 엑셀 다운로드", buf, file_name=f"branch_{datetime.datetime.now():%Y%m%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            tab1, tab2, tab3 = st.tabs(["조회/수정", "신규 추가", "영업소 삭제"])
            with tab1:
                term = st.text_input("검색어", key='b_search')
                sort = st.selectbox("정렬 컬럼", options=df_b.columns, key='b_sort')
                asc = st.checkbox("오름차순", value=True, key='b_asc')
                disp = df_b.copy()
                if term:
                    disp = disp[disp.apply(lambda r: r.astype(str).str.contains(term, False).any(), axis=1)]
                try:
                    disp['_k'] = pd.to_numeric(disp[sort], errors='coerce')
                    disp = disp.sort_values(by=['_k', sort], ascending=[asc,asc]).drop(columns=['_k'])
                except:
                    disp = disp.sort_values(by=sort, ascending=asc)
                edited = st.data_editor(disp, use_container_width=True, height=400)
                if st.button("변경사항 저장", key='save_branch'):
                    conn = get_connection()
                    if conn:
                        cur = conn.cursor()
                        changes = 0
                        for idx, row in edited.iterrows():
                            orig = df_b.loc[idx]
                            for c in df_b.columns:
                                if str(orig[c]) != str(row[c]):
                                    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                    cur.execute(
                                        "INSERT INTO `history_branch` (row_id,column_name,old_value,new_value,timestamp) VALUES (%s,%s,%s,%s,%s)",
                                        (int(idx), c, str(orig[c]), str(row[c]), ts)
                                    )
                                    cur.execute(f"UPDATE `branch_table` SET `{c}`=%s WHERE id=%s", (row[c], idx))
                                    changes += 1
                        conn.commit()
                        cur.close()
                        conn.close()
                        st.success(f"{changes}개 셀 업데이트 완료")
            with tab2:
                st.subheader("신규 데이터 추가")
                cols = df_b.columns.tolist()
                new = {}
                cols_input = st.columns(2)
                for i, c in enumerate(cols):
                    with cols_input[i%2]:
                        if c == 'No':
                            new[c] = st.number_input(c, value=int(df_b['No'].max()+1), step=1)
                        else:
                            new[c] = st.text_input(c)
                if st.button("추가", key='add_branch'):
                    new_id = add_new_row('branch_table', new)
                    if new_id:
                        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        conn = get_connection()
                        if conn:
                            cur = conn.cursor()
                            for c, v in new.items():
                                cur.execute(
                                    "INSERT INTO `history_branch` (row_id,column_name,old_value,new_value,timestamp) VALUES (%s,%s,%s,%s,%s)",
                                    (new_id, c, '[NEW]', str(v), ts)
                                )
                            conn.commit()
                            cur.close()
                            conn.close()
                            st.success(f"새 영업소 데이터 추가 완료 (ID: {new_id})")
            with tab3:
                st.subheader("영업소 삭제")
                st.warning("⚠️ 주의: 영업소를 삭제하면 관련된 배송구역을 재할당해야 합니다.")
                branch_to_delete = st.selectbox(
                    "삭제할 영업소",
                    options=df_b.index.tolist(),
                    format_func=lambda x: f"{df_b.loc[x, branch_col]} (ID: {x})"
                )
                confirm = st.checkbox("정말로 이 영업소를 삭제하시겠습니까?", key="confirm_delete")
                if confirm and st.button("영업소 삭제", key='delete_branch'):
                    success, result = delete_branch(branch_to_delete, df_b.loc[branch_to_delete, branch_col])
                    if success:
                        st.success("영업소가 삭제되었습니다.")
                        if result:
                            st.warning(f"이 영업소와 연결된 {len(result)}개의 배송구역이 있습니다.")
                            remaining = df_b[df_b.index != branch_to_delete][branch_col].tolist()
                            for zone_id, post, addr in result:
                                st.write(f"- 우편번호: {post}, 주소: {addr}")
                                new_branch = st.selectbox(f"새 영업소 ({post})", options=remaining, key=f"reassign_{zone_id}")
                                if st.button("재할당 적용", key=f"apply_{zone_id}"):
                                    ok, err = update_zone_branch(zone_id, new_branch)
                                    if ok:
                                        st.success("배송구역 재할당 완료")
                                    else:
                                        st.error(f"재할당 실패: {err}")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(f"삭제 실패: {result}")
    elif menu == "배송구역 관리":
        st.header("배송구역 데이터 관리")
        branches = []
        try:
            with DatabaseConnection() as cursor:
                if cursor:
                    cursor.execute("SHOW COLUMNS FROM `branch_table`")
                    branch_columns = [col[0] for col in cursor.fetchall()]
                    branch_col = next((c for c in branch_columns if '영업' in c or '지점' in c or '사업' in c), branch_columns[1] if len(branch_columns) > 1 else branch_columns[0])
                    cursor.execute(f"SELECT DISTINCT `{branch_col}` FROM `branch_table` ORDER BY `{branch_col}`")
                    branches = [r[0] for r in cursor.fetchall() if r[0]]
        except Exception as e:
            st.warning(f"영업소 목록을 불러올 수 없습니다: {e}")
        col1, col2, col3 = st.columns([1, 1, 0.5])
        with col1:
            search_term_input = st.text_input("주소/우편번호", key='z_search')
        with col2:
            branch_input = st.text_input("영업소명", value="", key='z_branch_input', placeholder="영업소명을 입력하거나 일부만 입력하세요")
        with col3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("검색", key='zone_search_btn'):
                st.session_state['zone_search_term'] = search_term_input
                # 입력값이 있으면 그대로 branch_name에 저장 (부분일치 필터)
                if branch_input.strip() == "":
                    st.session_state['zone_search_branch'] = "전체"
                else:
                    st.session_state['zone_search_branch'] = branch_input.strip()
                st.session_state['zone_page'] = 1
                st.session_state['zone_search_started'] = True  # 조회 시작
        if 'zone_search_term' not in st.session_state:
            st.session_state['zone_search_term'] = ''
        if 'zone_search_branch' not in st.session_state:
            st.session_state['zone_search_branch'] = '전체'
        if 'zone_search_started' not in st.session_state:
            st.session_state['zone_search_started'] = False
        search_term = st.session_state['zone_search_term']
        selected_branch = st.session_state['zone_search_branch']
        branch_filter = None if selected_branch == "전체" else selected_branch
        page_size = 50
        if 'zone_page' not in st.session_state:
            st.session_state.zone_page = 1
        if not st.session_state['zone_search_started']:
            st.info("검색 조건을 입력하고 [검색] 버튼을 눌러주세요.")
        else:
            df_z, total_count = load_zone_data_paged(search_term, st.session_state.zone_page, page_size, branch_filter)
            total_pages = (total_count + page_size - 1) // page_size
            col1, col2, col3 = st.columns([1,3,1])
            with col1:
                if st.button("이전", disabled=st.session_state.zone_page <= 1):
                    st.session_state.zone_page -= 1
                    st.rerun()
            with col2:
                st.write(f"페이지 {st.session_state.zone_page} / {total_pages} (총 {total_count}개)")
            with col3:
                if st.button("다음", disabled=st.session_state.zone_page >= total_pages):
                    st.session_state.zone_page += 1
                    st.rerun()
            if df_z.empty:
                st.warning("검색 결과가 없습니다.")
            else:
                # 배송구역 엑셀 다운로드 버튼: 전체 데이터로 변경
                df_z_all = load_zone_data_all(search_term, branch_filter)
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    df_z_all.to_excel(w, index=False, sheet_name='배송구역')
                buf.seek(0)
                st.download_button("배송구역 엑셀 다운로드", buf, file_name=f"zone_{datetime.datetime.now():%Y%m%d}.xlsx")
                tab1, tab2, tab3 = st.tabs(["조회/수정", "신규 추가", "일괄 처리"])
                with tab1:
                    edited = st.data_editor(
                        df_z,
                        use_container_width=True,
                        height=400,
                        num_rows="dynamic",
                        key='zone_editor'
                    )
                    if st.button("변경사항 저장", key='save_zone'):
                        conn = get_connection()
                        if conn:
                            try:
                                cur = conn.cursor()
                                changes = 0
                                for idx, row in edited.iterrows():
                                    orig = df_z.loc[idx]
                                    for c in df_z.columns:
                                        if str(orig[c]) != str(row[c]):
                                            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                            cur.execute(
                                                "INSERT INTO `history_zone` (row_id, column_name, old_value, new_value, timestamp) "
                                                "VALUES (%s, %s, %s, %s, %s)",
                                                (int(idx), c, str(orig[c]), str(row[c]), ts)
                                            )
                                            cur.execute(
                                                f"UPDATE `delivery_zone_table` SET `{c}` = %s WHERE id = %s",
                                                (row[c], idx)
                                            )
                                            changes += 1
                                conn.commit()
                                st.success(f"{changes}개 셀 업데이트 완료")
                            except Exception as e:
                                conn.rollback()
                                st.error(f"저장 중 오류 발생: {e}")
                            finally:
                                cur.close()
                                conn.close()
                                st.cache_data.clear()
                                st.rerun()
                with tab2:
                    st.subheader("신규 데이터 추가")
                    cols = df_z.columns.tolist()
                    new = {}
                    cols_input = st.columns(2)
                    for i, c in enumerate(cols):
                        with cols_input[i%2]:
                            if c == 'No':
                                with DatabaseConnection() as cursor:
                                    if cursor:
                                        cursor.execute("SELECT MAX(`No`) FROM `delivery_zone_table`")
                                        max_no = cursor.fetchone()[0] or 0
                                    else:
                                        max_no = 0
                                new[c] = st.number_input(c, value=max_no+1, step=1)
                            else:
                                new[c] = st.text_input(c)
                    if st.button("추가", key='add_zone'):
                        new_id = add_new_row('delivery_zone_table', new)
                        if new_id:
                            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            with DatabaseConnection() as cursor:
                                if cursor:
                                    for c, v in new.items():
                                        cursor.execute(
                                            "INSERT INTO `history_zone` (row_id,column_name,old_value,new_value,timestamp) VALUES (%s,%s,%s,%s,%s)",
                                            (new_id, c, '[NEW]', str(v), ts)
                                        )
                                    st.success(f"새 배송구역 데이터 추가 완료 (ID: {new_id})")
                                    st.cache_data.clear()
                                    st.rerun()
                with tab3:
                    st.subheader("일괄 처리")
                    address_col = next((c for c in df_z.columns if '주소' in c), None)
                    postal_col = next((c for c in df_z.columns if '우편' in c), None)
                    def zone_label(x):
                        postal = df_z.loc[x, postal_col] if postal_col else str(x)
                        address = df_z.loc[x, address_col] if address_col else ""
                        return f"{postal} - {address}"
                    if branch_filter:
                        action = st.radio("작업 선택", ["영업소 변경", "삭제"])
                        selected = st.multiselect(
                            "처리할 배송구역 선택",
                            options=df_z.index.tolist(),
                            format_func=zone_label
                        )
                        if action == "영업소 변경":
                            new_branch = st.selectbox(
                                "새로운 영업소 선택",
                                [b for b in branches if b != branch_filter]
                            )
                            if st.button("선택한 배송구역 영업소 변경", disabled=not selected):
                                success, error = update_zones_branch(selected, new_branch, branch_filter)
                                if success:
                                    st.success(f"{len(selected)}개 배송구역의 영업소를 변경했습니다.")
                                    st.cache_data.clear()
                                    st.rerun()
                                else:
                                    st.error(f"변경 실패: {error}")
                        else:
                            if st.button("선택한 배송구역 삭제", disabled=not selected):
                                if st.checkbox("정말로 삭제하시겠습니까?"):
                                    success, error = delete_zones(selected, branch_filter)
                                    if success:
                                        st.success(f"{len(selected)}개 배송구역을 삭제했습니다.")
                                        st.cache_data.clear()
                                        st.rerun()
                                    else:
                                        st.error(f"삭제 실패: {error}")
                    else:
                        st.info("일괄 처리를 위해 먼저 영업소를 선택해주세요.")

# ====== 메인 메뉴 ======
st.sidebar.title("메인 메뉴")
menu = st.sidebar.radio("기능 선택", [
    "엑셀 오류 체크", "엑셀 파일 비교", "엑셀 DB 관리 시스템"
], key="main_menu")

if menu == "엑셀 오류 체크":
    excel_error_checker()
elif menu == "엑셀 파일 비교":
    excel_file_compare()
elif menu == "엑셀 DB 관리 시스템":
    excel_db_manager() 