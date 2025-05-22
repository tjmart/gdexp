import streamlit as st
import pandas as pd
import sqlite3
import io
import datetime
import re
import numpy as np

# DB 연결 및 테이블 생성
conn = sqlite3.connect('data.db', check_same_thread=False)
cursor = conn.cursor()
# 히스토리 테이블 생성
cursor.execute('''
CREATE TABLE IF NOT EXISTS history_branch (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    row_id INTEGER,
    column_name TEXT,
    old_value TEXT,
    new_value TEXT,
    timestamp TEXT
)
''')
cursor.execute('''
CREATE TABLE IF NOT EXISTS history_zone (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    row_id INTEGER,
    column_name TEXT,
    old_value TEXT,
    new_value TEXT,
    timestamp TEXT
)
''')
conn.commit()

st.set_page_config(page_title="엑셀 DB 관리 시스템", layout="wide")
st.title("엑셀 DB 관리 시스템")

# 사이드바에 메뉴 추가
st.sidebar.title("메뉴")
menu = st.sidebar.radio(
    "작업 선택",
    ["데이터 업로드", "영업소 관리", "배송구역 관리"]
)

# --- 업로드 및 DB 저장 함수 ---
def load_and_save(file, table_name):
    df = pd.read_excel(file, engine='openpyxl') if file.name.endswith('xlsx') else pd.read_excel(file)
    df = df.fillna('')
    
    # No 컬럼이 있으면 정수형으로 변환 (No 컬럼을 올바르게 정렬하기 위함)
    if 'No' in df.columns:
        df['No'] = pd.to_numeric(df['No'], errors='coerce').fillna(0).astype(int)
    
    if '우편번호' in df.columns:
        df['우편번호'] = df['우편번호'].astype(str).str.replace(r'\.0$', '', regex=True)
    
    cols = df.columns.tolist()
    cursor.execute(f'DROP TABLE IF EXISTS {table_name}')
    
    # No 컬럼은 INTEGER로 저장, 나머지는 TEXT로 저장
    col_defs = []
    for c in cols:
        if c == 'No':
            col_defs.append(f'"{c}" INTEGER')
        else:
            col_defs.append(f'"{c}" TEXT')
    
    col_defs_str = ','.join(col_defs)
    cursor.execute(f'CREATE TABLE {table_name} (id INTEGER PRIMARY KEY AUTOINCREMENT, {col_defs_str})')
    
    for _, row in df.iterrows():
        vals = []
        for c in cols:
            if c == 'No':
                vals.append(int(row[c]) if pd.notna(row[c]) else 0)
            else:
                vals.append(str(row[c]))
        
        placeholders = ','.join('?' for _ in cols)
        col_list = ','.join(f'"{c}"' for c in cols)
        cursor.execute(f'INSERT INTO {table_name} ({col_list}) VALUES ({placeholders})', vals)
    
    conn.commit()
    return pd.read_sql(f'SELECT * FROM {table_name}', conn, index_col='id')

# 신규 데이터 추가 함수
def add_new_row(table_name, data_dict):
    cols = list(data_dict.keys())
    vals = []
    
    for c in cols:
        if c == 'No':
            try:
                vals.append(int(data_dict[c]))
            except:
                vals.append(0)
        else:
            vals.append(str(data_dict[c]))
    
    placeholders = ','.join('?' for _ in cols)
    col_list = ','.join(f'"{c}"' for c in cols)
    
    cursor.execute(f'INSERT INTO {table_name} ({col_list}) VALUES ({placeholders})', vals)
    conn.commit()
    
    # 마지막으로 추가된 행의 ID 반환
    return cursor.lastrowid

# 테이블 구조 가져오기
def get_table_columns(table_name):
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = [row[1] for row in cursor.fetchall() if row[1] != 'id']
    return columns

# -------------------- 페이지: 데이터 업로드 --------------------
if menu == "데이터 업로드":
    st.header("데이터 업로드")
    
    branch_file = st.file_uploader("영업소 업데이트 엑셀 업로드", type=['xlsx','xls'], key='branch')
    if branch_file:
        df_branch = load_and_save(branch_file, 'branch_table')
        st.success("영업소 DB 저장 완료")
        
    zone_file = st.file_uploader("배송구역 업데이트 엑셀 업로드", type=['xlsx','xls'], key='zone')
    if zone_file:
        df_zone = load_and_save(zone_file, 'delivery_zone_table')
        st.success("배송구역 DB 저장 완료")

# -------------------- 페이지: 영업소 관리 --------------------
elif menu == "영업소 관리":
    st.header("영업소 데이터 관리")

    # 영업소 데이터 불러오기
    try:
        df_branch = pd.read_sql('SELECT * FROM branch_table', conn, index_col='id')
        if df_branch.empty:
            st.warning("영업소 데이터가 없습니다. 데이터를 먼저 업로드해주세요.")
    except:
        st.warning("영업소 데이터가 없습니다. 데이터를 먼저 업로드해주세요.")
        df_branch = pd.DataFrame()
    
    if not df_branch.empty:
        # 탭 구성: 조회/수정, 추가
        tab1, tab2 = st.tabs(["조회 및 수정", "신규 추가"])
        
        # 탭 1: 조회 및 수정
        with tab1:
            # 다운로드 버튼
            buf_b = io.BytesIO()
            with pd.ExcelWriter(buf_b, engine='openpyxl') as writer:
                df_branch.to_excel(writer, index=False, sheet_name='영업소')
            buf_b.seek(0)
            st.download_button(
                "영업소 엑셀 다운로드",
                buf_b,
                file_name=f"영업소_업데이트_{datetime.datetime.now():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='dl_branch'
            )

            # 검색 및 정렬 옵션
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1:
                term_b = st.text_input("영업소 검색", key='search_branch')
            with col2:
                sort_col_b = st.selectbox("정렬 컬럼", options=df_branch.columns.tolist(), key='sort_branch')
            with col3:
                order_b = st.radio("정렬 순서", ['오름차순','내림차순'], key='order_branch')
            asc_b = (order_b == '오름차순')

            disp_b = df_branch.copy()
            if term_b:
                mask = disp_b.apply(lambda r: r.astype(str).str.contains(term_b, case=False).any(), axis=1)
                disp_b = disp_b[mask]
            
            # 'No' 컬럼이 있으면 숫자로 정렬, 아니면 선택한 컬럼으로 정렬
            if not disp_b.empty:
                if sort_col_b == 'No' and 'No' in disp_b.columns:
                    # No 컬럼은 이미 숫자로 저장되어 있으므로 직접 정렬
                    disp_b = disp_b.sort_values(by='No', ascending=asc_b)
                else:
                    # 다른 컬럼은 기존 방식으로 정렬
                    # 숫자 변환 시도
                    try:
                        # 먼저 숫자로 변환 시도
                        disp_b['__temp_sort_key'] = pd.to_numeric(disp_b[sort_col_b], errors='coerce')
                        # 변환 실패한 경우를 위해 원래 값도 보존
                        disp_b['__temp_sort_key_str'] = disp_b[sort_col_b]
                        # 숫자 + 문자열 복합 정렬
                        disp_b = disp_b.sort_values(
                            by=['__temp_sort_key', '__temp_sort_key_str'], 
                            ascending=[asc_b, asc_b], 
                            na_position='last'
                        )
                        # 임시 정렬 키 제거
                        disp_b = disp_b.drop(['__temp_sort_key', '__temp_sort_key_str'], axis=1)
                    except:
                        # 숫자 변환 실패 시 일반 정렬
                        disp_b = disp_b.sort_values(by=sort_col_b, ascending=asc_b)

            # 선택된 행 삭제 기능
            if not disp_b.empty:
                # 선택할 행의 인덱스를 표시
                st.write("행 선택 (삭제할 행 선택):")
                selected_indices = []
                
                # 10개씩 데이터 보기 (페이징)
                items_per_page = 10
                total_pages = max(1, (len(disp_b) + items_per_page - 1) // items_per_page)
                current_page = st.number_input("페이지", min_value=1, max_value=total_pages, value=1) - 1  # 0-인덱싱을 위해 1 빼기
                
                start_idx = current_page * items_per_page
                end_idx = min(start_idx + items_per_page, len(disp_b))
                
                page_indices = list(disp_b.index)[start_idx:end_idx]
                
                # 체크박스로 행 선택 - 접근성 문제 해결
                for idx in page_indices:
                    row = disp_b.loc[idx]
                    col1, col2 = st.columns([1, 9])
                    with col1:
                        # 체크박스에 레이블 추가 (숨김 처리)
                        if st.checkbox(f"항목 {idx}", key=f"select_{idx}", label_visibility="collapsed"):
                            selected_indices.append(idx)
                    with col2:
                        # 행 정보 간단히 표시
                        display_info = " | ".join([f"{col}: {row[col]}" for col in disp_b.columns[:3]])
                        st.write(f"ID: {idx}, {display_info}")
                
                # 삭제 버튼
                if selected_indices and st.button("선택한 행 삭제", key="delete_branch"):
                    for idx in selected_indices:
                        # 이력 기록 (전체 행 삭제)
                        for col in df_branch.columns:
                            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            cursor.execute(
                                "INSERT INTO history_branch (row_id,column_name,old_value,new_value,timestamp) VALUES (?,?,?,?,?)",
                                (idx, col, str(disp_b.loc[idx, col]), "[DELETED]", ts)
                            )
                        # DB에서 삭제
                        cursor.execute(f'DELETE FROM branch_table WHERE id=?', (idx,))
                    
                    conn.commit()
                    st.success(f"{len(selected_indices)}개 행이 삭제되었습니다.")
                    # 데이터 리로드
                    df_branch = pd.read_sql('SELECT * FROM branch_table', conn, index_col='id')
                    st.experimental_rerun()

            # 데이터 에디터로 편집
            edited_b = st.data_editor(
                disp_b,
                num_rows="dynamic",
                use_container_width=True,
                key='editor_branch',
                height=400
            )
            
            # 변경 사항 저장 버튼
            if st.button("변경사항 저장", key="save_branch_changes"):
                if not edited_b.equals(disp_b):
                    changes_count = 0
                    for idx in edited_b.index:
                        if idx in disp_b.index:  # 인덱스가 있는지 확인
                            old = disp_b.loc[idx]
                            new = edited_b.loc[idx]
                            for col in df_branch.columns:
                                if str(old[col]) != str(new[col]):
                                    changes_count += 1
                                    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                    cursor.execute(
                                        "INSERT INTO history_branch (row_id,column_name,old_value,new_value,timestamp) VALUES (?,?,?,?,?)",
                                        (idx, col, str(old[col]), str(new[col]), ts)
                                    )
                                    
                                    # No 컬럼인 경우 정수형으로 저장
                                    if col == 'No':
                                        try:
                                            value = int(float(new[col]))
                                        except:
                                            value = 0
                                        cursor.execute(f'UPDATE branch_table SET "{col}"=? WHERE id=?', (value, idx))
                                    else:
                                        cursor.execute(f'UPDATE branch_table SET "{col}"=? WHERE id=?', (str(new[col]), idx))
                    
                    conn.commit()
                    if changes_count > 0:
                        st.success(f"영업소 데이터 {changes_count}개 셀 업데이트 완료")
                    else:
                        st.info("변경된 내용이 없습니다.")
                    
                    # 데이터 리로드
                    df_branch = pd.read_sql('SELECT * FROM branch_table', conn, index_col='id')
                else:
                    st.info("변경된 내용이 없습니다.")
        
        # 탭 2: 신규 추가
        with tab2:
            st.subheader("영업소 신규 데이터 추가")
            
            # 컬럼 정보 가져오기
            branch_cols = get_table_columns('branch_table')
            
            # 새 데이터 입력 폼
            new_branch_data = {}
            
            # 컬럼별 입력 필드 표시 (2열로 배치)
            cols = st.columns(2)
            for i, col in enumerate(branch_cols):
                with cols[i % 2]:
                    if col == 'No':
                        # No 컬럼은 기본값으로 최대 값 + 1 제공
                        max_no = 0
                        if 'No' in df_branch.columns and not df_branch.empty:
                            max_no = df_branch['No'].max()
                        new_branch_data[col] = st.number_input(f"{col}", value=max_no+1, step=1, key=f"new_{col}")
                    else:
                        new_branch_data[col] = st.text_input(f"{col}", key=f"new_{col}")
            
            # 추가 버튼
            if st.button("데이터 추가", key="add_branch"):
                # 필수 입력 체크 (여기서는 No 컬럼만 필수로 가정)
                if 'No' in new_branch_data and new_branch_data['No']:
                    try:
                        # 데이터 추가
                        new_id = add_new_row('branch_table', new_branch_data)
                        
                        # 이력 기록 (신규 추가)
                        for col, value in new_branch_data.items():
                            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            cursor.execute(
                                "INSERT INTO history_branch (row_id,column_name,old_value,new_value,timestamp) VALUES (?,?,?,?,?)",
                                (new_id, col, "[NEW]", str(value), ts)
                            )
                        
                        conn.commit()
                        st.success(f"새 영업소 데이터가 추가되었습니다. (ID: {new_id})")
                        
                        # 폼 초기화 및 데이터 리로드
                        df_branch = pd.read_sql('SELECT * FROM branch_table', conn, index_col='id')
                        st.experimental_rerun()
                    except Exception as e:
                        # 오류 발생 시 롤백 및 에러 메시지 표시
                        conn.rollback()
                        st.error(f"데이터 추가 중 오류가 발생했습니다: {str(e)}")
                else:
                    st.error("No 필드는 필수 입력입니다.")
    else:
        st.info("영업소 데이터가 없습니다. 먼저 데이터를 업로드해주세요.")

# -------------------- 페이지: 배송구역 관리 --------------------
elif menu == "배송구역 관리":
    st.header("배송구역 데이터 관리")

    # 배송구역 데이터 불러오기
    try:
        df_zone = pd.read_sql('SELECT * FROM delivery_zone_table', conn, index_col='id')
        if df_zone.empty:
            st.warning("배송구역 데이터가 없습니다. 데이터를 먼저 업로드해주세요.")
    except:
        st.warning("배송구역 데이터가 없습니다. 데이터를 먼저 업로드해주세요.")
        df_zone = pd.DataFrame()
    
    if not df_zone.empty:
        # 탭 구성: 조회/수정, 추가
        tab1, tab2 = st.tabs(["조회 및 수정", "신규 추가"])
        
        # 탭 1: 조회 및 수정
        with tab1:
            # 다운로드 버튼
            buf_z = io.BytesIO()
            with pd.ExcelWriter(buf_z, engine='openpyxl') as writer:
                df_zone.to_excel(writer, index=False, sheet_name='배송구역')
            buf_z.seek(0)
            st.download_button(
                "배송구역 엑셀 다운로드",
                buf_z,
                file_name=f"배송구역_업데이트_{datetime.datetime.now():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='dl_zone'
            )

            # 검색 및 정렬 옵션
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1:
                term_z = st.text_input("배송구역 검색", key='search_zone')
            with col2:
                sort_col_z = st.selectbox("정렬 컬럼", options=df_zone.columns.tolist(), key='sort_zone')
            with col3:
                order_z = st.radio("정렬 순서", ['오름차순','내림차순'], key='order_zone')
            asc_z = (order_z == '오름차순')

            disp_z = df_zone.copy()
            if term_z:
                mask = disp_z.apply(lambda r: r.astype(str).str.contains(term_z, case=False).any(), axis=1)
                disp_z = disp_z[mask]
            
            # 'No' 컬럼이 있으면 숫자로 정렬, 아니면 선택한 컬럼으로 정렬
            if not disp_z.empty:
                if sort_col_z == 'No' and 'No' in disp_z.columns:
                    # No 컬럼은 이미 숫자로 저장되어 있으므로 직접 정렬
                    disp_z = disp_z.sort_values(by='No', ascending=asc_z)
                else:
                    # 다른 컬럼은 기존 방식으로 정렬
                    # 숫자 변환 시도
                    try:
                        # 먼저 숫자로 변환 시도
                        disp_z['__temp_sort_key'] = pd.to_numeric(disp_z[sort_col_z], errors='coerce')
                        # 변환 실패한 경우를 위해 원래 값도 보존
                        disp_z['__temp_sort_key_str'] = disp_z[sort_col_z]
                        # 숫자 + 문자열 복합 정렬
                        disp_z = disp_z.sort_values(
                            by=['__temp_sort_key', '__temp_sort_key_str'], 
                            ascending=[asc_z, asc_z], 
                            na_position='last'
                        )
                        # 임시 정렬 키 제거
                        disp_z = disp_z.drop(['__temp_sort_key', '__temp_sort_key_str'], axis=1)
                    except:
                        # 숫자 변환 실패 시 일반 정렬
                        disp_z = disp_z.sort_values(by=sort_col_z, ascending=asc_z)

            # 선택된 행 삭제 기능
            if not disp_z.empty:
                # 선택할 행의 인덱스를 표시
                st.write("행 선택 (삭제할 행 선택):")
                selected_indices = []
                
                # 10개씩 데이터 보기 (페이징)
                items_per_page = 10
                total_pages = max(1, (len(disp_z) + items_per_page - 1) // items_per_page)
                current_page = st.number_input("페이지", min_value=1, max_value=total_pages, value=1) - 1  # 0-인덱싱을 위해 1 빼기
                
                start_idx = current_page * items_per_page
                end_idx = min(start_idx + items_per_page, len(disp_z))
                
                page_indices = list(disp_z.index)[start_idx:end_idx]
                
                # 체크박스로 행 선택 - 접근성 문제 해결
                for idx in page_indices:
                    row = disp_z.loc[idx]
                    col1, col2 = st.columns([1, 9])
                    with col1:
                        # 체크박스에 레이블 추가 (숨김 처리)
                        if st.checkbox(f"항목 {idx}", key=f"select_z_{idx}", label_visibility="collapsed"):
                            selected_indices.append(idx)
                    with col2:
                        # 행 정보 간단히 표시
                        display_info = " | ".join([f"{col}: {row[col]}" for col in disp_z.columns[:3]])
                        st.write(f"ID: {idx}, {display_info}")
                
                # 삭제 버튼
                if selected_indices and st.button("선택한 행 삭제", key="delete_zone"):
                    for idx in selected_indices:
                        # 이력 기록 (전체 행 삭제)
                        for col in df_zone.columns:
                            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            cursor.execute(
                                "INSERT INTO history_zone (row_id,column_name,old_value,new_value,timestamp) VALUES (?,?,?,?,?)",
                                (idx, col, str(disp_z.loc[idx, col]), "[DELETED]", ts)
                            )
                        # DB에서 삭제
                        cursor.execute(f'DELETE FROM delivery_zone_table WHERE id=?', (idx,))
                    
                    conn.commit()
                    st.success(f"{len(selected_indices)}개 행이 삭제되었습니다.")
                    # 데이터 리로드
                    df_zone = pd.read_sql('SELECT * FROM delivery_zone_table', conn, index_col='id')
                    st.experimental_rerun()

            # 데이터 에디터로 편집
            edited_z = st.data_editor(
                disp_z,
                num_rows="dynamic",
                use_container_width=True,
                key='editor_zone',
                height=400
            )
            
            # 변경 사항 저장 버튼
            if st.button("변경사항 저장", key="save_zone_changes"):
                if not edited_z.equals(disp_z):
                    changes_count = 0
                    for idx in edited_z.index:
                        if idx in disp_z.index:  # 인덱스가 있는지 확인
                            old = disp_z.loc[idx]
                            new = edited_z.loc[idx]
                            for col in df_zone.columns:
                                if str(old[col]) != str(new[col]):
                                    changes_count += 1
                                    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                    cursor.execute(
                                        "INSERT INTO history_zone (row_id,column_name,old_value,new_value,timestamp) VALUES (?,?,?,?,?)",
                                        (idx, col, str(old[col]), str(new[col]), ts)
                                    )
                                    
                                    # No 컬럼인 경우 정수형으로 저장
                                    if col == 'No':
                                        try:
                                            value = int(float(new[col]))
                                        except:
                                            value = 0
                                        cursor.execute(f'UPDATE delivery_zone_table SET "{col}"=? WHERE id=?', (value, idx))
                                    else:
                                        cursor.execute(f'UPDATE delivery_zone_table SET "{col}"=? WHERE id=?', (str(new[col]), idx))
                    
                    conn.commit()
                    if changes_count > 0:
                        st.success(f"배송구역 데이터 {changes_count}개 셀 업데이트 완료")
                    else:
                        st.info("변경된 내용이 없습니다.")
                    
                    # 데이터 리로드
                    df_zone = pd.read_sql('SELECT * FROM delivery_zone_table', conn, index_col='id')
                else:
                    st.info("변경된 내용이 없습니다.")
        
        # 탭 2: 신규 추가
        with tab2:
            st.subheader("배송구역 신규 데이터 추가")
            
            # 컬럼 정보 가져오기
            zone_cols = get_table_columns('delivery_zone_table')
            
            # 새 데이터 입력 폼
            new_zone_data = {}
            
            # 컬럼별 입력 필드 표시 (2열로 배치)
            cols = st.columns(2)
            for i, col in enumerate(zone_cols):
                with cols[i % 2]:
                    if col == 'No':
                        # No 컬럼은 기본값으로 최대 값 + 1 제공
                        max_no = 0
                        if 'No' in df_zone.columns and not df_zone.empty:
                            max_no = df_zone['No'].max()
                        new_zone_data[col] = st.number_input(f"{col}", value=max_no+1, step=1, key=f"new_z_{col}")
                    else:
                        new_zone_data[col] = st.text_input(f"{col}", key=f"new_z_{col}")
            
            # 추가 버튼
            if st.button("데이터 추가", key="add_zone"):
                # 필수 입력 체크 (여기서는 No 컬럼만 필수로 가정)
                if 'No' in new_zone_data and new_zone_data['No']:
                    try:
                        # 데이터 추가
                        new_id = add_new_row('delivery_zone_table', new_zone_data)
                        
                        # 이력 기록 (신규 추가)
                        for col, value in new_zone_data.items():
                            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            cursor.execute(
                                "INSERT INTO history_zone (row_id,column_name,old_value,new_value,timestamp) VALUES (?,?,?,?,?)",
                                (new_id, col, "[NEW]", str(value), ts)
                            )
                        
                        conn.commit()
                        st.success(f"새 배송구역 데이터가 추가되었습니다. (ID: {new_id})")
                        
                        # 폼 초기화 및 데이터 리로드
                        df_zone = pd.read_sql('SELECT * FROM delivery_zone_table', conn, index_col='id')
                        st.experimental_rerun()
                    except Exception as e:
                        # 오류 발생 시 롤백 및 에러 메시지 표시
                        conn.rollback()
                        st.error(f"데이터 추가 중 오류가 발생했습니다: {str(e)}")
                else:
                    st.error("No 필드는 필수 입력입니다.")