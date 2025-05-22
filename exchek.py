import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

class ExcelSafetyChecker(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 안전 검사기")
        self.geometry("800x700")
        self.resizable(False, False)

        # 도착용 표준 헤더
        self.standard_headers_arrival = [
            "No", "알림", "발송접수일", "발송접수시간", "구분", "법인", "발송구분", "회수", "미착", "과착", "변상", "운송장번호",
            "고객사주문번호", "인수자타입", "인수자명", "인수완료일시", "접수계정", "발송지", "발송지전화번호", "도착지", "도착지전화번호",
            "보내는분", "보낸분전화번호", "보낸분기타전화번호", "보낸분우편번호", "보낸분주소", "보낸분상세주소", "받는분",
            "도착고객관리", "받는분전화번호", "받는분기타전화번호", "받는분우편번호", "받는분주소", "받는분상세주소", "품목명",
            "배송구분", "포장상태", "수량", "운임구분", "결재구분", "신용", "결재여부", "결재일시", "개별단가", "가로", "세로",
            "높이", "무게", "CBM", "할증운임", "배송운임", "도서운임", "기타운임", "별도운임", "운임합계", "출고비",
            "출고비결재수단", "출고비결재여부", "출고비결재일시", "메모", "노선번호", "발송연선번호", "도착연선번호", "발송지지역",
            "도착지지역", "발송지관할지역", "도착지관할지역", "발송터미널", "도착터미널", "발송터미널하차번호", "도착터미널하차번호"
        ]
        # 발송용 표준 헤더
        self.standard_headers_departure = [
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
        # 발송 파일 필수 필드
        self.DEPARTURE_REQUIRED_FIELDS = ["고객관리번호", "발송고객관리", "거래처 체크"]

        # 초기 상태
        self.current_standard_headers = []
        self.modified_cells = []
        self.modified_dfs = {}
        self.file_type = ''
        self.file_path = ''
        self.wb = None

        # UI 구성
        main = ttk.Frame(self, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # 파일 업로드
        upf = ttk.LabelFrame(main, text="1. 파일 업로드")
        upf.pack(fill=tk.X, pady=5)
        ttk.Button(upf, text="도착 파일 선택", command=self.select_arrival_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(upf, text="발송 파일 선택", command=self.select_departure_file).pack(side=tk.LEFT, padx=5)
        self.file_label = ttk.Label(upf, text="선택된 파일: 없음", foreground="blue")
        self.file_label.pack(side=tk.LEFT, padx=10)

        # 상세 보기
        self.toggle_btn = ttk.Button(main, text="상세 보기", command=self.toggle_details)
        self.toggle_btn.pack(anchor=tk.W, pady=5)
        self.detail_frame = ttk.Frame(main)

        # 표준 헤더 표시
        hdr = ttk.LabelFrame(self.detail_frame, text="2. 표준 헤더")
        hdr.pack(fill=tk.X, pady=5)
        self.std_text = tk.Text(hdr, height=4)
        self.std_text.pack(fill=tk.X, padx=5, pady=2)
        self.std_text.config(state="disabled")

        # 수정된 셀 미리보기
        prv = ttk.LabelFrame(self.detail_frame, text="3. 수정된 셀 미리보기")
        prv.pack(fill=tk.BOTH, expand=True, pady=5)
        cols = ("셀 위치","변경 전","변경 후")
        self.tree = ttk.Treeview(prv, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=200)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(prv, orient=tk.VERTICAL, command=self.tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        # 헤더 비교
        cmpf = ttk.LabelFrame(self.detail_frame, text="4. 표준 헤더 비교")
        cmpf.pack(fill=tk.BOTH, expand=True, pady=5)
        cmp_cols = ("순번","표준","실제","상태")
        self.cmp_tree = ttk.Treeview(cmpf, columns=cmp_cols, show="headings", height=6)
        for c in cmp_cols:
            self.cmp_tree.heading(c, text=c)
            self.cmp_tree.column(c, width=150)
        self.cmp_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(cmpf, orient=tk.VERTICAL, command=self.cmp_tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        # 결과 및 저장
        res = ttk.Frame(main)
        res.pack(fill=tk.X, pady=5)
        self.result_label = ttk.Label(res, text="처리 결과: -", foreground="green", font=(None,12,'bold'))
        self.result_label.pack(side=tk.LEFT)
        self.save_btn = ttk.Button(res, text="파일 저장", command=self.save_file, state=tk.DISABLED)
        self.save_btn.pack(side=tk.RIGHT)

    def select_arrival_file(self):
        self.current_standard_headers = self.standard_headers_arrival
        self.refresh_std_text()
        path = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")], title="도착 파일 선택")
        if path:
            self.process_file(path, mode='arrival')

    def select_departure_file(self):
        self.current_standard_headers = self.standard_headers_departure
        self.refresh_std_text()
        path = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")], title="발송 파일 선택")
        if path:
            self.process_file(path, mode='departure')

    def refresh_std_text(self):
        self.std_text.config(state='normal')
        self.std_text.delete('1.0', tk.END)
        self.std_text.insert(tk.END, '\n'.join(self.current_standard_headers))
        self.std_text.config(state='disabled')

    def toggle_details(self):
        if self.detail_frame.winfo_ismapped():
            self.detail_frame.pack_forget()
            self.toggle_btn.config(text='상세 보기')
        else:
            self.detail_frame.pack(fill=tk.BOTH, expand=True)
            self.toggle_btn.config(text='상세 숨기기')

    def process_file(self, file_path, mode):
        # 초기화
        self.file_path = file_path
        self.file_label.config(text=f"선택된 파일: {os.path.basename(file_path)}")
        self.modified_cells.clear()
        self.cmp_tree.delete(*self.cmp_tree.get_children())
        total = modified = 0

        ext = os.path.splitext(file_path)[1].lower()
        if ext not in ['.xlsx',' .xls']:
            messagebox.showwarning('경고','.xlsx/.xls만 지원')
            return

        # .xlsx: COM으로 직접 처리
        if ext == '.xlsx':
            import win32com.client as win32
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = False
            wb_com = excel.Workbooks.Open(os.path.abspath(file_path))
            for sheet in wb_com.Sheets:
                used = sheet.UsedRange
                rows = used.Rows.Count
                cols = used.Columns.Count
                for r in range(2, rows+1):
                    for c in range(1, cols+1):
                        total += 1
                        cell = sheet.Cells(r,c)
                        val = cell.Formula
                        if isinstance(val,str) and (val.startswith(('=','+')) or val.endswith('\\')):
                            new = val.lstrip('=+').rstrip('\\')
                            if new != val:
                                cell.NumberFormat = '@'
                                cell.Value = new
                                coord = f"{get_column_letter(c)}{r}"
                                self.modified_cells.append((coord,val,new))
                                modified +=1
            # 저장
            save_path = filedialog.asksaveasfilename(initialfile=os.path.basename(file_path), defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')])
            if save_path:
                wb_com.SaveAs(os.path.abspath(save_path))
            wb_com.Close(False)
            excel.Quit()
            headers = []
        else:
            # .xls 처리 (pandas)
            sheets = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str, engine='xlrd')
            for name,df in sheets.items():
                df2 = df.copy()
                for r in range(1, df2.shape[0]):
                    for c in range(df2.shape[1]):
                        total += 1
                        val = df2.iat[r,c] or ''
                        if isinstance(val,str) and (val.startswith(('=','+')) or val.endswith('\\')):
                            new = val.lstrip('=+').rstrip('\\')
                            df2.iat[r,c] = new
                            coord = f"{get_column_letter(c+1)}{r+1}"
                            self.modified_cells.append((coord,val,new))
                            modified +=1
                self.modified_dfs[name] = df2
            headers = [str(x) if pd.notna(x) else '' for x in sheets[next(iter(sheets))].iloc[0]]
            self.save_btn.config(state=tk.NORMAL)

        # 모드별 헤더 검증 및 비교
        for idx,exp in enumerate(self.current_standard_headers,1):
            act = headers[idx-1] if idx-1<len(headers) else ''
            status = '일치' if exp==act else '불일치'
            self.cmp_tree.insert('',tk.END,values=(idx,exp,act,status))

        # 결과 표시
        self.result_label.config(text=f"총 셀: {total}, 수정: {modified}")
        self.update_treeview()

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for coord,old,new in self.modified_cells:
            self.tree.insert('',tk.END,values=(coord,old,new))

    def save_file(self):
        if not self.modified_cells:
            return
        save_path = filedialog.asksaveasfilename(initialfile=os.path.basename(self.file_path), defaultextension=self.file_type, filetypes=[('Excel',f'*{self.file_type}')])
        if save_path and self.file_type=='.xls':
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                for name,df in self.modified_dfs.items():
                    df.to_excel(writer, sheet_name=name, header=False, index=False)
            messagebox.showinfo('완료',f"저장됨: {save_path}")

if __name__=='__main__':
    app=ExcelSafetyChecker()
    app.mainloop()
