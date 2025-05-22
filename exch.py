import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import sys

# -----------------------------------------
# 구도착자료(최종) 양식
# -----------------------------------------
old_arrival_cols = [
    "No", "도착구분", "미착구분", "접수일", "운송장번호", "인수자타입", "인수완료시간", "접수시간", "입력ID",
    "발송지", "발송지 전화번호", "도착지", "도착지 전화번호", "보내는분", "보낸분 전화번호", "보낸분 기타전화번호",
    "보낸분 우편번호", "보낸분 주소", "보낸분 상세주소", "받는분", "고객입력", "받는분 전화번호", "받는분 기타전화번호",
    "받는분 우편번호", "받는분 주소", "받는분 상세주소", "품목명", "귀중품", "포장상태", "배송구분", "신용", "수량",
    "개별단가", "배송예정일", "배송예정시간", "할증운임", "배송운임", "도서운임", "기타운임", "별도운임",
    "하역비", "메모", "배송기사", "배송기사 전화", "도선구분", "도서택배요율", "도서화물요율", "노선번호",
    "발송연선번호", "도착연선번호", "발송지 지역", "도착지 지역", "발송지 관할지점", "도착지 관할지점",
    "발송 터미널", "도착 터미널", "발송 터미널 하차번호", "도착 터미널 하차번호", "수정시간", "수정내역",
    "담당기사", "인수자명", "법인 구분", "예약번호", "도착지 수수료", "도착지 운임삭제"
]

# -----------------------------------------------------
# 기본 매핑 (일부 예시)
# 여기서 "인수완료시간" -> "인수완료일시" 로 매핑해두면,
# 해당 항목은 날짜만 남기도록 처리됨.
# -----------------------------------------------------
default_mapping = {
    "No": "No",
    "접수일": "발송접수일",
    "운송장번호": "운송장번호",
    "인수자타입": "인수자타입",
    "인수완료시간": "인수완료일시",
    "접수시간": "발송접수시간",
    "입력ID": "접수계정",
    "발송지": "발송지",
    "발송지 전화번호": "발송지전화번호",
    "도착지": "도착지",
    "도착지 전화번호": "도착지전화번호",
    "보내는분": "보내는분",
    "보낸분 전화번호": "보낸분전화번호",
    "보낸분 기타전화번호": "보낸분기타전화번호",
    "보낸분 우편번호": "보낸분우편번호",
    "보낸분 주소": "보낸분주소",
    "보낸분 상세주소": "보낸분상세주소",
    "받는분": "받는분",
    "고객입력": "도착고객관리",
    "받는분 전화번호": "받는분전화번호",
    "받는분 기타전화번호": "받는분기타전화번호",
    "받는분 우편번호": "받는분우편번호",
    "받는분 주소": "받는분주소",
    "받는분 상세주소": "받는분상세주소",
    "품목명": "품목명",
    "포장상태": "포장상태",
    "배송구분": "운임구분",
    "신용": "신용",
    "수량": "수량",
    "개별단가": "개별단가",
    "할증운임": "할증운임",
    "배송운임": "배송운임",
    "도서운임": "도서운임",
    "기타운임": "기타운임",
    "별도운임": "별도운임",
    "하역비": "출고비",
    "메모": "메모",
    "노선번호": "노선번호",
    "발송연선번호": "발송연선번호",
    "도착연선번호": "도착연선번호",
    "발송지 지역": "발송지지역",
    "도착지 지역": "도착지지역",
    "발송지 관할지점": "발송지관할지역",
    "도착지 관할지점": "도착지관할지역",
    "발송 터미널": "발송터미널",
    "도착 터미널": "도착터미널",
    "발송 터미널 하차번호": "발송터미널하차번호",
    "도착 터미널 하차번호": "도착터미널하차번호",
    "담당기사": "배달기사",
    "인수자명": "인수자명",
    "법인 구분": "법인",
    "도착지 수수료": "도착지수수료"
}

class OldArrivalConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("구도착자료 변환")

        self.df = None
        self.comboboxes = {}

        self.create_main_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_main_ui(self):
        """메인 화면: 파일 선택 버튼"""
        label_title = tk.Label(self.root, text="구도착자료 변환 프로그램", font=("Arial", 12, "bold"))
        label_title.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        btn_load = tk.Button(self.root, text="파일 선택", command=self.load_excel_file)
        btn_load.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

    def load_excel_file(self):
        """신도착자료 엑셀 파일 선택 및 로드"""
        file_path = filedialog.askopenfilename(
            title="신도착자료 엑셀 파일 선택",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if not file_path:
            messagebox.showwarning("경고", "파일이 선택되지 않았습니다.")
            return

        try:
            self.df = pd.read_excel(file_path)
            messagebox.showinfo("알림", "엑셀 파일을 성공적으로 불러왔습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일을 읽는 도중 오류 발생: {e}")
            return

        # 매핑 변경 여부
        change_mapping = messagebox.askyesno("매핑 설정", "매핑을 변경하시겠습니까?")
        if change_mapping:
            self.show_mapping_ui()
        else:
            # 기본 매핑 사용
            user_mapping = {}
            for col_name in old_arrival_cols:
                user_mapping[col_name] = default_mapping.get(col_name, "")
            self.convert_and_save(user_mapping)

    def show_mapping_ui(self):
        """구도착자료 항목 ↔ 신도착자료 컬럼 매핑 설정 UI"""
        mapping_window = tk.Toplevel(self.root)
        mapping_window.title("매핑 설정")
        mapping_window.protocol("WM_DELETE_WINDOW", self.on_close)

        pairs_per_row = 2
        for i, col_name in enumerate(old_arrival_cols):
            row_idx = i // pairs_per_row
            col_idx = (i % pairs_per_row) * 2

            lbl = tk.Label(mapping_window, text=col_name)
            lbl.grid(row=row_idx, column=col_idx, padx=5, pady=3, sticky="w")

            cb = ttk.Combobox(mapping_window, values=list(self.df.columns), width=15)
            cb.grid(row=row_idx, column=col_idx+1, padx=5, pady=3)

            # 기본 매핑
            if col_name in default_mapping:
                default_val = default_mapping[col_name]
                if default_val in self.df.columns:
                    cb.set(default_val)

            self.comboboxes[col_name] = cb

        row_end = (len(old_arrival_cols) - 1) // pairs_per_row + 1
        btn_confirm = tk.Button(mapping_window, text="확인", command=lambda: self.on_confirm(mapping_window))
        btn_confirm.grid(row=row_end, column=0, columnspan=2, pady=10)

    def on_confirm(self, mapping_window):
        """매핑 창에서 확인 버튼 -> 매핑 정보 추출 -> 변환 & 저장"""
        user_mapping = {}
        for col_name, cb in self.comboboxes.items():
            selected = cb.get().strip()
            user_mapping[col_name] = selected

        mapping_window.destroy()
        self.convert_and_save(user_mapping)

    def convert_and_save(self, user_mapping):
        """구도착자료 형태로 변환 후 엑셀로 저장"""
        try:
            new_df = pd.DataFrame()
            for col_name in old_arrival_cols:
                mapped_col = user_mapping.get(col_name, "")
                if mapped_col and (mapped_col in self.df.columns):
                    # [핵심 로직] "인수완료시간" 항목인 경우 => 날짜만 추출
                    if col_name == "인수완료시간":
                        # 엑셀에서 읽은 열을 datetime으로 파싱 후, 날짜 부분만
                        new_df[col_name] = (
                            pd.to_datetime(self.df[mapped_col], errors="coerce")
                              .dt.strftime("%Y-%m-%d")  # 'YYYY-MM-DD' 형태
                              .fillna("")
                        )
                    else:
                        # 일반 매핑
                        new_df[col_name] = self.df[mapped_col]
                else:
                    # 매핑이 없거나, 신도착자료에 해당 컬럼이 없으면 빈 칸
                    new_df[col_name] = ""
        except Exception as e:
            messagebox.showerror("오류", f"데이터 변환 중 오류 발생: {e}")
            return

        # 저장 파일 경로
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            title="구도착자료 엑셀 파일 저장",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if not save_path:
            messagebox.showwarning("경고", "저장할 파일이 선택되지 않았습니다.")
            return

        # 엑셀 저장
        try:
            new_df.to_excel(save_path, index=False)
            messagebox.showinfo("완료", "구도착자료 양식으로 파일이 성공적으로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 중 오류 발생: {e}")
            return

        self.on_close()

    def on_close(self):
        """창을 닫으면 프로그램 완전 종료"""
        self.root.destroy()
        sys.exit(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = OldArrivalConverterApp(root)
    root.mainloop()
