# Author : 문은영
# Revision : 2025-03-31
# Modified : 2025-06-26 (리팩터링: 파일 단일화, 속도 개선, 폰트 확대, 실시간 파일 감시, 전화번호 자동정리, ESC 닫기)

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkFont
import threading
import time
import os
import re

# ============================================================
# 설정
# ============================================================
file_path = '북돋움관리대장.xlsx'
sheet_name = 'Sheet1'

df = None
last_mtime = None


# ============================================================
# 전화번호 자동 정리 함수
# ============================================================
def format_phone(phone):
    """01012345678 → 010-1234-5678 자동 변환"""
    if not isinstance(phone, str):
        return phone

    digits = re.sub(r'[^0-9]', "", phone)

    if len(digits) == 11:  # 01012345678
        return f"{digits[0:3]}-{digits[3:7]}-{digits[7:11]}"
    elif len(digits) == 10:  # 02 등 지역번호 포함한 형태
        return f"{digits[0:3]}-{digits[3:6]}-{digits[6:10]}"
    else:
        return phone


# ============================================================
# GUI 생성
# ============================================================
root = tk.Tk()
root.title("북돋움 관리 검색 프로그램")
root.geometry("600x380")

# ============================================================
# 폰트 확대
# ============================================================
default_font = tkFont.nametofont("TkDefaultFont")
default_font.configure(size=12)
tkFont.nametofont("TkTextFont").configure(size=12)
tkFont.nametofont("TkFixedFont").configure(size=12)


# ============================================================
# 엑셀 로딩
# ============================================================
def load_excel():
    """엑셀 파일 불러오기"""
    global df, last_mtime

    try:
        df_local = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        df_local.columns = df_local.columns.str.strip()

        # 생년월일 자리수 정리
        if '생년월일' in df_local.columns:
            df_local['생년월일'] = df_local['생년월일'].astype(str).str.zfill(6)

        # 전화번호 자동 변환
        if "연락처" in df_local.columns:
            df_local["연락처"] = df_local["연락처"].apply(format_phone)

        df = df_local
        last_mtime = os.path.getmtime(file_path)

    except FileNotFoundError:
        messagebox.showerror("파일 오류", f"엑셀 파일을 찾을 수 없습니다:\n{file_path}")
        root.destroy()
        exit()


# ============================================================
# 파일 변경 실시간 감시
# ============================================================
def watch_file():
    global df, last_mtime

    while True:
        time.sleep(1)
        try:
            current_mtime = os.path.getmtime(file_path)
            if last_mtime is None or current_mtime != last_mtime:

                df_local = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
                df_local.columns = df_local.columns.str.strip()

                if '생년월일' in df_local.columns:
                    df_local['생년월일'] = df_local['생년월일'].astype(str).str.zfill(6)

                # 전화번호 자동 변환
                if "연락처" in df_local.columns:
                    df_local["연락처"] = df_local["연락처"].apply(format_phone)

                df = df_local
                last_mtime = current_mtime

        except Exception:
            pass


# ============================================================
# 검색 기능
# ============================================================
result_window = None
tree = None

def search_data(event=None):
    global result_window, tree, df

    if df is None:
        messagebox.showwarning("로딩 중", "엑셀 파일 로딩 중입니다. 잠시 후 다시 검색하세요.")
        return

    # 기존 결과창 제거
    if result_window is not None:
        result_window.destroy()
        result_window = None

    category = category_var.get()
    search_term = search_entry.get().strip()

    if not search_term:
        messagebox.showwarning("입력 오류", "검색어를 입력하세요!")
        return

    if category == "생년월일":
        search_term = search_term.zfill(6)

    if category not in df.columns:
        messagebox.showerror("오류", f"'{category}' 항목이 엑셀 파일에 없습니다.")
        return

    filtered = df[df[category].str.contains(search_term, case=False, na=False, regex=False)]

    if filtered.empty:
        messagebox.showinfo("검색 결과", f"'{search_term}'에 대한 검색 결과가 없습니다.")
        return

    # 결과창 생성
    result_window = tk.Toplevel(root)
    result_window.title("검색 결과")

    # ✔ 요청: 창 가로 크기 확대
    result_window.geometry("1300x300")

    tk.Label(result_window, text=f"{len(filtered)}건 검색됨", font=("Arial", 12)).pack(pady=5)

    # ESC로 결과창 닫기
    result_window.bind("<Escape>", lambda e: result_window.destroy())

    cols = [
        '부모이름', '생년월일', '임신부/부모', '임신확인일/출산예정일',
        '영아이름', '영아생년월일', '주소', '연락처'
    ]
    existing_cols = [col for col in cols if col in filtered.columns]

    style = ttk.Style()
    style.configure("Treeview", font=("Arial", 13), rowheight=32)
    style.configure("Treeview.Heading", font=("Arial", 13, "bold"))

    tree = ttk.Treeview(result_window, columns=existing_cols, show="headings")
    tree.pack(expand=True, fill="both")

    for col in existing_cols:
        if col == "주소":
            tree.column(col, width=450)
        elif col == "연락처":
            tree.column(col, width=130)
        else:
            tree.column(col, width=80)
        tree.heading(col, text=col)

    # 데이터 삽입
    for _, row in filtered.iterrows():
        tree.insert("", "end", values=[row[col] for col in existing_cols])


# ============================================================
# GUI 구성
# ============================================================
tk.Label(root, text="검색할 분류", font=("Arial", 12, "bold")).pack(pady=5)

categories = ['부모이름', '생년월일', '임신부/부모', '영아이름', '영아생년월일', '주소', '연락처']
category_var = tk.StringVar()
dropdown = ttk.Combobox(root, textvariable=category_var, values=categories, state="readonly", width=20)
dropdown.pack(pady=5)
dropdown.current(0)

tk.Label(root, text="검색어 입력", font=("Arial", 12, "bold")).pack(pady=5)
search_entry = tk.Entry(root, font=("Arial", 12), width=30)
search_entry.pack(pady=5)
search_entry.focus_set()

search_button = tk.Button(root, text="검색", command=search_data,
                          bg="orange", font=("Arial", 14, "bold"), width=18, height=2)
search_button.pack(pady=12)

root.bind("<Return>", search_data)

# ============================================================
# 엑셀 로딩 + 감시 쓰레드 실행
# ============================================================
threading.Thread(target=load_excel, daemon=True).start()
threading.Thread(target=watch_file, daemon=True).start()

root.mainloop()
