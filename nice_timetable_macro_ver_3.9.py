import os
import subprocess
import time
import json
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

import openpyxl


# ============================================================
# [설정]
# ============================================================

CONFIG_FILE = "mapping_config.json"

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

DEBUG_PORT = 9224

DEBUG_USER_DATA_DIR = r"C:\sel_chrome"

EDUPTL_URL = "https://ice.eduptl.kr/"

NEIS_URL = "https://ice.neis.go.kr/jsp/main.jsp"


# ============================================================
# [기본 과목 매핑]
# ============================================================

DEFAULT_MAPPING = {
    "국어": "국어(",
    "사회": "사회(",
    "수학": "수학(",
    "실과": "실과(",
    "체육": "체육(",
    "음악": "음악(",
    "진로": "진로활동(",
    "자율": "자율·자치활동(",

    "과학(성)": "과학(성",
    "과학": "과학(성",

    "도덕(김)": "도덕(김승",
    "도덕": "도덕(김승",

    "영어(이)": "영어(이",
    "영어": "영어(이",

    "동아리": "동아리활동(도",
    "미술": "미술(",

    "영(이)": "영어(이",
    "과(성)": "과학(성",
    "과(담)": "과학(담",
    "체(해)": "체육(해",
    "창체": "창체",
}


# ============================================================
# [소개]
# ============================================================

def print_intro():

    intro_text = """
============================================================
[시간표 입력 매크로 v3.9]
============================================================
- 9224 디버깅 Chrome 자동 실행
- 업무포털 자동 접속
- 로그인 버튼 자동 클릭
- 로그인 및 NEIS 이동은 사용자가 직접 진행
- Enter 입력 후 시간표 자동 입력
- Excel / 한글(HWP) 시간표 붙여넣기 지원
- 한글 표의 열 우선 복사 문제 자동 보정
- 빈 셀 보존
- EVPN 독립형
- 제작: 인천용현남초등학교 왘도
============================================================
"""

    print(intro_text)


# ============================================================
# [매핑 설정 불러오기]
# ============================================================

def load_mapping():

    if os.path.exists(CONFIG_FILE):

        try:

            with open(
                CONFIG_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                return json.load(f)

        except Exception:

            return DEFAULT_MAPPING.copy()

    return DEFAULT_MAPPING.copy()


# ============================================================
# [매핑 설정 저장]
# ============================================================

def save_mapping(mapping_data):

    try:

        with open(
            CONFIG_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                mapping_data,
                f,
                ensure_ascii=False,
                indent=4
            )

        return True

    except Exception as e:

        print(f"❌ 설정 저장 실패: {e}")

        return False


# ============================================================
# [시간표 파싱]
#
# Excel:
#   행 우선
#
# HWP:
#   열 우선으로 들어오는 경우
#   30개의 셀을 6개씩 잘라서
#   다시 6행 × 5열로 변환
# ============================================================

def parse_timetable(raw_text):

    # ========================================================
    # 줄바꿈 통일
    # ========================================================

    raw_text = raw_text.replace("\r\n", "\n")
    raw_text = raw_text.replace("\r", "\n")


    # ========================================================
    # ① Excel에서 복사한 경우
    #
    # Excel:
    # 미술\t국어\t자율\t국어\t국어
    # 미술\t실과\t국어\t실과\t실과
    # ...
    # ========================================================

    if "\t" in raw_text:

        lines = raw_text.split("\n")

        grid = []

        for line in lines:

            # 완전히 빈 줄은 Excel에서 생기는
            # 불필요한 마지막 줄일 가능성이 높음
            if not line.strip():
                continue

            row = line.split("\t")

            row = [
                cell.strip()
                for cell in row
            ]

            # 마지막 빈 셀 제거
            while row and row[-1] == "":
                row.pop()

            grid.append(row)


        # 최소 5행
        if len(grid) < 5:

            return None, (
                "Excel 시간표를 인식했지만 "
                "행의 개수가 부족합니다."
            )


        # 각 행 5열 확인
        for row_index, row in enumerate(
            grid[:6],
            start=1
        ):

            if len(row) < 5:

                return None, (
                    f"{row_index}교시 행의 "
                    f"과목 수가 {len(row)}개입니다."
                )


        return grid[:6], "Excel 형식"


    # ========================================================
    # ② 한글(HWP)에서 복사한 경우
    #
    # 중요!
    #
    # 현재 사용자의 한글 표는 이미
    # '행 우선'으로 들어온다.
    #
    # 따라서 절대로 전치하지 않는다.
    #
    # 예:
    #
    # 미술
    # 국어
    # 자율
    # 국어
    # 국어
    # 미술
    # 실과
    # ...
    #
    # 그대로
    #
    # [0]~[4]   = 1교시
    # [5]~[9]   = 2교시
    # [10]~[14] = 3교시
    # ...
    # ========================================================

    lines = raw_text.split("\n")


    # --------------------------------------------------------
    # 마지막 자동 개행만 제거
    #
    # 중요:
    # 중간의 빈 줄은 절대 제거하지 않는다.
    #
    # 사용자의 시간표에서
    #
    # 과(담)
    # 체육
    #
    # [빈칸]
    #
    # 사회
    # 체(해)
    #
    # 여기의 빈 줄이 수요일 6교시이다.
    # --------------------------------------------------------

    while lines and lines[-1] == "":
        lines.pop()


    # 앞쪽의 불필요한 빈 줄만 제거
    while lines and lines[0] == "":
        lines.pop(0)


    # --------------------------------------------------------
    # 각 셀의 앞뒤 공백만 제거
    #
    # 빈 문자열은 그대로 유지
    # --------------------------------------------------------

    cells = [
        line.strip()
        for line in lines
    ]


    # ========================================================
    # HWP 정상 형태
    #
    # 6교시 × 5요일 = 30셀
    #
    # 사용자의 경우:
    #
    # 30개 셀 중
    # 수요일 6교시가 빈칸
    #
    # 따라서 빈 문자열도 반드시 포함되어야 한다.
    # ========================================================

    if len(cells) == 30:

        grid = []


        # ----------------------------------------------------
        # 5개씩 잘라서 6교시 구성
        # ----------------------------------------------------

        for row_index in range(6):

            start = row_index * 5
            end = start + 5

            row = cells[start:end]

            grid.append(row)


        return grid, "한글(HWP) 행 우선 형식"


    # ========================================================
    # ③ 혹시 30셀이 아닌 경우
    # ========================================================

    return None, (
        f"시간표 셀 개수를 정확하게 인식하지 못했습니다.\n\n"
        f"현재 인식된 셀 개수: {len(cells)}개\n"
        f"정상적인 시간표는 30개(6교시 × 5요일)여야 합니다.\n\n"
        f"※ 빈칸도 하나의 셀로 계산됩니다."
    )


# ============================================================
# [시간표 GUI]
# ============================================================

def show_main_gui():

    current_mapping = load_mapping()

    result = {
        "data": None,
        "timetable": None,
        "status": "cancel"
    }


    root = tk.Tk()

    root.title(
        "시간표 입력 매크로 v3.9"
    )

    root.geometry(
        "720x760"
    )

    root.resizable(
        False,
        False
    )


    # ========================================================
    # Notebook
    # ========================================================

    notebook = ttk.Notebook(root)

    notebook.pack(
        fill="both",
        expand=True,
        padx=10,
        pady=10
    )


    # ========================================================
    # TAB 1
    # ========================================================

    tab1 = tk.Frame(notebook)

    notebook.add(
        tab1,
        text=" 📅 시간표 붙여넣기 "
    )


    info_lbl1 = tk.Label(
        tab1,
        text=(
            "Excel 또는 한글(HWP)의 시간표 표를 그대로 복사한 후\n"
            "아래 입력창에 Ctrl+V로 붙여넣으세요.\n\n"
            "※ 한글 표는 열 단위로 복사되는 문제를 자동으로 보정합니다.\n"
            "※ 빈칸도 그대로 유지합니다."
        ),
        font=(
            "Malgun Gothic",
            10,
            "bold"
        ),
        justify="left",
        fg="#1f4e78"
    )

    info_lbl1.pack(
        pady=10,
        anchor="w",
        padx=10
    )


    # ========================================================
    # 시간표 입력 Text
    # ========================================================

    txt_timetable = tk.Text(
        tab1,
        width=80,
        height=18,
        font=(
            "Malgun Gothic",
            10
        )
    )

    txt_timetable.pack(
        padx=10,
        pady=5
    )


    # ========================================================
    # 안내 문구
    # ========================================================

    placeholder = (
        "여기에 Excel 또는 한글 시간표를\n"
        "Ctrl+V로 붙여넣으세요."
    )

    txt_timetable.insert(
        "1.0",
        placeholder
    )


    def clear_placeholder(event):

        current = txt_timetable.get(
            "1.0",
            tk.END
        )

        if placeholder in current:

            txt_timetable.delete(
                "1.0",
                tk.END
            )


    txt_timetable.bind(
        "<FocusIn>",
        clear_placeholder
    )


    # ========================================================
    # 인식 결과 표시
    # ========================================================

    preview_frame = tk.LabelFrame(
        tab1,
        text=" 시간표 인식 결과 "
    )

    preview_frame.pack(
        fill="both",
        padx=10,
        pady=8
    )


    preview_text = tk.Text(
        preview_frame,
        width=80,
        height=8,
        font=(
            "Consolas",
            10
        ),
        state="disabled"
    )

    preview_text.pack(
        padx=5,
        pady=5
    )


    # ========================================================
    # 시간표 미리보기
    # ========================================================

    def preview_timetable():

        raw_text = txt_timetable.get(
            "1.0",
            tk.END
        ).strip("\n")


        if not raw_text.strip():

            messagebox.showwarning(
                "시간표",
                "시간표를 붙여넣어 주세요."
            )

            return


        grid, mode = parse_timetable(
            raw_text
        )


        preview_text.config(
            state="normal"
        )

        preview_text.delete(
            "1.0",
            tk.END
        )


        if grid is None:

            preview_text.insert(
                "1.0",
                "❌ 인식 실패\n\n"
                + mode
            )

            preview_text.config(
                state="disabled"
            )

            return


        days = [
            "월",
            "화",
            "수",
            "목",
            "금"
        ]


        preview_text.insert(
            tk.END,
            f"✅ {mode}으로 인식했습니다.\n\n"
        )


        preview_text.insert(
            tk.END,
            "      "
        )


        for day in days:

            preview_text.insert(
                tk.END,
                f"{day:^10}"
            )


        preview_text.insert(
            tk.END,
            "\n"
        )


        preview_text.insert(
            tk.END,
            "-" * 58
        )

        preview_text.insert(
            tk.END,
            "\n"
        )


        for row_index, row in enumerate(
            grid[:6],
            start=1
        ):

            preview_text.insert(
                tk.END,
                f"{row_index}교시 "
            )


            for col_index in range(5):

                if col_index < len(row):

                    value = row[col_index]

                else:

                    value = ""


                if value == "":

                    value = "(빈칸)"


                preview_text.insert(
                    tk.END,
                    f"{value:^10}"
                )


            preview_text.insert(
                tk.END,
                "\n"
            )


        preview_text.config(
            state="disabled"
        )


    btn_preview = tk.Button(
        tab1,
        text="🔍 시간표 인식 확인",
        command=preview_timetable,
        bg="#e3f2fd",
        font=(
            "Malgun Gothic",
            10,
            "bold"
        )
    )

    btn_preview.pack(
        fill="x",
        padx=10,
        pady=5
    )


    # ========================================================
    # TAB 2 : 과목 매핑
    # ========================================================

    tab2 = tk.Frame(notebook)

    notebook.add(
        tab2,
        text=" ⚙️ 과목/교사 매핑 설정 "
    )


    info_lbl2 = tk.Label(
        tab2,
        text=(
            "엑셀 또는 한글 시간표의 과목명과\n"
            "NEIS 시스템 메뉴의 실제 과목명을 연결합니다."
        ),
        font=(
            "Malgun Gothic",
            9
        ),
        justify="left",
        fg="#2c3e50"
    )

    info_lbl2.pack(
        pady=5,
        anchor="w",
        padx=10
    )


    # ========================================================
    # Treeview
    # ========================================================

    tree_frame = tk.Frame(tab2)

    tree_frame.pack(
        fill="both",
        expand=True,
        padx=10,
        pady=5
    )


    tree = ttk.Treeview(
        tree_frame,
        columns=(
            "excel",
            "system"
        ),
        show="headings",
        height=15
    )


    tree.heading(
        "excel",
        text="시간표 과목명"
    )

    tree.heading(
        "system",
        text="NEIS 매칭 문자"
    )


    tree.column(
        "excel",
        width=180,
        anchor="center"
    )

    tree.column(
        "system",
        width=300,
        anchor="w"
    )


    tree.pack(
        side="left",
        fill="both",
        expand=True
    )


    scrollbar = ttk.Scrollbar(
        tree_frame,
        orient="vertical",
        command=tree.yview
    )


    tree.configure(
        yscrollcommand=scrollbar.set
    )


    scrollbar.pack(
        side="right",
        fill="y"
    )


    def update_tree():

        tree.delete(
            *tree.get_children()
        )

        for key, value in current_mapping.items():

            tree.insert(
                "",
                "end",
                values=(
                    key,
                    value
                )
            )


    update_tree()


    # ========================================================
    # Excel 매핑 업로드
    # ========================================================

    def upload_mapping_excel():

        file_path = filedialog.askopenfilename(
            title="매핑 설정 엑셀 파일 선택",
            filetypes=[
                (
                    "Excel files",
                    "*.xlsx"
                )
            ]
        )


        if not file_path:
            return


        try:

            wb = openpyxl.load_workbook(
                file_path
            )

            sheet = wb.active

            new_count = 0


            for row in sheet.iter_rows(
                values_only=True
            ):

                if (
                    len(row) >= 2
                    and row[0]
                    and row[1]
                ):

                    key = str(
                        row[0]
                    ).strip()

                    value = str(
                        row[1]
                    ).strip()


                    current_mapping[key] = value

                    new_count += 1


            update_tree()


            messagebox.showinfo(
                "성공",
                (
                    f"{new_count}개의 "
                    "매핑 설정을 불러왔습니다."
                )
            )


        except Exception as e:

            messagebox.showerror(
                "오류",
                f"엑셀 파일 불러오기 실패:\n{e}"
            )


    btn_upload = tk.Button(
        tab2,
        text="📁 매핑 설정 엑셀 파일 업로드",
        command=upload_mapping_excel,
        bg="#fff9c4",
        font=(
            "Malgun Gothic",
            9,
            "bold"
        )
    )


    btn_upload.pack(
        fill="x",
        padx=10,
        pady=5
    )


    # ========================================================
    # 수동 매핑 입력
    # ========================================================

    input_frame = tk.LabelFrame(
        tab2,
        text=" 수동 항목 수정 "
    )


    input_frame.pack(
        fill="x",
        padx=10,
        pady=5
    )


    tk.Label(
        input_frame,
        text="시간표:"
    ).grid(
        row=0,
        column=0,
        padx=5,
        pady=5
    )


    entry_excel = tk.Entry(
        input_frame,
        width=12
    )


    entry_excel.grid(
        row=0,
        column=1,
        padx=5,
        pady=5
    )


    tk.Label(
        input_frame,
        text="➡️ NEIS:"
    ).grid(
        row=0,
        column=2,
        padx=5,
        pady=5
    )


    entry_system = tk.Entry(
        input_frame,
        width=20
    )


    entry_system.grid(
        row=0,
        column=3,
        padx=5,
        pady=5
    )


    def add_or_update():

        e_val = entry_excel.get().strip()

        s_val = entry_system.get().strip()


        if not e_val or not s_val:

            return


        current_mapping[e_val] = s_val

        update_tree()


        entry_excel.delete(
            0,
            tk.END
        )

        entry_system.delete(
            0,
            tk.END
        )


    def delete_item():

        selected = tree.selection()


        if not selected:
            return


        e_val = tree.item(
            selected[0]
        )["values"][0]


        if e_val in current_mapping:

            del current_mapping[e_val]


        update_tree()


    tk.Button(
        input_frame,
        text="추가/수정",
        command=add_or_update,
        bg="#e1f5fe"
    ).grid(
        row=0,
        column=4,
        padx=5,
        pady=5
    )


    tk.Button(
        input_frame,
        text="삭제",
        command=delete_item,
        bg="#ffebee",
        fg="red"
    ).grid(
        row=0,
        column=5,
        padx=5,
        pady=5
    )


    # ========================================================
    # 실행
    # ========================================================

    def on_run():

        raw_text = txt_timetable.get(
            "1.0",
            tk.END
        )


        if (
            not raw_text.strip()
            or placeholder in raw_text
        ):

            messagebox.showwarning(
                "시간표",
                "시간표를 붙여넣어 주세요."
            )

            notebook.select(tab1)

            return


        grid, mode = parse_timetable(
            raw_text
        )


        if grid is None:

            messagebox.showerror(
                "시간표 인식 실패",
                mode
            )

            notebook.select(tab1)

            return


        # ----------------------------------------------------
        # 매핑 저장
        # ----------------------------------------------------

        save_mapping(
            current_mapping
        )


        # ----------------------------------------------------
        # 결과 저장
        # ----------------------------------------------------

        result["data"] = current_mapping

        result["timetable"] = grid

        result["status"] = "run"


        root.destroy()


    btn_run = tk.Button(
        root,
        text="🚀 시간표 파싱 및 매크로 시작",
        font=(
            "Malgun Gothic",
            12,
            "bold"
        ),
        bg="#c8e6c9",
        fg="black",
        height=2,
        command=on_run
    )


    btn_run.pack(
        fill="x",
        padx=15,
        pady=10
    )


    root.mainloop()


    return result


# ============================================================
# [NEIS 시간표 자동 입력]
# ============================================================

def run_timetable_macro(
    driver,
    timetable_grid,
    name_mapping
):

    wait = WebDriverWait(
        driver,
        0.5
    )


    days = [
        "월",
        "화",
        "수",
        "목",
        "금"
    ]


    print()
    print("=" * 70)
    print("🚀 [4단계] 시간표 자동 입력 시작")
    print("=" * 70)


    # ========================================================
    # 1교시 ~ 6교시
    # ========================================================

    for row_idx in range(1, 7):

        # ----------------------------------------------------
        # 월 ~ 금
        # ----------------------------------------------------

        for col_idx in range(1, 6):

            day_name = days[
                col_idx - 1
            ]


            # ------------------------------------------------
            # 수요일 6교시 자동 제외
            # ------------------------------------------------

            if (
                day_name == "수"
                and row_idx == 6
            ):

                print(
                    "⏭️ 수요일 6교시는 건너뜁니다."
                )

                continue


            # ------------------------------------------------
            # 배열 범위 확인
            # ------------------------------------------------

            grid_row_index = row_idx - 1

            grid_col_index = col_idx - 1


            if (
                grid_row_index
                >= len(timetable_grid)
            ):

                continue


            if (
                grid_col_index
                >= len(
                    timetable_grid[
                        grid_row_index
                    ]
                )
            ):

                continue


            target_text = timetable_grid[
                grid_row_index
            ][
                grid_col_index
            ]


            # ------------------------------------------------
            # 빈 셀
            # ------------------------------------------------

            if not target_text:

                print(
                    f"⏭️ {day_name}요일 "
                    f"{row_idx}교시 = 빈칸"
                )

                continue


            target_text = str(
                target_text
            ).strip()


            # ------------------------------------------------
            # 과목 매핑
            # ------------------------------------------------

            search_text = name_mapping.get(
                target_text,
                target_text
            )


            print(
                f"=== [작업] "
                f"{day_name}요일 "
                f"{row_idx}교시 "
                f"➡️ '{target_text}' "
                f"→ '{search_text}' ==="
            )


            try:

                # =================================================
                # 월요일
                # =================================================

                if day_name == "월":

                    target_xpath = (
                        '//*[contains(@id, "uuid-")]'
                        '/div/div[3]/div/div[2]/div/'
                        'div[1]/div/div/div['
                        f'{row_idx}'
                        ']/div[3]/div/div/div'
                    )


                    element = wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                target_xpath
                            )
                        )
                    )


                    driver.execute_script(
                        (
                            "arguments[0].scrollIntoView("
                            "{block: 'center'});"
                        ),
                        element
                    )


                    width = element.size[
                        "width"
                    ]


                    action = ActionChains(
                        driver
                    )


                    action \
                        .move_to_element(
                            element
                        ) \
                        .move_by_offset(
                            -width,
                            0
                        ) \
                        .context_click() \
                        .perform()


                # =================================================
                # 화 ~ 금
                # =================================================

                else:

                    target_xpath = (
                        '//*[contains(@id, "uuid-")]'
                        '/div/div[3]/div/div[2]/div/'
                        'div[1]/div/div/div['
                        f'{row_idx}'
                        ']/div['
                        f'{col_idx + 1}'
                        ']/div/div/div'
                    )


                    element = wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                target_xpath
                            )
                        )
                    )


                    driver.execute_script(
                        (
                            "arguments[0].scrollIntoView("
                            "{block: 'center'});"
                        ),
                        element
                    )


                    ActionChains(
                        driver
                    ).context_click(
                        element
                    ).perform()


                # =================================================
                # 우클릭 메뉴
                # =================================================

                menu_xpath = (
                    f"//*[contains(text(), "
                    f"'{search_text}')]"
                    f" | "
                    f"//*[contains(@aria-label, "
                    f"'{search_text}')]"
                )


                menu_element = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            menu_xpath
                        )
                    )
                )


                try:

                    menu_element.click()

                except Exception:

                    ActionChains(
                        driver
                    ).move_to_element(
                        menu_element
                    ).click().perform()


                print(
                    "🎯 등록 성공"
                )


            except Exception as item_error:

                print(
                    f"❌ {day_name}요일 "
                    f"{row_idx}교시 처리 실패"
                )

                print(
                    f"   사유: {item_error}"
                )


            # ------------------------------------------------
            # 너무 빠른 연속 입력 방지
            # ------------------------------------------------

##            time.sleep(0.2)


    print()
    print("=" * 70)
    print("🎉 모든 시간표 입력 작업이 완료되었습니다!")
    print("=" * 70)


# ============================================================
# [START]
# ============================================================

def main():

    # --------------------------------------------------------
    # GUI 실행
    # --------------------------------------------------------

    gui_result = show_main_gui()


    if gui_result["status"] != "run":

        print(
            "⏹️ 사용자가 취소했습니다."
        )

        return


    name_mapping = gui_result[
        "data"
    ]


    timetable_grid = gui_result[
        "timetable"
    ]


    print_intro()


    # ========================================================
    # Chrome 사용자 데이터 폴더
    # ========================================================

    if not os.path.exists(
        DEBUG_USER_DATA_DIR
    ):

        os.makedirs(
            DEBUG_USER_DATA_DIR
        )


    # ========================================================
    # [1단계]
    # 9224 디버깅 Chrome 실행
    # ========================================================

    print(
        "\n🌐 [1단계] 디버깅 Chrome 실행"
    )


    try:

        cmd_command = (
            f'"{CHROME_PATH}" '
            f'--remote-debugging-port={DEBUG_PORT} '
            f'--user-data-dir="{DEBUG_USER_DATA_DIR}"'
        )


        subprocess.Popen(
            cmd_command,
            shell=True
        )


        time.sleep(2)


    except Exception as e:

        print(
            f"❌ Chrome 실행 실패: {e}"
        )

        input(
            "종료하려면 Enter를 누르세요..."
        )

        return


    # ========================================================
    # Selenium 연결
    # ========================================================

    options = Options()


    options.add_experimental_option(
        "debuggerAddress",
        f"127.0.0.1:{DEBUG_PORT}"
    )


    try:

        driver = webdriver.Chrome(
            options=options
        )


        print(
            "✅ 9224 디버깅 Chrome에 연결되었습니다."
        )


    except Exception as e:

        print(
            "❌ 디버깅 Chrome 연결 실패"
        )

        print(
            f"   {e}"
        )

        input(
            "종료하려면 Enter를 누르세요..."
        )

        return


    # ========================================================
    # [2단계]
    # 업무포털 자동 접속
    # ========================================================

    try:

        print()
        print(
            "=" * 70
        )

        print(
            "🌐 [2단계] 업무포털 접속"
        )

        print(
            "=" * 70
        )


        driver.get(
            EDUPTL_URL
        )


        time.sleep(2)


        # ====================================================
        # 로그인 버튼
        #
        # 사용자가 알려준 XPath:
        #
        # //*[@id="btnLgn"]/text()
        #
        # Selenium에서는 text() 노드 자체를 클릭할 수 없으므로
        # 실제 클릭 대상은 부모 element인
        #
        # //*[@id="btnLgn"]
        #
        # ====================================================

        login_xpath = (
            '//*[@id="btnLgn"]'
        )


        try:

            login_button = WebDriverWait(
                driver,
                5
            ).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        login_xpath
                    )
                )
            )


            driver.execute_script(
                (
                    "arguments[0].scrollIntoView("
                    "{block: 'center'});"
                ),
                login_button
            )


            try:

                login_button.click()

            except Exception:

                driver.execute_script(
                    "arguments[0].click();",
                    login_button
                )


            print(
                "✅ 로그인 버튼을 클릭했습니다."
            )


        except Exception as e:

            print(
                "⚠️ 로그인 버튼을 자동으로 클릭하지 못했습니다."
            )

            print(
                f"   {e}"
            )

            print(
                "   현재 페이지에서 직접 로그인 버튼을 클릭해도 됩니다."
            )


        # ====================================================
        # 사용자 로그인
        # ====================================================

        print()
        print(
            "=" * 70
        )

        print(
            "🔐 [사용자 작업]"
        )

        print(
            "=" * 70
        )

        print(
            "이제부터 로그인 및 NEIS 이동은 직접 진행하세요."
        )

        print()
        print(
            "① 필요하면 EVPN 인증"
        )

        print(
            "② 업무포털 로그인"
        )

        print(
            "③ NEIS(4세대 나이스) 접속"
        )

        print(
            "④ 시간표 입력 화면까지 이동"
        )

        print()
        print(
            "※ 프로그램은 로그인 정보를 입력하지 않습니다."
        )

        print(
            "※ EVPN 사용 여부와 관계없이 사용 가능합니다."
        )

        print()
        print(
            "시간표 입력 화면까지 이동한 후"
        )

        input(
            "➡️ 준비가 끝났다면 이 터미널에서 [Enter] : "
        )


        # ====================================================
        # [3단계]
        # 현재 열린 Chrome 탭 확인
        # ====================================================

        print()
        print(
            "=" * 70
        )

        print(
            "🔎 [3단계] NEIS 시간표 화면 찾기"
        )

        print(
            "=" * 70
        )


        handles = driver.window_handles


        print(
            f"현재 Chrome 탭/창 개수: {len(handles)}"
        )


        for index, handle in enumerate(
            handles
        ):

            try:

                driver.switch_to.window(
                    handle
                )


                print(
                    f"  [{index}] "
                    f"{driver.title} | "
                    f"{driver.current_url}"
                )

            except Exception:

                pass


        # ====================================================
        # NEIS URL을 포함한 탭 찾기
        # ====================================================

        neis_handle = None


        for handle in driver.window_handles:

            try:

                driver.switch_to.window(
                    handle
                )


                current_url = (
                    driver.current_url
                    or ""
                )


                if (
                    "ice.neis.go.kr"
                    in current_url
                ):

                    neis_handle = handle

                    break


            except Exception:

                continue


        # ====================================================
        # NEIS 탭을 찾은 경우
        # ====================================================

        if neis_handle:

            driver.switch_to.window(
                neis_handle
            )


            print()
            print(
                "🎯 NEIS 탭을 찾았습니다."
            )

            print(
                f"   URL: {driver.current_url}"
            )


        else:

            # ------------------------------------------------
            # NEIS 탭을 못 찾았으면 직접 NEIS 메인으로 이동
            # ------------------------------------------------

            print()
            print(
                "⚠️ NEIS 탭을 자동으로 찾지 못했습니다."
            )

            print(
                "🌐 NEIS 메인 페이지를 엽니다."
            )


            driver.get(
                NEIS_URL
            )


            time.sleep(2)


            print(
                f"📌 현재 URL: {driver.current_url}"
            )


        # ====================================================
        # 사용자가 시간표 화면까지 이동
        # ====================================================

        print()
        print(
            "=" * 70
        )

        print(
            "📅 [4단계] 시간표 화면으로 이동"
        )

        print(
            "=" * 70
        )


        print(
            "현재 Chrome에서 NEIS 시간표 입력 화면으로 이동하세요."
        )

        print()
        print(
            "로그인 / 메뉴 이동 / 시간표 화면 진입은"
        )

        print(
            "모두 직접 진행하면 됩니다."
        )

        print()
        print(
            "시간표 화면이 완전히 열린 후 Enter를 누르세요."
        )


        input(
            "➡️ 시간표 화면 준비 완료 [Enter] : "
        )


        # ====================================================
        # 현재 탭 다시 확인
        # ====================================================

        current_url = (
            driver.current_url
            or ""
        )


        print()
        print(
            "📌 자동 입력 대상:"
        )

        print(
            f"   URL: {current_url}"
        )

        print(
            f"   제목: {driver.title}"
        )


        # ====================================================
        # NEIS 여부 확인
        # ====================================================

        if "ice.neis.go.kr" not in current_url:

            print()
            print(
                "⚠️ 현재 페이지가 NEIS URL이 아닙니다."
            )

            print(
                "현재 페이지에서 계속 진행합니다."
            )


        # ====================================================
        # 최종 확인
        # ====================================================

        print()
        print(
            "=" * 70
        )

        print(
            "🚀 시간표 자동 입력을 시작합니다."
        )

        print(
            "=" * 70
        )


        input(
            "➡️ 시작하려면 Enter를 누르세요..."
        )


        # ====================================================
        # 실제 시간표 자동 입력
        # ====================================================

        run_timetable_macro(
            driver,
            timetable_grid,
            name_mapping
        )


        input(
            "\n프로그램을 종료하려면 [Enter]를 누르세요..."
        )


    except Exception as e:

        print()
        print(
            "=" * 70
        )

        print(
            "❌ 프로그램 실행 중 오류가 발생했습니다."
        )

        print(
            "=" * 70
        )

        print(
            f"{e}"
        )


        input(
            "\n오류를 확인한 후 [Enter]를 누르세요..."
        )


# ============================================================
# 프로그램 시작
# ============================================================

if __name__ == "__main__":

    main()
