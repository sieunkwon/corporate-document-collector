import os
import re
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def normalize_text(text):
    return unicodedata.normalize("NFC", text or "")


def compact_text(text):
    return re.sub(r"[\s_\-*·.()\[\]]", "", normalize_text(text))


def split_names(text):
    return [name.strip() for name in re.split(r"[,，;；\r\n]+", text) if name.strip()]


def safe_name(name):
    # Windows에서 사용할 수 없는 문자까지 함께 제거해 양쪽 OS에서 같은 결과를 만든다.
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip()


def is_month_match(name, month_range):
    digits = re.findall(r"\d+", month_range or "")
    if not digits:
        return False

    normalized_name = normalize_text(name)
    if len(digits) == 1:
        month = re.escape(digits[0])
        pattern = (
            rf"(?<!\d)(?:\(\s*{month}\s*\)\s*월|"
            rf"\(\s*{month}\s*월\s*\)|{month}\s*월)(?!\d)"
        )
    else:
        pattern = (
            rf"(?<!\d){re.escape(digits[0])}\s*(?:월)?\s*[-~]\s*"
            rf"{re.escape(digits[1])}\s*(?:월)?(?!\d)"
        )
    return re.search(pattern, normalized_name) is not None


def starts_with_company(compact_name, company_compact):
    return compact_name.startswith(company_compact) or compact_name.startswith(
        "주" + company_compact
    )
def cohort_numbers(name):
    return {int(value) for value in re.findall(r"(?<!\d)(\d+)\s*기(?!\d)", normalize_text(name))}


def cohort_number(value):
    match = re.search(r"\d+", value or "")
    return int(match.group()) if match else None


def cohort_score(path, selected_cohort, job, policy, company_compact):
    selected = cohort_number(selected_cohort)
    found = cohort_numbers(str(path))
    if policy == "strict":
        if selected and selected > 1 and selected not in found:
            return None
        if found and selected not in found:
            return None
    elif policy == "normal" and found and selected not in found:
        return None

    if selected in found:
        score = 300
    elif not found:
        score = 100
    elif policy == "common":
        score = 20
    else:
        return None

    compact_full_path = compact_text(str(path))
    if company_compact in compact_full_path:
        score += 1000
    if job and compact_text(job) in compact_full_path:
        score += 100
    return score



class DocumentCollector(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("기업서류 PDF 자동 수집기")
        self.geometry("900x840")
        self.minsize(820, 780)
        self.configure(bg="#f5f7fb")
        self.last_result_folder = None

        desktop = Path.home() / "Desktop"
        if not desktop.exists():
            desktop = Path.home()

        self.search_var = tk.StringVar(value=str(Path.home()))
        self.output_var = tk.StringVar(value=str(desktop / "기업별_PDF_수집결과"))
        self.company_var = tk.StringVar()
        self.cohort_var = tk.StringVar(value="1기")
        self.job_var = tk.StringVar()
        self.round_var = tk.StringVar(value="1차")
        self.month_var = tk.StringVar()
        self.participants_var = tk.StringVar()
        self.mentors_var = tk.StringVar()

        self._configure_style()
        self._build_ui()

    def _configure_style(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TFrame", background="#ffffff")
        style.configure("TLabel", background="#ffffff", foreground="#0f172a", font=("", 10))
        style.configure("Field.TLabel", font=("", 10, "bold"))
        style.configure("Hint.TLabel", foreground="#64748b")
        style.configure("Primary.TButton", background="#2563eb", foreground="#ffffff", borderwidth=0, font=("", 10))
        style.map("Primary.TButton", background=[("active", "#1d4ed8"), ("pressed", "#1e40af")], foreground=[("disabled", "#ffffff")])
        style.configure("Result.TButton", background="#ffffff", foreground="#334155", bordercolor="#cbd5e1", lightcolor="#cbd5e1", darkcolor="#cbd5e1", borderwidth=1, font=("", 10))
        style.map("Result.TButton", background=[("active", "#f8fafc")], foreground=[("disabled", "#94a3b8")])
        style.configure("Result.TLabelframe", background="#ffffff", bordercolor="#d8dee9", lightcolor="#d8dee9", darkcolor="#d8dee9")
        style.configure("Result.TLabelframe.Label", background="#ffffff", foreground="#334155")

    def _build_ui(self):
        header = tk.Frame(self, bg="#0f172a", height=92)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="기업서류 PDF 자동 수집기",
            bg="#0f172a",
            fg="white",
            font=("", 18, "bold"),
        ).pack(anchor="w", padx=28, pady=(18, 2))
        tk.Label(
            header,
            text="원본은 그대로 두고, 필요한 PDF 복사본만 새 폴더에 정리합니다.",
            bg="#0f172a",
            fg="#bec8d7",
            font=("", 10),
        ).pack(anchor="w", padx=28)

        card = ttk.Frame(self, padding=(22, 18))
        card.pack(fill="x", padx=24, pady=(20, 8))
        card.columnconfigure(1, weight=1)

        self._folder_row(card, 0, "검색할 폴더", self.search_var)
        self._folder_row(card, 1, "결과 저장 폴더", self.output_var)
        self._entry_row(card, 2, "기업명", self.company_var)

        ttk.Label(card, text="기수 / 직무", style="Field.TLabel").grid(
            row=3, column=0, sticky="w", pady=9
        )
        cohort_frame = ttk.Frame(card)
        cohort_frame.grid(row=3, column=1, columnspan=2, sticky="ew", pady=5)
        cohort_frame.columnconfigure(1, weight=1)
        ttk.Combobox(
            cohort_frame,
            textvariable=self.cohort_var,
            values=("1기", "2기", "3기", "4기", "5기"),
            width=10,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Entry(cohort_frame, textvariable=self.job_var).grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Label(
            cohort_frame, text="직무는 선택 입력  예: 마케팅", style="Hint.TLabel"
        ).grid(row=0, column=2, sticky="w", padx=(12, 0))
        ttk.Label(card, text="수집 차수", style="Field.TLabel").grid(
            row=4, column=0, sticky="w", pady=9
        )
        round_frame = ttk.Frame(card)
        round_frame.grid(row=4, column=1, columnspan=2, sticky="ew", pady=5)
        round_frame.columnconfigure(3, weight=1)
        ttk.Combobox(
            round_frame,
            textvariable=self.round_var,
            values=("1차", "2차"),
            state="readonly",
            width=12,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Label(round_frame, text="월 구간", style="Field.TLabel").grid(
            row=0, column=1, padx=(0, 10)
        )
        ttk.Entry(round_frame, textvariable=self.month_var, width=12).grid(
            row=0, column=2, padx=(0, 12)
        )
        ttk.Label(round_frame, text="예: 7 또는 6-7", style="Hint.TLabel").grid(
            row=0, column=3, sticky="w"
        )

        self._entry_row(card, 5, "참여자명", self.participants_var)
        self._entry_row(card, 6, "멘토명", self.mentors_var)
        ttk.Label(
            card,
            text="참여자와 멘토를 같은 순서로 쉼표 구분하세요.  예: 홍길동, 김철수",
            style="Hint.TLabel",
        ).grid(row=7, column=1, columnspan=2, sticky="w", pady=(2, 10))
        ttk.Label(
            card,
            text="월 구간이 일치하는 인턴형 일경험 결과보고만 수집합니다.",
            style="Hint.TLabel",
        ).grid(row=8, column=1, columnspan=2, sticky="w", pady=(0, 8))

        actions = tk.Frame(self, bg="#f5f7fb")
        actions.pack(fill="x", pady=(4, 8))
        self.run_button = ttk.Button(
            actions, text="PDF 복사본 수집", command=self.collect, style="Primary.TButton"
        )
        self.run_button.pack(side="left", padx=(170, 6), ipadx=35, ipady=8)
        self.open_button = ttk.Button(
            actions, text="결과 폴더 열기", command=self.open_result, state="disabled"
        )
        self.open_button.configure(style="Result.TButton")
        self.open_button.pack(side="left", padx=6, ipadx=28, ipady=8)

        log_frame = ttk.LabelFrame(self, text="수집 결과", padding=(18, 14), style="Result.TLabelframe")
        log_frame.pack(fill="both", expand=True, padx=24, pady=(4, 22))
        self.log = tk.Text(
            log_frame,
            wrap="word",
            relief="flat",
            borderwidth=0,
            bg="#ffffff",
            fg="#334155",
            font=("", 10),
            state="disabled",
        )
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _folder_row(self, parent, row, label, variable):
        ttk.Label(parent, text=label, style="Field.TLabel").grid(
            row=row, column=0, sticky="w", pady=9, padx=(0, 18)
        )
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", pady=6, padx=(0, 10)
        )
        ttk.Button(
            parent,
            text="찾아보기",
            command=lambda: self.pick_folder(variable),
        ).grid(row=row, column=2, sticky="ew", pady=6)

    def _entry_row(self, parent, row, label, variable):
        ttk.Label(parent, text=label, style="Field.TLabel").grid(
            row=row, column=0, sticky="w", pady=9, padx=(0, 18)
        )
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, columnspan=2, sticky="ew", pady=6
        )

    def pick_folder(self, variable):
        initial = variable.get()
        if not Path(initial).is_dir():
            initial = str(Path.home())
        chosen = filedialog.askdirectory(initialdir=initial)
        if chosen:
            variable.set(chosen)

    def set_log(self, text):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("1.0", text)
        self.log.configure(state="disabled")
        self.update_idletasks()

    def open_result(self):
        if not self.last_result_folder or not self.last_result_folder.exists():
            return
        if sys.platform == "darwin":
            subprocess.run(["open", str(self.last_result_folder)], check=False)
        elif os.name == "nt":
            os.startfile(str(self.last_result_folder))
        else:
            subprocess.run(["xdg-open", str(self.last_result_folder)], check=False)

    def validate_inputs(self):
        search_root = Path(self.search_var.get().strip()).expanduser()
        output_text = self.output_var.get().strip()
        company = re.sub(r"^\(주\)", "", normalize_text(self.company_var.get()).strip())
        cohort = self.cohort_var.get().strip()
        job = normalize_text(self.job_var.get()).strip()
        month_range = self.month_var.get().strip()
        participants = split_names(self.participants_var.get())
        mentors = split_names(self.mentors_var.get())

        if not search_root.is_dir():
            raise ValueError("검색할 폴더를 확인해 주세요.")
        if not output_text or not company:
            raise ValueError("결과 폴더와 기업명을 입력해 주세요.")
        if not re.search(r"\d+", month_range):
            raise ValueError("월을 7 또는 6-7과 같은 형식으로 입력해 주세요.")
        if not re.fullmatch(r"\d+\s*기", cohort):
            raise ValueError("기수를 1기, 2기와 같은 형식으로 입력해 주세요.")
        if not participants or not mentors:
            raise ValueError("참여자명과 멘토명을 입력해 주세요.")
        if len(mentors) != 1 and len(mentors) != len(participants):
            raise ValueError("멘토는 한 명만 입력하거나 참여자 수와 같게 입력해 주세요.")

        return (
            search_root,
            Path(output_text).expanduser(),
            company,
            cohort,
            job,
            self.round_var.get(),
            month_range,
            participants,
            mentors,
        )

    def collect(self):
        try:
            (
                search_root,
                output_root,
                company,
                cohort,
                job,
                round_name,
                month_range,
                participants,
                mentors,
            ) = self.validate_inputs()
        except ValueError as error:
            messagebox.showwarning("입력 확인", str(error))
            return

        self.run_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.set_log("PDF 파일을 검색하는 중입니다...")

        try:
            output_resolved = output_root.resolve(strict=False)
            all_pdfs = []
            for folder, _, files in os.walk(search_root):
                folder_path = Path(folder)
                try:
                    if folder_path.resolve(strict=False).is_relative_to(output_resolved):
                        continue
                except ValueError:
                    pass
                for filename in files:
                    if filename.lower().endswith(".pdf"):
                        all_pdfs.append(folder_path / filename)

            company_compact = compact_text(company)
            company_files = [
                path
                for path in all_pdfs
                if company_compact in compact_text(str(path))
            ]

            operation_name = "_".join(value for value in (job, cohort) if value)
            result_folder = output_root / safe_name(f"(주){company}") / safe_name(operation_name)
            result_folder.mkdir(parents=True, exist_ok=True)
            overall = [
                f"기업: (주){company}",
                f"기수: {cohort}",
                f"직무: {job or '미입력'}",
                f"차수: {round_name}",
                f"월 구간: {month_range}",
                "",
            ]
            total_found = 0
            total_missing = 0

            for index, participant in enumerate(participants):
                mentor = mentors[0] if len(mentors) == 1 else mentors[index]
                participant_compact = compact_text(participant)
                mentor_compact = compact_text(mentor)

                def needs_second_round(compact_name):
                    return round_name != "2차" or "2차" in compact_name

                requirements = [
                    (
                        1,
                        "인턴형 일경험 실시보고",
                        lambda p: starts_with_company(compact_text(p.stem), company_compact)
                        and re.search(r"인턴형?일경험실시보고", compact_text(p.stem))
                        and not is_month_match(p.stem, month_range),
                    ),
                    (
                        2,
                        f"({month_range})월 인턴형 일경험 결과보고",
                        lambda p: is_month_match(p.stem, month_range)
                        and re.search(r"인턴형?일경험결과보고", compact_text(p.stem)),
                    ),
                    (
                        3,
                        f"출석부_{participant}_{round_name}",
                        lambda p: "출석부" in compact_text(p.stem)
                        and participant_compact in compact_text(p.stem)
                        and round_name in compact_text(p.stem),
                    ),
                    (
                        4,
                        "기업통장사본",
                        lambda p: re.search(
                            r"기업통장사본|국민통장사본", compact_text(p.stem)
                        ),
                    ),
                    (
                        5,
                        f"참여자({participant}) 멘토({mentor}) 면담일지",
                        lambda p: "면담일지" in compact_text(p.stem)
                        and participant_compact in compact_text(p.stem)
                        and mentor_compact in compact_text(p.stem)
                        and needs_second_round(compact_text(p.stem)),
                    ),
                    (
                        6,
                        f"멘토수당 신청서_{mentor}",
                        lambda p: "멘토수당신청서" in compact_text(p.stem)
                        and mentor_compact in compact_text(p.stem)
                        and needs_second_round(compact_text(p.stem)),
                    ),
                    (
                        7,
                        f"멘토({mentor}) 증빙묶음",
                        lambda p: all(
                            value in compact_text(p.stem)
                            for value in (
                                "멘토",
                                mentor_compact,
                                "통장사본",
                                "신분증사본",
                                "재직증명서",
                            )
                        ),
                    ),
                    (
                        8,
                        f"인턴형 프로그램 표준계약서_{participant}",
                        lambda p: "인턴형프로그램표준계약서" in compact_text(p.stem)
                        and participant_compact in compact_text(p.stem),
                    ),
                    (
                        9,
                        f"참여자({participant}) 증빙묶음",
                        lambda p: all(
                            value in compact_text(p.stem)
                            for value in (
                                "참여자",
                                participant_compact,
                                "통장사본",
                                "신분증사본",
                                "주민등록등본",
                                "길찾기결과",
                            )
                        ),
                    ),
                    (
                        10,
                        f"개인정보 활용에 관한 동의서_{participant}",
                        lambda p: all(
                            value in compact_text(p.stem)
                            for value in (
                                "개인정보",
                                "동의서",
                                participant_compact,
                            )
                        ),
                    ),
                ]

                if round_name == "2차":
                    requirements.extend(
                        [
                            (
                                11,
                                "인턴형 일경험 실시 결과보고",
                                lambda p: "인턴형일경험실시결과보고"
                                in compact_text(p.stem),
                            ),
                            (
                                12,
                                f"인턴형 일경험 종합 보고서_참여자({participant})",
                                lambda p: "인턴형일경험종합보고서"
                                in compact_text(p.stem)
                                and "참여자" in compact_text(p.stem)
                                and participant_compact in compact_text(p.stem),
                            ),
                            (
                                13,
                                f"멘토({mentor}) 결과보고서",
                                lambda p: "멘토" in compact_text(p.stem)
                                and mentor_compact in compact_text(p.stem)
                                and "결과보고서" in compact_text(p.stem),
                            ),
                        ]
                    )

                participant_folder = result_folder / safe_name(participant)
                participant_folder.mkdir(parents=True, exist_ok=True)
                found_count = 0
                missing_labels = []

                for order, label, matcher in requirements:
                    company_orders = {1, 2, 4, 11}
                    common_orders = {4, 7}
                    strict_orders = {1, 2, 11}
                    pool = company_files if order in company_orders else all_pdfs
                    policy = "common" if order in common_orders else ("strict" if order in strict_orders else "normal")
                    ranked_matches = []
                    for path in pool:
                        if not matcher(path):
                            continue
                        score = cohort_score(path, cohort, job, policy, company_compact)
                        if score is not None:
                            ranked_matches.append((score, path.stat().st_mtime, path))
                    ranked_matches.sort(key=lambda item: (item[0], item[1]), reverse=True)
                    matches = [item[2] for item in ranked_matches]
                    if matches:
                        chosen = matches[0]
                        for old_file in participant_folder.glob(f"{order}. *.pdf"):
                            old_file.unlink()
                        target = participant_folder / safe_name(f"{order}. {chosen.name}")
                        shutil.copy2(chosen, target)
                        found_count += 1
                        total_found += 1
                    else:
                        missing_labels.append(f"{order}. {label}")
                        total_missing += 1

                overall.append(
                    f"[{participant} / {mentor}] 수집 {found_count}개, "
                    f"누락 {len(missing_labels)}개"
                )
                if missing_labels:
                    overall.extend(f"  누락: {label}" for label in missing_labels)
                else:
                    overall.append("  누락 없음")
                overall.append("")

            overall.extend(
                ["", f"전체 수집: {total_found}개 / 전체 누락: {total_missing}개"]
            )
            self.last_result_folder = result_folder
            self.open_button.configure(state="normal")
            self.set_log(os.linesep.join(overall))
            messagebox.showinfo(
                "완료",
                f"참여자별 수집 완료: {total_found}개 / 누락 {total_missing}개",
            )
        except Exception as error:
            self.set_log(f"오류: {error}")
            messagebox.showerror("오류", str(error))
        finally:
            self.run_button.configure(state="normal")


if __name__ == "__main__":
    DocumentCollector().mainloop()
