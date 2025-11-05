import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Combobox
import os
import json
import datetime
import calendar
import pandas as pd
from conclusion_form.form import ConclusionForm
from search_form.form import SearchForm
import sys, os, json

APP_NAME = "Medoctor"

def appdata_dir():
    base = os.environ.get("APPDATA", os.path.expanduser("~"))
    path = os.path.join(base, APP_NAME)
    os.makedirs(path, exist_ok=True)
    return path

def setup_logging():
    log_dir = os.getcwd()
    log_file = os.path.join(log_dir, "log.txt")

    # Чтобы старые логи не затирались, можно дописывать
    sys.stdout = open(log_file, "a", encoding="utf-8")
    sys.stderr = sys.stdout

    print("\n=== Запуск программы:", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "===\n")

def settings_path():
    return os.path.join(appdata_dir(), "settings.json")

def resource_path(rel_path: str) -> str:
    """
    Возвращает путь к ресурсу и в dev-режиме, и внутри PyInstaller.
    rel_path: относительный путь внутри проекта (например 'conclusion_form/res/template.docx')
    """
    if hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS  # временная папка PyInstaller
    else:
        base = os.path.abspath(".")
    return os.path.join(base, rel_path)

SETTINGS_PATH = settings_path()
XML_PATH      = resource_path("conclusion_form/res/data.xml")
PRIKAZ_XLSX   = resource_path("search_form/input/prikaz29n.xlsx")\

def user_prikaz_path():
    return os.path.join(appdata_dir(), "prikaz29n.xlsx")

def get_prikaz_read_path():
    """
    Путь, откуда читать приказ:
    если в %APPDATA% есть пользовательская копия — берём её,
    иначе — оригинал из пакета.
    """
    up = user_prikaz_path()
    return up if os.path.exists(up) else PRIKAZ_XLSX



def load_settings():
    if os.path.exists(SETTINGS_PATH):
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return {"save_dir": os.getcwd()}

def save_settings(settings):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)

def load_data():
    import xml.etree.ElementTree as ET
    if not os.path.exists(XML_PATH):
        return {}
    tree = ET.parse(XML_PATH)
    root = tree.getroot()
    data = {}
    for person in root.findall("person"):
        org_name = person.findtext("organization", default="")
        record = {
            "name": person.findtext("name", default=""),
            "birthday": person.findtext("birthday", default=""),
            "sex": person.findtext("sex", default=""),
            "division": person.findtext("division", default=""),
            "profession": person.findtext("profession", default=""),
            "factors": person.findtext("factors", default=""),
            "typework": person.findtext("typework", default=""),
            "id": person.findtext("id", default=""),
            "ids_date": person.findtext("ids_date", default=""),
            "diagnosis": person.findtext("diagnosis", default="")
        }
        if org_name not in data:
            data[org_name] = []
        data[org_name].append(record)
    return data

def sanitize_filename(name: str) -> str:
    import re
    return re.sub(r'[\\\/\:\*\?"<>\|]', '_', name)

def is_valid_date(date_str):
    try:
        datetime.datetime.strptime(date_str, "%d.%m.%Y")
        return True
    except ValueError:
        return False

def open_calendar(parent, entry_widget):
    from tkcalendar import Calendar
    import ctypes
    from ctypes import wintypes

    mouse_x = parent.winfo_pointerx()
    mouse_y = parent.winfo_pointery()
    top = tk.Toplevel(parent)
    top.overrideredirect(False)
    top.title("Выберите дату")
    def pick_date():
        date = cal.selection_get()
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, date.strftime("%d.%m.%Y"))
        top.destroy()
    cal = Calendar(top, date_pattern='dd.mm.yyyy')
    cal.pack(padx=10, pady=10)
    tk.Button(top, text="Выбрать", command=pick_date).pack(pady=5)
    top.update_idletasks()
    win_w = top.winfo_width()
    win_h = top.winfo_height()
    try:
        SPI_GETWORKAREA = 0x0030
        rect = wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
        work_w = rect.right - rect.left
        work_h = rect.bottom - rect.top
    except Exception:
        work_w = parent.winfo_screenwidth()
        work_h = parent.winfo_screenheight()
    x = mouse_x + 10
    y = mouse_y + 10
    if x + win_w > work_w:
        x = work_w - win_w
    if y + win_h > work_h:
        y = work_h - win_h
    if x < 0:
        x = 0
    if y < 0:
        y = 0
    top.geometry(f"{win_w}x{win_h}+{x}+{y}")

def show_notification(parent, text, duration=3000, x_offset=10, y_offset=10):
    label = tk.Label(
        parent,
        text=text,
        fg="white",
        bg="#333",
        bd=1,
        relief="solid",
        padx=10, pady=5
    )
    label.place(
        relx=1.0, rely=1.0,
        anchor="se",
        x=-x_offset,
        y=-y_offset
    )
    parent.after(duration, label.place_forget)

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Генератор медицинских заключений")
        self.geometry("680x550")
        self.settings = load_settings()
        self.current_form_frame = None

        self.create_menubar()
        self.settings_window = None
        self.forms_area = tk.Frame(self)
        self.forms_area.pack(fill="both", expand=True)
        self.show_form("search")



    def create_menubar(self):
        menubar = tk.Menu(self)

        # Файл (настройки)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Настройки", command=self.open_settings)
        menubar.add_cascade(label="Файл", menu=file_menu)

        # Формы (переключение форм через меню)
        forms_menu = tk.Menu(menubar, tearoff=0)
        forms_menu.add_command(label="Заключение", command=lambda: self.show_form("conclusion"))
        forms_menu.add_command(label="Анализ по приказу 29н", command=lambda: self.show_form("search"))
        menubar.add_cascade(label="Формы", menu=forms_menu)

        # Отчеты
        reports_menu = tk.Menu(menubar, tearoff=0)
        reports_menu.add_command(label="Отчёт по организации", command=self.report_by_organization)
        reports_menu.add_command(label="Отчёт за месяц", command=self.report_by_month)
        reports_menu.add_command(label="Отчёт по врачам", command=self.report_doctors)
        menubar.add_cascade(label="Отчёты", menu=reports_menu)

        self.config(menu=menubar)

    def report_doctors(self):
        import re, os, datetime, calendar
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        from tkinter import messagebox

        data = load_data()

        # --- Окно выбора дат (как в отчёте за месяц) ---
        rpt = tk.Toplevel(self)
        rpt.title("Отчет по врачам")
        rpt.resizable(False, False)
        padx, pady = 10, 5

        tk.Label(rpt, text="Период с:").grid(row=0, column=0, sticky="w", padx=padx, pady=pady)
        start_var = tk.StringVar()
        start_entry = tk.Entry(rpt, width=20, textvariable=start_var)
        start_entry.grid(row=0, column=1, sticky="w", padx=padx, pady=pady)
        tk.Button(rpt, text="📅", command=lambda: open_calendar(self, start_entry)).grid(row=0, column=2, padx=0,
                                                                                        pady=pady)

        tk.Label(rpt, text="По:").grid(row=1, column=0, sticky="w", padx=padx, pady=pady)
        end_var = tk.StringVar()
        end_entry = tk.Entry(rpt, width=20, textvariable=end_var)
        end_entry.grid(row=1, column=1, sticky="w", padx=padx, pady=pady)
        tk.Button(rpt, text="📅", command=lambda: open_calendar(self, end_entry)).grid(row=1, column=2, padx=0,
                                                                                      pady=pady)

        def on_start_changed(*_):
            s = start_var.get().strip()
            if is_valid_date(s):
                dt = datetime.datetime.strptime(s, "%d.%m.%Y")
                last = calendar.monthrange(dt.year, dt.month)[1]
                end_var.set(f"{last:02d}.{dt.month:02d}.{dt.year}")

        start_var.trace_add("write", on_start_changed)

        # -------- helpers --------
        def parse_points_from_text(text: str):
            if not text:
                return []
            norm = text.replace(';', ',')
            return [p.strip() for p in re.findall(r'\d+(?:\.\d+)?', norm) if p.strip()]

        def calc_age_on(date_birth: str, at_date: str):
            try:
                b = datetime.datetime.strptime(date_birth, "%d.%m.%Y").date()
                d = datetime.datetime.strptime(at_date, "%d.%m.%Y").date()
            except Exception:
                return None
            return d.year - b.year - ((d.month, d.day) < (b.month, b.day))

        def base_point_for_gender_age(sex: str, age):
            if age is None:
                return None
            if (sex or "").strip().upper() == "М":
                return "0.12" if age >= 40 else "0.11"
            else:
                return "0.22" if age >= 40 else "0.21"

        def make_report_doctors():
            start = start_var.get().strip()
            end = end_var.get().strip()

            if not is_valid_date(start) or not is_valid_date(end):
                messagebox.showerror("Ошибка даты", "Даты в формате ДД.ММ.ГГГГ")
                return
            d0 = datetime.datetime.strptime(start, "%d.%m.%Y")
            d1 = datetime.datetime.strptime(end, "%d.%m.%Y")
            if d1 < d0:
                messagebox.showerror("Ошибка", "Конечная дата меньше начальной")
                return

            map_path = get_prikaz_read_path()
            if not os.path.exists(map_path):
                messagebox.showerror("Ошибка", f"Не найден файл: {map_path}")
                return
            df_map = pd.read_excel(map_path)
            df_map['n'] = df_map['n'].astype(str).str.replace(',', '.').str.strip()

            target_cols = [
                "Врач-терапевт 2 каб",
                "Врач-психиатр __ каб",
                "Врач-психиатр-нарколог __ каб",
                "Врач-профпатолог __ каб",
                "Врач-невролог 7 каб",
                "Врач-гинеколог 9 каб",
                "Врач-оториноларинголог 6 каб",
                "Врач-дерматовенеролог 9 каб",
                "Врач-офтальмолог 4 каб",
                "Врач-хирург 8 каб",
                "Врач-стоматолог",
                "ФОГ",
                "Мамография",
                "Спирометрия",
                "Тональная пороговая аудиометрия",
            ]

            doctor_patterns = {
                "Врач-терапевт 2 каб": ["терапевт"],
                "Врач-психиатр __ каб": ["психиатр"],
                "Врач-психиатр-нарколог __ каб": ["нарколог"],
                "Врач-профпатолог __ каб": ["профпатолог"],
                "Врач-невролог 7 каб": ["невролог"],
                "Врач-гинеколог 9 каб": ["гинеколог"],
                "Врач-оториноларинголог 6 каб": ["оториноларинголог", "лор"],
                "Врач-дерматовенеролог 9 каб": ["дерматовенеролог", "дерматолог", "венеролог"],
                "Врач-офтальмолог 4 каб": ["офтальмолог"],
                "Врач-хирург 8 каб": ["хирург"],
                "Врач-стоматолог": ["стоматолог"],
            }
            test_patterns = {
                "ФОГ": ["фог", "флюорограф", "флюорография", "рентген грудной"],
                "Мамография": ["маммограф", "маммография"],
                "Спирометрия": ["спирометр"],
                "Тональная пороговая аудиометрия": ["аудиометр", "тональная пороговая аудиометрия"],
            }

            def split_to_set(series):
                items = []
                for v in series.fillna("").astype(str):
                    parts = [p.strip() for p in v.split(",") if p.strip()]
                    items.extend(parts)
                return set(p.lower() for p in items if p)

            def contains_any(terms_set, patterns):
                return any(p in t for t in terms_set for p in patterns)

            summary_rows = []
            for org_name, recs in data.items():
                for r in recs:
                    ids = (r.get("ids_date") or "").strip()
                    if not ids:
                        continue
                    try:
                        d_ids = datetime.datetime.strptime(ids, "%d.%m.%Y")
                    except ValueError:
                        continue
                    if not (d0 <= d_ids <= d1):
                        continue

                    age = calc_age_on(r.get("birthday", ""), ids)
                    base_pt = base_point_for_gender_age(r.get("sex", ""), age)

                    pts = []
                    pts += parse_points_from_text(r.get("factors", ""))
                    pts += parse_points_from_text(r.get("typework", ""))
                    if base_pt:
                        pts.append(base_pt)

                    items = sorted(set(p for p in pts if p))
                    subset = df_map[df_map['n'].isin(items)] if items else df_map.iloc[0:0]

                    required_doctors = split_to_set(subset.get('doctors_name', pd.Series(dtype=str)))
                    required_inspections = split_to_set(subset.get('inspection', pd.Series(dtype=str)))
                    required_analyses = split_to_set(subset.get('analysis', pd.Series(dtype=str)))

                    row = {
                        "Дата": ids,
                        "ФИО": r.get("name", ""),
                        "Дата рождения": r.get("birthday", ""),
                        "Организация": org_name,
                    }
                    for col, pats in doctor_patterns.items():
                        row[col] = '+' if contains_any(required_doctors, pats) else ''
                    for col, pats in test_patterns.items():
                        has = contains_any(required_inspections, pats) or contains_any(required_analyses, pats)
                        row[col] = '+' if has else ''
                    summary_rows.append(row)

            if not summary_rows:
                messagebox.showinfo("Пустой отчет", "Нет записей за выбранный период.")
                return

            summary_rows.sort(key=lambda x: datetime.datetime.strptime(x["Дата"], "%d.%m.%Y"))
            columns = ["Дата", "ФИО", "Дата рождения", "Организация"] + target_cols
            df = pd.DataFrame(summary_rows, columns=columns)

            save_dir = self.settings.get("save_dir", os.getcwd())
            save_path = os.path.join(save_dir, sanitize_filename(f"Отчет_по_врачам_{start}_{end}.xlsx"))

            try:
                from openpyxl.utils import get_column_letter
                from openpyxl.styles import Alignment
                from openpyxl import load_workbook

                with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Отчет")

                wb = load_workbook(save_path)
                ws = wb["Отчет"]

                # Высота первой строки
                ws.row_dimensions[1].height = 136

                for col_idx, col_name in enumerate(columns, 1):
                    # Вертикальный текст для всех колонок
                    ws.cell(row=1, column=col_idx).alignment = Alignment(
                        textRotation=90,
                        vertical="center",
                        horizontal="center",
                        wrap_text=True
                    )
                    if col_name in target_cols:
                        # Мед. колонки фикс ширина 60
                        ws.column_dimensions[get_column_letter(col_idx)].width = 8
                    else:
                        # Остальные — автоподбор
                        max_len = max(
                            (len(str(cell.value)) if cell.value else 0) for cell in ws[get_column_letter(col_idx)])
                        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

                wb.save(save_path)
                show_notification(self, f"Отчет сохранён: {save_path}")
                rpt.destroy()

            except Exception as e:
                messagebox.showerror("Ошибка записи", f"Не удалось сохранить файл:\n{e}")

        tk.Button(
            rpt,
            text="Сформировать",
            command=make_report_doctors,
            bg="#4CAF50",
            fg="white"
        ).grid(row=2, column=0, columnspan=3, pady=(10, 10), padx=padx, sticky="ew")

        rpt.grid_columnconfigure(1, weight=1)
        rpt.update_idletasks()
        w, h = rpt.winfo_width(), rpt.winfo_height()
        sw, sh = rpt.winfo_screenwidth(), rpt.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        rpt.geometry(f"{w}x{h}+{x}+{y}")

    def create_forms_panel(self):
        panel = tk.Frame(self, bd=1, relief="raised")
        panel.pack(fill="x")
        tk.Label(panel, text="Формы:", font="Arial 10 bold").pack(side="left", padx=10, pady=5)
        tk.Button(panel, text="Заключение", command=lambda: self.show_form("conclusion")).pack(side="left", padx=5, pady=5)

    def show_form(self, form_key):
        if self.current_form_frame is not None:
            self.current_form_frame.destroy()
        if form_key == "conclusion":
            self.current_form_frame = ConclusionForm(self.forms_area, main_app=self)
            self.current_form_frame.pack(fill="both", expand=True)
        elif form_key == "search":
            self.current_form_frame = SearchForm(self.forms_area, main_app=self)
            self.current_form_frame.pack(fill="both", expand=True)

    def open_settings(self):
        # если окно уже существует — поднять и сфокусировать
        if self.settings_window and self.settings_window.winfo_exists():
            self.settings_window.deiconify()
            self.settings_window.lift()
            self.settings_window.focus_force()
            # кратко сделать topmost, чтобы наверняка всплыло
            self.settings_window.attributes('-topmost', True)
            self.settings_window.after(100, lambda: self.settings_window.attributes('-topmost', False))
            return

        # иначе — создать новое и сохранить ссылку
        top = tk.Toplevel(self)
        self.settings_window = top
        top.title("Настройки")
        top.resizable(False, False)
        top.geometry("400x350")

        # при закрытии окна обнуляем ссылку
        def on_close():
            if self.settings_window and self.settings_window.winfo_exists():
                self.settings_window.destroy()
            self.settings_window = None

        top.protocol("WM_DELETE_WINDOW", on_close)

        tk.Label(top, text="Папка для сохранения документов:").pack(anchor="w", padx=10, pady=(10, 0))
        path_var = tk.StringVar(value=self.settings.get("save_dir", os.getcwd()))
        path_entry = tk.Entry(top, textvariable=path_var, width=50)
        path_entry.pack(padx=10, pady=5)

        def select_directory():
            from tkinter import filedialog
            # временно убираем topmost, чтобы диалог не оказался за окном
            self.settings_window.attributes("-topmost", False)
            path = filedialog.askdirectory(parent=self.settings_window)
            # после выбора — вернуть окно наверх
            self.settings_window.lift()
            self.settings_window.focus_force()
            self.settings_window.attributes("-topmost", True)
            self.settings_window.after(100, lambda: self.settings_window.attributes("-topmost", False))

            if path:
                path_var.set(path)

        def save_and_close():
            selected_path = path_var.get()
            if not os.path.exists(selected_path):
                try:
                    os.makedirs(selected_path)
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось создать папку:\n{e}")
                    return
            self.settings["save_dir"] = selected_path
            save_settings(self.settings)
            on_close()

        tk.Button(top, text="Выбрать...", command=select_directory).pack(pady=5)
        tk.Button(top, text="Сохранить", command=save_and_close).pack(pady=(5, 10))

        tk.Button(top, text="Редактировать приказ 29н…", command=self.open_prikaz_for_edit).pack(pady=(5, 0))

    def open_prikaz_for_edit(self):
        import shutil
        dst = user_prikaz_path()
        src = PRIKAZ_XLSX

        try:
            os.makedirs(appdata_dir(), exist_ok=True)
            # если пользовательской копии нет — создаём из оригинала (или пустую, если оригинал исчез)
            if not os.path.exists(dst):
                if os.path.exists(src):
                    shutil.copyfile(src, dst)
                else:
                    # создаём пустой xlsx с нужными колонками, чтобы не упасть при чтении
                    import pandas as pd
                    pd.DataFrame(columns=["n", "doctors_name", "inspection", "analysis"]).to_excel(dst, index=False)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подготовить файл приказа для редактирования:\n{e}")
            return

        # открываем пользовательскую копию в ассоциированном приложении (Excel)
        try:
            os.startfile(dst)
            messagebox.showinfo(
                "Редактирование приказа",
                "Открыта копия приказа в %APPDATA%.\nВсе отчёты будут использовать именно её."
            )
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

    # ============ ОТЧЁТ ПО ОРГАНИЗАЦИИ ============
    def report_by_organization(self):
        data = load_data()
        rpt = tk.Toplevel(self)
        rpt.title("Отчет по организации")
        rpt.resizable(False, False)
        padx, pady = 10, 5
        org_var_report = tk.StringVar()
        tk.Label(rpt, text="Организация:").grid(row=0, column=0, sticky="w", padx=padx, pady=pady)
        org_list = sorted(data.keys())
        org_cb = Combobox(
            rpt,
            values=org_list,
            textvariable=org_var_report,
            width=40,
            state="readonly"
        )
        org_cb.grid(row=0, column=1, padx=padx, pady=pady)
        org_var_report.set("")
        tk.Label(rpt, text="Период с:").grid(row=1, column=0, sticky="w", padx=padx, pady=pady)
        start_var = tk.StringVar()
        start_entry = tk.Entry(rpt, width=20, textvariable=start_var)
        start_entry.grid(row=1, column=1, sticky="w", padx=padx, pady=pady)
        tk.Button(rpt, text="📅", command=lambda: open_calendar(self, start_entry)).grid(row=1, column=2, padx=0, pady=pady)
        tk.Label(rpt, text="По:").grid(row=2, column=0, sticky="w", padx=padx, pady=pady)
        end_var = tk.StringVar()
        end_entry = tk.Entry(rpt, width=20, textvariable=end_var)
        end_entry.grid(row=2, column=1, sticky="w", padx=padx, pady=pady)
        tk.Button(rpt, text="📅", command=lambda: open_calendar(self, end_entry)).grid(row=2, column=2, padx=0, pady=pady)

        def on_start_changed(*_):
            s = start_var.get().strip()
            if is_valid_date(s):
                dt = datetime.datetime.strptime(s, "%d.%m.%Y")
                last = calendar.monthrange(dt.year, dt.month)[1]
                end_var.set(f"{last:02d}.{dt.month:02d}.{dt.year}")
        start_var.trace_add("write", on_start_changed)

        def make_report():
            org_sel = org_var_report.get().strip()
            start = start_entry.get().strip()
            end = end_entry.get().strip()
            if not org_sel:
                messagebox.showerror("Ошибка ввода", "Выберите организацию")
                return
            if not is_valid_date(start) or not is_valid_date(end):
                messagebox.showerror("Ошибка даты", "Даты в формате ДД.ММ.ГГГГ")
                return
            d0 = datetime.datetime.strptime(start, "%d.%m.%Y")
            d1 = datetime.datetime.strptime(end, "%d.%m.%Y")
            if d1 < d0:
                messagebox.showerror("Ошибка", "Конечная дата меньше начальной")
                return
            rows = []
            for r in data.get(org_sel, []):
                ids = r.get("ids_date", "").strip()
                if not ids:
                    continue
                try:
                    d_ids = datetime.datetime.strptime(ids, "%d.%m.%Y")
                except ValueError:
                    continue
                if d0 <= d_ids <= d1:
                    rows.append({
                        "Организация": org_sel,
                        "ФИО": r["name"],
                        "Дата рожд.": r["birthday"],
                        "Пол": r["sex"],
                        "Подразделение": r["division"],
                        "Должность": r["profession"],
                        "Факторы": r["factors"],
                        "Виды работ": r["typework"],
                        "Дата ИДС": ids,
                        "Диагноз": r.get("diagnosis", "")
                    })
            if not rows:
                messagebox.showinfo("Пустой отчет", "Нет записей за выбранный период.")
                return
            df = pd.DataFrame(rows)
            save_dir = self.settings.get("save_dir", os.getcwd())
            fname = sanitize_filename(f"Отчет_{org_sel}_{start}_{end}.xlsx")
            save_path = os.path.join(save_dir, fname)
            try:
                from openpyxl.utils import get_column_letter
                with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Report")
                    sheet = writer.sheets["Report"]
                    for idx, col in enumerate(df.columns, start=1):
                        width = max(df[col].astype(str).map(len).max(), len(col)) + 2
                        sheet.column_dimensions[get_column_letter(idx)].width = width
                show_notification(self, f"Отчет сохранён: {save_path}")
                rpt.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка записи", f"Не удалось сохранить файл:\n{e}")

        tk.Button(
            rpt,
            text="Сформировать",
            command=make_report,
            bg="#4CAF50",
            fg="white"
        ).grid(row=3, column=0, columnspan=3, pady=(10, 10), padx=padx, sticky="ew")
        rpt.grid_columnconfigure(1, weight=1)
        rpt.update_idletasks()
        w, h = rpt.winfo_width(), rpt.winfo_height()
        sw, sh = rpt.winfo_screenwidth(), rpt.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        rpt.geometry(f"{w}x{h}+{x}+{y}")

    # ============ ОТЧЁТ ПО МЕСЯЦУ ============
    def report_by_month(self):
        data = load_data()
        rpt = tk.Toplevel(self)
        rpt.title("Отчет по дате ИДС")
        rpt.resizable(False, False)
        padx, pady = 10, 5
        tk.Label(rpt, text="Период с:").grid(row=0, column=0, sticky="w", padx=padx, pady=pady)
        start_var = tk.StringVar()
        start_entry = tk.Entry(rpt, width=20, textvariable=start_var)
        start_entry.grid(row=0, column=1, sticky="w", padx=padx, pady=pady)
        tk.Button(rpt, text="📅", command=lambda: open_calendar(self, start_entry)).grid(row=0, column=2, padx=0, pady=pady)
        tk.Label(rpt, text="По:").grid(row=1, column=0, sticky="w", padx=padx, pady=pady)
        end_var = tk.StringVar()
        end_entry = tk.Entry(rpt, width=20, textvariable=end_var)
        end_entry.grid(row=1, column=1, sticky="w", padx=padx, pady=pady)
        tk.Button(rpt, text="📅", command=lambda: open_calendar(self, end_entry)).grid(row=1, column=2, padx=0, pady=pady)
        def on_start_changed(*_):
            s = start_var.get().strip()
            if is_valid_date(s):
                dt = datetime.datetime.strptime(s, "%d.%m.%Y")
                last = calendar.monthrange(dt.year, dt.month)[1]
                end_var.set(f"{last:02d}.{dt.month:02d}.{dt.year}")
        start_var.trace_add("write", on_start_changed)
        def make_report_month():
            start = start_var.get().strip()
            end = end_var.get().strip()
            d0 = datetime.datetime.strptime(start, "%d.%m.%Y")
            d1 = datetime.datetime.strptime(end, "%d.%m.%Y")
            rows = []
            for org_name, recs in data.items():
                for r in recs:
                    ids = r.get("ids_date", "").strip()
                    if not ids:
                        continue
                    try:
                        d_ids = datetime.datetime.strptime(ids, "%d.%m.%Y")
                    except ValueError:
                        continue
                    if d0 <= d_ids <= d1:
                        rows.append({
                            "Организация": org_name,
                            "ФИО": r["name"],
                            "Дата рожд.": r["birthday"],
                            "Пол": r["sex"],
                            "Подразделение": r["division"],
                            "Должность": r["profession"],
                            "Факторы": r["factors"],
                            "Виды работ": r["typework"],
                            "Дата ИДС": ids,
                            "Диагноз": r.get("diagnosis", "")
                        })
            if not rows:
                messagebox.showinfo("Пустой отчет", "Нет записей за выбранный период.")
                return
            df = pd.DataFrame(rows)
            save_dir = self.settings.get("save_dir", os.getcwd())
            fname = sanitize_filename(f"Отчет_по_месяцу_{start}_{end}.xlsx")
            save_path = os.path.join(save_dir, fname)
            try:
                from openpyxl.utils import get_column_letter
                with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Report")
                    sheet = writer.sheets["Report"]
                    for idx, col in enumerate(df.columns, start=1):
                        width = max(df[col].astype(str).map(len).max(), len(col)) + 2
                        sheet.column_dimensions[get_column_letter(idx)].width = width
                show_notification(self, f"Отчет сохранён: {save_path}")
                rpt.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка записи", f"Не удалось сохранить файл:\n{e}")
        tk.Button(
            rpt,
            text="Сформировать",
            command=make_report_month,
            bg="#4CAF50",
            fg="white"
        ).grid(row=2, column=0, columnspan=3, pady=(10, 10), padx=padx, sticky="ew")
        rpt.grid_columnconfigure(1, weight=1)
        rpt.update_idletasks()
        w, h = rpt.winfo_width(), rpt.winfo_height()
        sw, sh = rpt.winfo_screenwidth(), rpt.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        rpt.geometry(f"{w}x{h}+{x}+{y}")




if __name__ == "__main__":
    app = MainApp()
    setup_logging()
    app.mainloop()
