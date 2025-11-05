import tkinter as tk
from tkinter import messagebox, BooleanVar
from tkinter.ttk import Combobox
from tkcalendar import Calendar
import datetime
import xml.etree.ElementTree as ET
from docx import Document
from docxcompose.composer import Composer
import tempfile
import shutil
import re
import xml.dom.minidom as minidom
import ctypes
from ctypes import wintypes
import calendar

import sys, os, json

APP_NAME = "Medoctor"


def setup_logging():
    log_dir = os.getcwd()
    log_file = os.path.join(log_dir, "log.txt")

    # Чтобы старые логи не затирались, можно дописывать
    sys.stdout = open(log_file, "a", encoding="utf-8")
    sys.stderr = sys.stdout

    print("\n=== Запуск программы:", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "===\n")

def appdata_dir():
    base = os.environ.get("APPDATA", os.path.expanduser("~"))
    path = os.path.join(base, APP_NAME)
    os.makedirs(path, exist_ok=True)
    return path

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

SETTINGS_PATH = settings_path()  # теперь в %APPDATA%\Medoctor\settings.json
TEMPLATE_PATH = resource_path("conclusion_form/res/template.docx")
USER_XML_PATH = resource_path("conclusion_form/res/data.xml")
CALENDAR_PNG  = resource_path("conclusion_form/res/calendar.png")
PRIKAZ_XLSX   = resource_path("search_form/input/prikaz29n.xlsx")
SUMMER_XLSX   = resource_path("search_form/input/summer.xlsx")

class ConclusionForm(tk.Frame):
    def __init__(self, parent, main_app):
        super().__init__(parent)
        self.main_app = main_app
        self.settings = main_app.settings
        self.suggestion_listbox = None
        self.report_org_window = None
        self.report_month_window = None

        # --- Переменные формы ---
        self.type_var = tk.StringVar(value="предварительный")
        self.organization = tk.StringVar()
        self.sex_var = tk.StringVar(value="М")
        self.division = tk.StringVar()
        self.profession = tk.StringVar()
        self.factors = tk.StringVar()
        self.typework = tk.StringVar()
        self.diagnosis = tk.StringVar()
        self.combine_all = BooleanVar(value=False)

        # Данные из XML
        self.data = self.load_data()

        # --- UI ---
        self.build_ui()



    # --------- UI строим здесь -------------
    def build_ui(self):
        self.pack(fill="both", expand=True)
        self.columnconfigure(0, weight=1)

        row = 0
        tk.Label(self, text="Дата ИДС").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.ids_frame = tk.Frame(self)
        self.ids_frame.grid(row=row, column=1, sticky="w", padx=10)
        self.ids_entry = tk.Entry(self.ids_frame, width=45)
        self.ids_entry.pack(side="left", fill="x", expand=True)
        self.ids_entry.bind("<KeyRelease>", self.format_date)
        self.calendar_icon = tk.PhotoImage(master=self, file=CALENDAR_PNG  )
        tk.Button(
            self.ids_frame,
            image=self.calendar_icon,
            command=lambda: self.open_calendar(self.ids_entry),
            bd=0,
            relief="flat"
        ).pack(side="left", padx=(5, 0))

        row += 1
        tk.Label(self, text="Тип осмотра").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.type_cb = Combobox(self, textvariable=self.type_var, values=["предварительный", "периодический"], width=50, state="readonly")
        self.type_cb.grid(row=row, column=1, padx=10, pady=(10, 0))

        row += 1
        tk.Label(self, text="Организация").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.organization_cb = Combobox(self, textvariable=self.organization, width=50)
        self.organization_cb.grid(row=row, column=1, padx=10)
        self.organization_cb.bind("<<ComboboxSelected>>", self.on_organization_selected)

        row += 1
        tk.Label(self, text="ФИО").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.name_entry = tk.Entry(self, width=53)
        self.name_entry.grid(row=row, column=1, padx=10)
        self.name_entry.bind("<FocusOut>", lambda e: self.sex_var.set(self.detect_sex_from_name(self.name_entry.get())))
        self.name_entry.bind("<KeyRelease>", self.show_name_suggestions)

        row += 1
        tk.Label(self, text="Дата рождения").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.birthday_frame = tk.Frame(self)
        self.birthday_frame.grid(row=row, column=1, sticky="w", padx=10)
        self.birthday_entry = tk.Entry(self.birthday_frame, width=45)
        self.birthday_entry.pack(side="left", fill="x", expand=True)
        self.birthday_entry.bind("<KeyRelease>", self.format_date)
        tk.Button(
            self.birthday_frame,
            image=self.calendar_icon,
            command=lambda: self.open_calendar(self.birthday_entry),
            bd=0,
            relief="flat"
        ).pack(side="left", padx=(5, 0))

        row += 1
        tk.Label(self, text="Пол").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.sex_cb = Combobox(self, textvariable=self.sex_var, values=["М", "Ж"], width=50, state="readonly")
        self.sex_cb.grid(row=row, column=1, padx=10)

        row += 1
        tk.Label(self, text="Подразделение").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.division_cb = Combobox(self, textvariable=self.division, width=50)
        self.division_cb.grid(row=row, column=1, padx=10)

        row += 1
        tk.Label(self, text="Должность").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.profession_cb = Combobox(self, textvariable=self.profession, width=50)
        self.profession_cb.grid(row=row, column=1, padx=10)

        row += 1
        tk.Label(self, text="Факторы").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.factors_cb = Combobox(self, textvariable=self.factors, width=50)
        self.factors_cb.grid(row=row, column=1, padx=10)

        row += 1
        tk.Label(self, text="Виды работ").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.typework_cb = Combobox(self, textvariable=self.typework, width=50)
        self.typework_cb.grid(row=row, column=1, padx=10)

        row += 1
        tk.Label(self, text="Диагноз").grid(row=row, column=0, sticky="w", padx=10, pady=(10, 0))
        self.diagnosis_cb = Combobox(self, textvariable=self.diagnosis, width=50)
        self.diagnosis_cb.grid(row=row, column=1, padx=10)
        self.diagnosis_cb.bind('<KeyRelease>', self.on_keyrelease)


        row += 1
        self.create_btn = tk.Button(self, text="Создать документ", command=self.generate_document, bg="#4CAF50", fg="white", height=2)
        self.create_btn.grid(row=row, column=0, columnspan=2, padx=10, pady=20, sticky="ew")

        row += 1
        self.combine_all = tk.BooleanVar(value=True)
        tk.Checkbutton(
            self,
            text="Объединить все в один файл",
            variable=self.combine_all
        ).grid(row=row, column=0, columnspan=2, pady=(0, 10), padx=10, sticky="w")

        # Уведомления (скрытый лейбл)
        row += 1
        self.notification_label = tk.Label(
            self,
            text="",
            fg="white",
            bg="#333",
            bd=1,
            relief="solid",
            padx=10, pady=5
        )
        self.notification_label.place_forget()

        # Автозаполнение
        self.organization_cb['values'] = sorted(self.data.keys())
        self.organization_cb.all_values = list(self.organization_cb['values'])

        for cb in (self.organization_cb, self.division_cb, self.profession_cb, self.factors_cb, self.typework_cb):
            cb.bind('<KeyRelease>', self.on_keyrelease)

        setup_logging()

    # ---------------- Логика ---------------------
    @staticmethod
    def sanitize_filename(name: str) -> str:
        return re.sub(r'[\\\/\:\*\?"<>\|]', '_', name)

    def show_name_suggestions(self, event):
        if self.suggestion_listbox:
            self.suggestion_listbox.destroy()
            self.suggestion_listbox = None

        text = self.name_entry.get().strip().lower()
        if not text:
            return

        all_names = []
        for records in self.data.values():
            for rec in records:
                fio = rec.get("name", "")
                if fio and fio.lower().startswith(text):
                    all_names.append(fio)
        suggestions = sorted(set(all_names))[:10]
        if not suggestions:
            return

        x = self.name_entry.winfo_rootx()
        y = self.name_entry.winfo_rooty() + self.name_entry.winfo_height()
        w = self.name_entry.winfo_width()
        h = min(200, len(suggestions) * 20)

        self.suggestion_listbox = tk.Toplevel(self)
        self.suggestion_listbox.overrideredirect(True)
        self.suggestion_listbox.transient(self)
        self.suggestion_listbox.geometry(f"{w}x{h}+{x}+{y}")
        self.suggestion_listbox.lift()

        lb = tk.Listbox(self.suggestion_listbox, exportselection=False)
        lb.pack(fill="both", expand=True)
        for item in suggestions:
            lb.insert(tk.END, item)

        def on_select(evt):
            sel = lb.get(lb.curselection())
            self.suggestion_listbox.destroy()
            self.suggestion_listbox = None
            self.fill_person_fields(sel)
            self.focus_force()
        lb.bind("<ButtonRelease-1>", on_select)

    def fill_person_fields(self, fio):
        for records in self.data.values():
            for rec in records:
                if rec.get("name") == fio:
                    self.name_entry.delete(0, tk.END)
                    self.name_entry.insert(0, rec["name"])
                    self.birthday_entry.delete(0, tk.END)
                    self.birthday_entry.insert(0, rec["birthday"])
                    self.sex_var.set(rec["sex"])
                    self.ids_entry.delete(0, tk.END)
                    self.ids_entry.insert(0, rec.get("ids_date", ""))
                    return

    @staticmethod
    def detect_sex_from_name(full_name):
        parts = full_name.strip().split()
        if len(parts) >= 3:
            middle = parts[2].lower()
            if middle.endswith(("вич", "льич", "ич")):
                return "М"
            elif middle.endswith(("вна", "чна", "инична", "овна", "евна", "ична")):
                return "Ж"
        return "М"

    def load_settings(self):
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        else:
            return {"save_dir": os.getcwd()}

    def save_settings(self, settings):
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)

    def clear_form(self):
        self.organization.set("")
        self.name_entry.delete(0, tk.END)
        self.birthday_entry.delete(0, tk.END)
        self.sex_var.set("М")
        self.division.set("")
        self.profession.set("")
        self.factors.set("")
        self.typework.set("")
        self.diagnosis.set("")
        self.ids_entry.delete(0, tk.END)

    def load_data(self):
        if not os.path.exists(USER_XML_PATH):
            return {}
        tree = ET.parse(USER_XML_PATH)
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

    def prettify_xml(self, xml_path):
        xml_str = open(xml_path, "r", encoding="utf-8").read()
        dom = minidom.parseString(xml_str)
        pretty_xml = dom.toprettyxml(indent="  ", encoding="utf-8")
        with open(xml_path, "wb") as f:
            f.write(pretty_xml)

    def save_record(self, org_name, division, profession, factors, typework,
                    name=None, birthday=None, sex_val=None, diagnosis=None, ids_date=None):
        if not (name and birthday and sex_val):
            return
        if os.path.exists(USER_XML_PATH):
            tree = ET.parse(USER_XML_PATH)
            root = tree.getroot()
        else:
            root = ET.Element("data")
            tree = ET.ElementTree(root)
        now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        p = ET.SubElement(root, "person")
        ET.SubElement(p, "organization").text = org_name
        ET.SubElement(p, "name").text = name
        ET.SubElement(p, "birthday").text = birthday
        ET.SubElement(p, "sex").text = sex_val
        ET.SubElement(p, "division").text = division
        ET.SubElement(p, "profession").text = profession
        ET.SubElement(p, "factors").text = factors
        ET.SubElement(p, "typework").text = typework
        ET.SubElement(p, "diagnosis").text = diagnosis if diagnosis else ""
        ET.SubElement(p, "ids_date").text = ids_date if ids_date else ""
        ET.SubElement(p, "id").text = str(int(datetime.datetime.now().timestamp()))
        ET.SubElement(p, "date_created").text = now_str
        tree.write(USER_XML_PATH, encoding="utf-8", xml_declaration=True)
        self.prettify_xml(USER_XML_PATH)

    def get_unique_values(self, field, org_name=None):
        """Вернёт уникальные значения поля. Если задана org_name — только для этой организации."""
        values = set()

        if org_name and org_name in self.data:
            records = self.data[org_name]
        else:
            # все записи по всем организациям
            records = [rec for recs in self.data.values() for rec in recs]

        for record in records:
            val = record.get(field)
            if val:
                values.add(val)

        return sorted(values)

    def replace_placeholders(self, doc, data_dict):
        def replace_in_paragraph(paragraph, data_dict):
            text = ''.join(run.text for run in paragraph.runs)
            parts = re.split(r'(\{.*?\})', text)
            if not any(part in data_dict for part in parts):
                return
            paragraph.clear()
            for part in parts:
                if part in data_dict:
                    run = paragraph.add_run(data_dict[part])
                    run.bold = True
                else:
                    paragraph.add_run(part)

        for paragraph in doc.paragraphs:
            replace_in_paragraph(paragraph, data_dict)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_in_paragraph(paragraph, data_dict)

    def on_keyrelease(self, event):
        cb = event.widget
        txt = cb.get().lower()
        if not hasattr(cb, 'all_values'):
            cb.all_values = list(cb['values'])
        if txt == '':
            vals = cb.all_values
        else:
            vals = [v for v in cb.all_values if txt in v.lower()]
        cb['values'] = vals
        try:
            cb.tk.call('ttk::combobox::post', cb._w)
            cb.focus_force()
        except Exception:
            pass

    def generate_document(self):
        type_raw = self.type_var.get()
        type_genitive = {
            "предварительный": "предварительного",
            "периодический": "периодического"
        }.get(type_raw, type_raw)

        form_data = {
            "{type}": type_genitive,
            "{organization}": self.organization.get(),
            "{name}": self.name_entry.get(),
            "{birthday}": self.birthday_entry.get(),
            "{sex}": self.sex_var.get(),
            "{division}": self.division.get(),
            "{profession}": self.profession.get(),
            "{factors}": self.factors.get(),
            "{typework}": self.typework.get(),
            "{ids_date}": self.ids_entry.get(),
            "{diagnosis}": self.diagnosis.get(),
            "{year}": str(datetime.datetime.now().year)   # текущий год
        }

        if (not form_data["{organization}"] or
                not form_data["{name}"] or
                not form_data["{birthday}"] or
                not form_data["{ids_date}"]):
            messagebox.showerror(
                "Ошибка ввода",
                "Пожалуйста, заполните обязательные поля:\n"
                "• Организация\n"
                "• ФИО\n"
                "• Дата рождения\n"
                "• Дата ИДС"
            )
            return

        if not self.is_valid_date(self.birthday_entry.get()):
            messagebox.showerror(
                "Ошибка даты",
                "Дата рождения должна быть в формате ДД.ММ.ГГГГ"
            )
            return
        if self.ids_entry.get().strip() and not self.is_valid_date(self.ids_entry.get()):
            messagebox.showerror(
                "Ошибка даты",
                "Дата ИДС должна быть в формате ДД.ММ.ГГГГ"
            )
            return

        temp_doc = Document(TEMPLATE_PATH)
        self.replace_placeholders(temp_doc, form_data)

        temp_doc_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
        temp_doc.save(temp_doc_path)

        if self.combine_all.get():
            combined_filename = os.path.join(
                self.settings.get("save_dir", os.getcwd()),
                f"заключения_{datetime.datetime.now().strftime('%d.%m.%Y')}.docx"
            )

            if os.path.exists(combined_filename):
                try:
                    master_doc = Document(combined_filename)
                    composer = Composer(master_doc)
                    composer.append(Document(temp_doc_path))
                    composer.save(combined_filename)
                    self.show_notification(f"Добавлено в файл: {combined_filename}")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось дописать в файл:\n{e}")
                    return
            else:
                # создаём новый файл из temp_doc
                try:
                    shutil.copyfile(temp_doc_path, combined_filename)
                    self.show_notification(f"Создан новый файл: {combined_filename}")
                    print("создан новый файл для дозаписи")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось создать файл:\n{e}")
                    print("файл не создан")
                    return

        else:
            filename = os.path.join(
                self.settings.get("save_dir", os.getcwd()),
                f"{form_data['{name}']} - заключение.docx"
            )
            try:
                temp_doc.save(filename)
                self.show_notification(f"Файл сохранён: {filename}")
            except PermissionError:
                messagebox.showerror(
                    "Ошибка записи",
                    "Невозможно создать запись, проверьте, закрыт ли Word-файл."
                )
                return
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
                return

        self.save_record(
            form_data["{organization}"],
            form_data["{division}"],
            form_data["{profession}"],
            form_data["{factors}"],
            form_data["{typework}"],
            name=form_data["{name}"],
            birthday=form_data["{birthday}"],
            sex_val=form_data["{sex}"],
            ids_date=self.ids_entry.get(),
            diagnosis=self.diagnosis.get()
        )
        self.data = self.load_data()
        self.organization_cb['values'] = sorted(self.data.keys())
        self.update_comboboxes()
        try:
            os.remove(temp_doc_path)
        except Exception as e:
            print(f"Не удалось удалить временный файл: {temp_doc_path}\n{e}")
        self.clear_form()

    def update_comboboxes(self):
        # всегда обновляем список организаций
        self.organization_cb["values"] = sorted(self.data.keys())

        org = self.organization.get().strip()
        if org and org in self.data:
            # значения ТОЛЬКО для выбранной организации
            self.division_cb["values"] = self.get_unique_values("division", org)
            self.profession_cb["values"] = self.get_unique_values("profession", org)
            self.factors_cb["values"] = self.get_unique_values("factors", org)
            self.typework_cb["values"] = self.get_unique_values("typework", org)
            self.diagnosis_cb["values"] = self.get_unique_values("diagnosis", org)
        else:
            # если организация не выбрана — показываем общие списки (как раньше)
            self.division_cb["values"] = self.get_unique_values("division")
            self.profession_cb["values"] = self.get_unique_values("profession")
            self.factors_cb["values"] = self.get_unique_values("factors")
            self.typework_cb["values"] = self.get_unique_values("typework")
            self.diagnosis_cb["values"] = self.get_unique_values("diagnosis")

        # обновляем кеш для живого поиска в combobox'ах
        for cb in (self.organization_cb, self.division_cb, self.profession_cb, self.factors_cb, self.typework_cb,
                   self.diagnosis_cb):
            cb.all_values = list(cb['values'])

    def on_organization_selected(self, event):
        self.update_comboboxes()

    @staticmethod
    def is_valid_date(date_str):
        try:
            datetime.datetime.strptime(date_str, "%d.%m.%Y")
            return True
        except ValueError:
            return False

    def format_date(self, event):
        widget = event.widget
        s = widget.get()
        digits = ''.join(filter(str.isdigit, s))[:8]
        parts = []
        if len(digits) >= 2:
            parts.append(digits[:2])
        else:
            parts.append(digits)
        if len(digits) >= 4:
            parts.append(digits[2:4])
        elif len(digits) > 2:
            parts.append(digits[2:])
        if len(digits) > 4:
            parts.append(digits[4:])
        new_text = '.'.join(parts)
        if new_text != s:
            widget.delete(0, tk.END)
            widget.insert(0, new_text)
            widget.icursor(tk.END)

    def open_calendar(self, entry_widget):
        mouse_x = self.winfo_pointerx()
        mouse_y = self.winfo_pointery()
        top = tk.Toplevel(self)
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
            work_w = self.winfo_screenwidth()
            work_h = self.winfo_screenheight()
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

    def show_notification(self, text, duration=3000, x_offset=10, y_offset=10):
        self.notification_label.config(text=text)
        self.notification_label.update_idletasks()
        self.notification_label.place(
            relx=1.0, rely=1.0,
            anchor="se",
            x=-x_offset,
            y=-y_offset
        )
        self.after(duration, self.notification_label.place_forget)

    # # ----------------- ОТЧЁТЫ --------------------
    # def report_by_organization(self):
    #     # -- Твой старый код "report_by_organization", только self везде --
    #     import pandas as pd
    #     if self.report_org_window and self.report_org_window.winfo_exists():
    #         self.report_org_window.focus_force()
    #         return
    #     self.report_org_window = tk.Toplevel(self)
    #     rpt = self.report_org_window
    #     rpt.title("Отчет по организации")
    #     rpt.resizable(False, False)
    #     rpt.protocol("WM_DELETE_WINDOW", lambda: (rpt.destroy(), self.set_report_none()))
    #     padx, pady = 10, 5
    #     org_var_report = tk.StringVar()
    #     tk.Label(rpt, text="Организация:").grid(row=0, column=0, sticky="w", padx=padx, pady=pady)
    #     org_list = sorted(self.data.keys())
    #     org_cb = Combobox(
    #         rpt,
    #         values=org_list,
    #         textvariable=org_var_report,
    #         width=40,
    #         state="readonly"
    #     )
    #     org_cb.grid(row=0, column=1, padx=padx, pady=pady)
    #     org_var_report.set("")
    #     tk.Label(rpt, text="Период с:").grid(row=1, column=0, sticky="w", padx=padx, pady=pady)
    #     start_var = tk.StringVar()
    #     start_entry = tk.Entry(rpt, width=20, textvariable=start_var)
    #     start_entry.grid(row=1, column=1, sticky="w", padx=padx, pady=pady)
    #     tk.Button(rpt, text="📅", command=lambda: self.open_calendar(start_entry)).grid(row=1, column=2, padx=0, pady=pady)
    #     tk.Label(rpt, text="По:").grid(row=2, column=0, sticky="w", padx=padx, pady=pady)
    #     end_var = tk.StringVar()
    #     end_entry = tk.Entry(rpt, width=20, textvariable=end_var)
    #     end_entry.grid(row=2, column=1, sticky="w", padx=padx, pady=pady)
    #     tk.Button(rpt, text="📅", command=lambda: self.open_calendar(end_entry)).grid(row=2, column=2, padx=0, pady=pady)
    #
    #     def on_start_changed(*_):
    #         s = start_var.get().strip()
    #         if self.is_valid_date(s):
    #             dt = datetime.datetime.strptime(s, "%d.%m.%Y")
    #             last = calendar.monthrange(dt.year, dt.month)[1]
    #             end_var.set(f"{last:02d}.{dt.month:02d}.{dt.year}")
    #
    #     start_var.trace_add("write", on_start_changed)
    #
    #     def make_report():
    #         org_sel = org_var_report.get().strip()
    #         start = start_entry.get().strip()
    #         end = end_entry.get().strip()
    #         if not org_sel:
    #             messagebox.showerror("Ошибка ввода", "Выберите организацию")
    #             return
    #         if not self.is_valid_date(start) or not self.is_valid_date(end):
    #             messagebox.showerror("Ошибка даты", "Даты в формате ДД.ММ.ГГГГ")
    #             return
    #         d0 = datetime.datetime.strptime(start, "%d.%m.%Y")
    #         d1 = datetime.datetime.strptime(end, "%d.%m.%Y")
    #         if d1 < d0:
    #             messagebox.showerror("Ошибка", "Конечная дата меньше начальной")
    #             return
    #         rows = []
    #         for r in self.data.get(org_sel, []):
    #             ids = r.get("ids_date", "").strip()
    #             if not ids:
    #                 continue
    #             try:
    #                 d_ids = datetime.datetime.strptime(ids, "%d.%m.%Y")
    #             except ValueError:
    #                 continue
    #             if d0 <= d_ids <= d1:
    #                 rows.append({
    #                     "Организация": org_sel,
    #                     "ФИО": r["name"],
    #                     "Дата рожд.": r["birthday"],
    #                     "Пол": r["sex"],
    #                     "Подразделение": r["division"],
    #                     "Должность": r["profession"],
    #                     "Факторы": r["factors"],
    #                     "Виды работ": r["typework"],
    #                     "Дата ИДС": ids,
    #                     "Диагноз": r.get("diagnosis", "")
    #                 })
    #         if not rows:
    #             messagebox.showinfo("Пустой отчет", "Нет записей за выбранный период.")
    #             return
    #         df = pd.DataFrame(rows)
    #         save_dir = self.settings.get("save_dir", os.getcwd())
    #         fname = self.sanitize_filename(f"Отчет_{org_sel}_{start}_{end}.xlsx")
    #         save_path = os.path.join(save_dir, fname)
    #         try:
    #             from openpyxl.utils import get_column_letter
    #             with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
    #                 df.to_excel(writer, index=False, sheet_name="Report")
    #                 sheet = writer.sheets["Report"]
    #                 for idx, col in enumerate(df.columns, start=1):
    #                     width = max(df[col].astype(str).map(len).max(), len(col)) + 2
    #                     sheet.column_dimensions[get_column_letter(idx)].width = width
    #             self.show_notification(f"Отчет сохранён: {save_path}")
    #         except Exception as e:
    #             messagebox.showerror("Ошибка записи", f"Не удалось сохранить файл:\n{e}")
    #
    #     tk.Button(
    #         rpt,
    #         text="Сформировать",
    #         command=make_report,
    #         bg="#4CAF50",
    #         fg="white"
    #     ).grid(row=3, column=0, columnspan=3, pady=(10, 10), padx=padx, sticky="ew")
    #     rpt.grid_columnconfigure(1, weight=1)
    #     rpt.update_idletasks()
    #     w, h = rpt.winfo_width(), rpt.winfo_height()
    #     sw, sh = rpt.winfo_screenwidth(), rpt.winfo_screenheight()
    #     x, y = (sw - w) // 2, (sh - h) // 2
    #     rpt.geometry(f"{w}x{h}+{x}+{y}")
    #
    # def set_report_none(self):
    #     self.report_org_window = None
    #     self.report_month_window = None
    #
    # def report_by_month(self):
    #     import pandas as pd
    #     if self.report_month_window and self.report_month_window.winfo_exists():
    #         self.report_month_window.focus_force()
    #         return
    #     self.report_month_window = tk.Toplevel(self)
    #     rpt = self.report_month_window
    #     rpt.title("Отчет по дате ИДС")
    #     rpt.resizable(False, False)
    #     rpt.protocol("WM_DELETE_WINDOW", lambda: (rpt.destroy(), self.set_report_none()))
    #     padx, pady = 10, 5
    #     tk.Label(rpt, text="Период с:").grid(row=0, column=0, sticky="w", padx=padx, pady=pady)
    #     start_var = tk.StringVar()
    #     start_entry = tk.Entry(rpt, width=20, textvariable=start_var)
    #     start_entry.grid(row=0, column=1, sticky="w", padx=padx, pady=pady)
    #     tk.Button(rpt, text="📅", command=lambda: self.open_calendar(start_entry)).grid(row=0, column=2, padx=0, pady=pady)
    #     tk.Label(rpt, text="По:").grid(row=1, column=0, sticky="w", padx=padx, pady=pady)
    #     end_var = tk.StringVar()
    #     end_entry = tk.Entry(rpt, width=20, textvariable=end_var)
    #     end_entry.grid(row=1, column=1, sticky="w", padx=padx, pady=pady)
    #     tk.Button(rpt, text="📅", command=lambda: self.open_calendar(end_entry)).grid(row=1, column=2, padx=0, pady=pady)
    #     def on_start_changed(*_):
    #         s = start_var.get().strip()
    #         if self.is_valid_date(s):
    #             dt = datetime.datetime.strptime(s, "%d.%m.%Y")
    #             last = calendar.monthrange(dt.year, dt.month)[1]
    #             end_var.set(f"{last:02d}.{dt.month:02d}.{dt.year}")
    #     start_var.trace_add("write", on_start_changed)
    #     def make_report_month():
    #         start = start_var.get().strip()
    #         end = end_var.get().strip()
    #         d0 = datetime.datetime.strptime(start, "%d.%m.%Y")
    #         d1 = datetime.datetime.strptime(end, "%d.%m.%Y")
    #         rows = []
    #         for org_name, recs in self.data.items():
    #             for r in recs:
    #                 ids = r.get("ids_date", "").strip()
    #                 if not ids:
    #                     continue
    #                 try:
    #                     d_ids = datetime.datetime.strptime(ids, "%d.%m.%Y")
    #                 except ValueError:
    #                     continue
    #                 if d0 <= d_ids <= d1:
    #                     rows.append({
    #                         "Организация": org_name,
    #                         "ФИО": r["name"],
    #                         "Дата рожд.": r["birthday"],
    #                         "Пол": r["sex"],
    #                         "Подразделение": r["division"],
    #                         "Должность": r["profession"],
    #                         "Факторы": r["factors"],
    #                         "Виды работ": r["typework"],
    #                         "Дата ИДС": ids,
    #                         "Диагноз": r.get("diagnosis", "")
    #                     })
    #         if not rows:
    #             messagebox.showinfo("Пустой отчет", "Нет записей за выбранный период.")
    #             return
    #         df = pd.DataFrame(rows)
    #         save_dir = self.settings.get("save_dir", os.getcwd())
    #         fname = self.sanitize_filename(f"Отчет_по_месяцу_{start}_{end}.xlsx")
    #         save_path = os.path.join(save_dir, fname)
    #         try:
    #             from openpyxl.utils import get_column_letter
    #             with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
    #                 df.to_excel(writer, index=False, sheet_name="Report")
    #                 sheet = writer.sheets["Report"]
    #                 for idx, col in enumerate(df.columns, start=1):
    #                     width = max(df[col].astype(str).map(len).max(), len(col)) + 2
    #                     sheet.column_dimensions[get_column_letter(idx)].width = width
    #             self.show_notification(f"Отчет сохранён: {save_path}")
    #         except Exception as e:
    #             messagebox.showerror("Ошибка записи", f"Не удалось сохранить файл:\n{e}")
    #     tk.Button(
    #         rpt,
    #         text="Сформировать",
    #         command=make_report_month,
    #         bg="#4CAF50",
    #         fg="white"
    #     ).grid(row=2, column=0, columnspan=3, pady=(10, 10), padx=padx, sticky="ew")
    #     rpt.grid_columnconfigure(1, weight=1)
    #     rpt.update_idletasks()
    #     w, h = rpt.winfo_width(), rpt.winfo_height()
    #     sw, sh = rpt.winfo_screenwidth(), rpt.winfo_screenheight()
    #     x, y = (sw - w) // 2, (sh - h) // 2
    #     rpt.geometry(f"{w}x{h}+{x}+{y}")
