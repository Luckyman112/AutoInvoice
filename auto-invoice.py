import customtkinter as ctk
from tkinter import filedialog
import openpyxl
import os
import re
import math

# НАСТРОЙКИ ОКНА
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.geometry("650x650")
app.title("Auto-invoice v1.2")

# ПЕРЕМЕННЫЕ ДЛЯ ПУТЕЙ
bi_files = []      
template_files = [] 
token_files = []

# ЛОГИКА КНОПОК ИНТЕРФЕЙСА
def load_bi_reports():
    global bi_files
    bi_files = filedialog.askopenfilenames(
        title="Выберите все 3 BI-отчета (ISS, ACQ, AUTH)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if bi_files:
        log_text.insert("end", f"[OK] Выбрано BI-отчетов: {len(bi_files)}\n")
        for f in bi_files:
            log_text.insert("end", f"  - {os.path.basename(f)}\n")
        log_text.insert("end", "\n")
        log_text.see("end")

def load_token_reports():
    global token_files
    token_files = filedialog.askopenfilenames(
        title="Выберите отчеты по токенам (.xlsx)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if token_files:
        log_text.insert("end", f"[OK] Выбрано отчетов по токенам: {len(token_files)}\n")
        for f in token_files:
            log_text.insert("end", f"  - {os.path.basename(f)}\n")
        log_text.insert("end", "\n")
        log_text.see("end")

def load_templates():
    global template_files
    template_files = filedialog.askopenfilenames(
        title="Выберите ВСЕ Шаблоны счетов",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if template_files:
        log_text.insert("end", f"[OK] Выбрано шаблонов: {len(template_files)}\n\n")
        log_text.see("end")

# Функция для расчета токенов (1-10=1, 11-20=2 и т.д.)
def calculate_billed_tokens(raw_tokens):
    try:
        val = int(float(raw_tokens))
        if val <= 0: return 0
        return math.ceil(val / 10)
    except (ValueError, TypeError):
        return 0

def generate_invoices():
    if not bi_files or not template_files:
        log_text.insert("end", "[ОШИБКА] Сначала загрузите BI-отчеты и Шаблоны!\n\n")
        log_text.see("end")
        return
    
    save_folder = filedialog.askdirectory(title="Выберите папку для сохранения счетов")
    if not save_folder:
        log_text.insert("end", "[ОТМЕНА] Папка не выбрана.\n\n")
        return
    
    log_text.insert("end", "\n====================================\n")
    log_text.insert("end", "[*] НАЧИНАЕМ МАГИЮ...\n")
    app.update()

    try:
        iss_data = {}
        acq_data = {}
        auth_data = {}
        total_tokens = 0

        # ЧИТАЕМ ОТЧЕТЫ ПО ТОКЕНАМ
        for path in token_files:
            filename = os.path.basename(path).upper()
            log_text.insert("end", f"[*] Считаем токены из файла: {filename}\n")
            
            month_match = re.search(r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*\_?\s*(\d{4})', filename)
            target_date_str = f"{month_match.group(1)}{month_match.group(2)}" if month_match else None
            
            wb = openpyxl.load_workbook(path, data_only=True)
            sheet = wb.active

            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) > 11:
                    col_g_date = str(row[6]).strip().upper().replace(" ", "") if row[6] else ""
                    if target_date_str and target_date_str not in col_g_date:
                        continue 

                    col_h_raw_tokens = row[7] 
                    col_k = str(row[10]).strip().upper() if row[10] else "" 
                    col_l = str(row[11]).strip().upper() if row[11] else "" 
                    
                    if col_l == "NEW":
                        total_tokens += calculate_billed_tokens(col_h_raw_tokens)
                    elif col_l == "0" and col_k == "CHANGES!":
                        total_tokens += calculate_billed_tokens(col_h_raw_tokens)

        if total_tokens > 0:
            log_text.insert("end", f"  -> Общее количество токенов насчитано: {total_tokens}\n\n")
        app.update()

        # ЧИТАЕМ BI-ОТЧЕТЫ
        for path in bi_files:
            filename = os.path.basename(path).upper()
            wb = openpyxl.load_workbook(path, data_only=True)
            sheet = wb.active

            if "ISS" in filename:
                log_text.insert("end", f"[*] Читаем базу Ишшуинга: {os.path.basename(path)}\n")
                for row in sheet.iter_rows(min_row=6, values_only=True):
                    if len(row) > 0 and row[0]: 
                        bank = str(row[0]).strip()
                        if bank not in iss_data: iss_data[bank] = {"txn": 0, "cards_total": 0, "cards_period": 0}
                        iss_data[bank]["txn"] = row[1] if len(row) > 1 and row[1] else 0
                        
                    if len(row) > 7 and row[5]: 
                        bank = str(row[5]).strip()
                        if bank not in iss_data: iss_data[bank] = {"txn": 0, "cards_total": 0, "cards_period": 0}
                        iss_data[bank]["cards_total"] = row[6] if len(row) > 6 and row[6] else 0
                        iss_data[bank]["cards_period"] = row[7] if len(row) > 7 and row[7] else 0

            elif "ACQ" in filename:
                log_text.insert("end", f"[*] Читаем базу Эквайринга: {os.path.basename(path)}\n")
                for row in sheet.iter_rows(min_row=5, values_only=True):
                    if len(row) > 0 and row[0]:
                        bank = str(row[0]).strip()
                        if bank not in acq_data: acq_data[bank] = {"txn": 0, "chargebacks": 0}
                        acq_data[bank]["txn"] = row[1] if len(row) > 1 and row[1] else 0
                        
                    if len(row) > 6 and row[5]:
                        bank = str(row[5]).strip()
                        if bank not in acq_data: acq_data[bank] = {"txn": 0, "chargebacks": 0}
                        acq_data[bank]["chargebacks"] = row[6] if len(row) > 6 and row[6] else 0

            elif "AUTH" in filename:
                log_text.insert("end", f"[*] Читаем базу Авторизаций: {os.path.basename(path)}\n")
                for row in sheet.iter_rows(min_row=5, values_only=True):
                    if len(row) > 0 and row[0]:
                        bank = str(row[0]).strip()
                        if bank not in auth_data: auth_data[bank] = {"auth_acq": 0, "auth_iss": 0}
                        auth_data[bank]["auth_acq"] = row[2] if len(row) > 2 and row[2] else 0
                        
                    if len(row) > 8 and row[7]:
                        bank = str(row[7]).strip()
                        if bank not in auth_data: auth_data[bank] = {"auth_acq": 0, "auth_iss": 0}
                        auth_data[bank]["auth_iss"] = row[8] if len(row) > 8 and row[8] else 0
        
        all_banks = list(set(iss_data.keys()).union(set(acq_data.keys()), set(auth_data.keys())))
        app.update()

        # ВАЛИДАЦИЯ ШАБЛОНОВ
        log_text.insert("end", "\n[*] Проверяем наличие всех шаблонов...\n")
        
        found_iss_templates = set()
        found_acq_templates = set()
        
        for tpl_path in template_files:
            tpl_filename = os.path.basename(tpl_path)
            possible_banks = [b for b in all_banks if b.lower() in tpl_filename.lower()]
            if len(possible_banks) > 1:
                possible_banks = [b for b in possible_banks if b.lower() != 'decta']
            possible_banks.sort(key=len, reverse=True)
            
            if possible_banks:
                bank_name = possible_banks[0]
                if "ISS" in tpl_filename.upper():
                    found_iss_templates.add(bank_name)
                elif "ACQ" in tpl_filename.upper():
                    found_acq_templates.add(bank_name)

        missing_warnings = []
        for bank in iss_data.keys():
            if bank not in found_iss_templates:
                missing_warnings.append(f"[{bank}] — забыт шаблон ISS")
        for bank in acq_data.keys():
            if bank not in found_acq_templates:
                missing_warnings.append(f"[{bank}] — забыт шаблон ACQ")
                
        if missing_warnings:
            log_text.insert("end", "⚠️ [ВНИМАНИЕ] Для некоторых банков не загружены шаблоны:\n")
            for warning in missing_warnings:
                log_text.insert("end", f"   ❌ {warning}\n")
            log_text.insert("end", "[*] Пропускаем их и генерируем счета для остальных...\n\n")
        else:
            log_text.insert("end", "[OK] Все необходимые шаблоны найдены!\n\n")
            
        log_text.see("end")
        app.update()

        # ПЕРЕБИРАЕМ И ЗАПОЛНЯЕМ ВСЕ ШАБЛОНЫ
        log_text.insert("end", "[*] НАЧИНАЕМ ЗАПОЛНЕНИЕ ШАБЛОНОВ:\n")
        
        for tpl_path in template_files:
            tpl_filename = os.path.basename(tpl_path)
            
            possible_banks = [b for b in all_banks if b.lower() in tpl_filename.lower()]
            if len(possible_banks) > 1:
                possible_banks = [b for b in possible_banks if b.lower() != 'decta']
            possible_banks.sort(key=len, reverse=True)
            
            if not possible_banks:
                log_text.insert("end", f"  [ПРОПУСК] В '{tpl_filename}' не найден банк!\n")
                continue
                
            bank_name = possible_banks[0]
            
            wb_tpl = openpyxl.load_workbook(tpl_path)
            target_sheet_name = None
            for name in wb_tpl.sheetnames:
                if "annex" in name.lower():
                    target_sheet_name = name
                    break
            
            if not target_sheet_name:
                log_text.insert("end", f"  [ОШИБКА] В '{tpl_filename}' нет листа Annex!\n")
                continue
                
            sheet_tpl = wb_tpl[target_sheet_name]
            tpl_type = "UNKNOWN"
            
            if "ISS" in tpl_filename.upper():
                h17_val = sheet_tpl["H17"].value
                h17_text = str(h17_val).strip() if h17_val else ""
                tpl_type = "ISS_2" if "Good Faith" in h17_text else "ISS_1"
                    
                b_iss = iss_data.get(bank_name, {"txn": 0, "cards_total": 0, "cards_period": 0})
                b_auth = auth_data.get(bank_name, {"auth_iss": 0})

                if tpl_type == "ISS_1" or tpl_type == "ISS_2":
                    log_text.insert("end", f"  [+] ISS: {bank_name}\n")
                    sheet_tpl["I14"] = b_iss["txn"]
                    sheet_tpl["I11"] = b_iss["cards_total"]
                    sheet_tpl["I10"] = b_iss["cards_period"]
                    sheet_tpl["I12"] = b_auth["auth_iss"]
                    
                    # ПРОВЕРКА И ЗАПИСЬ ТОКЕНОВ
                    h15_val = sheet_tpl["H15"].value
                    h15_text = str(h15_val).strip().lower() if h15_val else ""
                    
                    if "token" in h15_text and total_tokens > 0:
                        sheet_tpl["I15"] = total_tokens
                        log_text.insert("end", f"      -> Токены ({total_tokens}) записаны в ячейку I15\n")

            elif "ACQ" in tpl_filename.upper():
                h13_val = sheet_tpl["H13"].value
                h13_text = str(h13_val).strip() if h13_val else ""
                tpl_type = "ACQ_2" if "RDR" in h13_text else "ACQ_1"
                    
                b_acq = acq_data.get(bank_name, {"txn": 0, "chargebacks": 0})
                b_auth = auth_data.get(bank_name, {"auth_acq": 0})

                if tpl_type == "ACQ_1" or tpl_type == "ACQ_2":
                    log_text.insert("end", f"  [+] ACQ: {bank_name}\n")
                    sheet_tpl["I11"] = b_acq["txn"]
                    sheet_tpl["I12"] = b_acq["chargebacks"]
                    sheet_tpl["I10"] = b_auth["auth_acq"]

            new_filename = f"READY_{tpl_filename}"
            wb_tpl.save(os.path.join(save_folder, new_filename))
            log_text.insert("end", f"      -> Сохранен как {new_filename}\n")
            app.update()

        log_text.insert("end", "\n[УСПЕХ] Все доступные счета сгенерированы!\n\n")
        log_text.see("end")

    except Exception as e:
        log_text.insert("end", f"\n[Critical Error] Please reboot the program, if error still going contact to sergejs.kravecs@decta.com : {e}\n\n")
        log_text.see("end")

# ВИЗУАЛ 
label_title = ctk.CTkLabel(app, text="Mass Invoice Generator", font=("Arial", 20, "bold"))
label_title.pack(pady=(10, 10))

btn_bi = ctk.CTkButton(app, text="1. Select 3 BI-reports(AUTH,ISS,ACQ)", command=load_bi_reports, width=300)
btn_bi.pack(pady=5)

btn_tokens = ctk.CTkButton(app, text="2. Select Token reports (.xlsx)", command=load_token_reports, width=300)
btn_tokens.pack(pady=5)

btn_tpl = ctk.CTkButton(app, text="3. Select ALL Templates", command=load_templates, width=300)
btn_tpl.pack(pady=5)

btn_gen = ctk.CTkButton(app, text="4. GENERATE INVOICES", command=generate_invoices, fg_color="#28a745", hover_color="#218838", width=300)
btn_gen.pack(pady=(15, 10))

log_text = ctk.CTkTextbox(app, width=600, height=220)
log_text.pack(pady=10)
log_text.insert("end", "Программа готова! Загрузите нужные файлы.\n\n")

app.mainloop()