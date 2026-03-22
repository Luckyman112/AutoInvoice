import customtkinter as ctk
from tkinter import filedialog
import openpyxl
import os

# --- НАСТРОЙКИ ОКНА ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.geometry("650x600")
app.title("Auto-invoice v1.0")

# --- ПЕРЕМЕННЫЕ ДЛЯ ПУТЕЙ ---
bi_files = []      
template_files = [] 

# --- ЛОГИКА КНОПОК ИНТЕРФЕЙСА ---
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

def load_templates():
    global template_files
    template_files = filedialog.askopenfilenames(
        title="Выберите ВСЕ Шаблоны счетов",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if template_files:
        log_text.insert("end", f"[OK] Выбрано шаблонов: {len(template_files)}\n\n")
        log_text.see("end")

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

        # ==============================================================
        # 1. ЧИТАЕМ BI-ОТЧЕТЫ (С точными координатами колонок)
        # ==============================================================
        for path in bi_files:
            filename = os.path.basename(path).upper()
            wb = openpyxl.load_workbook(path, data_only=True)
            sheet = wb.active

            if "ISS" in filename:
                log_text.insert("end", f"[*] Читаем базу Ишшуинга: {os.path.basename(path)}\n")
                # Ишшуинг: стартуем с 6 строки (A6 и F6)
                for row in sheet.iter_rows(min_row=6, values_only=True):
                    # Колонка A (индекс 0)
                    if len(row) > 0 and row[0]: 
                        bank = str(row[0]).strip()
                        if bank not in iss_data: iss_data[bank] = {"txn": 0, "cards_total": 0, "cards_period": 0}
                        iss_data[bank]["txn"] = row[1] if len(row) > 1 and row[1] else 0
                        
                    # Колонка F (индекс 5)
                    if len(row) > 7 and row[5]: 
                        bank = str(row[5]).strip()
                        if bank not in iss_data: iss_data[bank] = {"txn": 0, "cards_total": 0, "cards_period": 0}
                        iss_data[bank]["cards_total"] = row[6] if len(row) > 6 and row[6] else 0
                        iss_data[bank]["cards_period"] = row[7] if len(row) > 7 and row[7] else 0

            elif "ACQ" in filename:
                log_text.insert("end", f"[*] Читаем базу Эквайринга: {os.path.basename(path)}\n")
                # Эквайринг: стартуем с 5 строки (A5 и F5)
                for row in sheet.iter_rows(min_row=5, values_only=True):
                    # Колонка A (индекс 0)
                    if len(row) > 0 and row[0]:
                        bank = str(row[0]).strip()
                        if bank not in acq_data: acq_data[bank] = {"txn": 0, "chargebacks": 0}
                        acq_data[bank]["txn"] = row[1] if len(row) > 1 and row[1] else 0
                        
                    # Колонка F (индекс 5)
                    if len(row) > 6 and row[5]:
                        bank = str(row[5]).strip()
                        if bank not in acq_data: acq_data[bank] = {"txn": 0, "chargebacks": 0}
                        acq_data[bank]["chargebacks"] = row[6] if len(row) > 6 and row[6] else 0

            elif "AUTH" in filename:
                log_text.insert("end", f"[*] Читаем базу Авторизаций: {os.path.basename(path)}\n")
                # Авторизации: стартуем с 5 строки (A5 и H5)
                for row in sheet.iter_rows(min_row=5, values_only=True):
                    # Колонка A (индекс 0)
                    if len(row) > 0 and row[0]:
                        bank = str(row[0]).strip()
                        if bank not in auth_data: auth_data[bank] = {"auth_acq": 0, "auth_iss": 0}
                        auth_data[bank]["auth_acq"] = row[2] if len(row) > 2 and row[2] else 0
                        
                    # Колонка H (индекс 7)
                    if len(row) > 8 and row[7]:
                        bank = str(row[7]).strip()
                        if bank not in auth_data: auth_data[bank] = {"auth_acq": 0, "auth_iss": 0}
                        auth_data[bank]["auth_iss"] = row[8] if len(row) > 8 and row[8] else 0
        
        all_banks = list(set(iss_data.keys()).union(set(acq_data.keys()), set(auth_data.keys())))
        app.update()

        # ==============================================================
        # 1.5 ВАЛИДАЦИЯ: ПРОВЕРЯЕМ НАЛИЧИЕ ВСЕХ НУЖНЫХ ШАБЛОНОВ
        # ==============================================================
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

        # ==============================================================
        # 2. ПЕРЕБИРАЕМ И ЗАПОЛНЯЕМ ВСЕ ШАБЛОНЫ
        # ==============================================================
        log_text.insert("end", "[*] НАЧИНАЕМ ЗАПОЛНЕНИЕ ШАБЛОНОВ:\n")
        
        for tpl_path in template_files:
            tpl_filename = os.path.basename(tpl_path)
            
            # --- ИЩЕМ БАНК ---
            possible_banks = [b for b in all_banks if b.lower() in tpl_filename.lower()]
            if len(possible_banks) > 1:
                possible_banks = [b for b in possible_banks if b.lower() != 'decta']
            possible_banks.sort(key=len, reverse=True)
            
            if not possible_banks:
                log_text.insert("end", f"  [ПРОПУСК] В '{tpl_filename}' не найден банк!\n")
                continue
                
            bank_name = possible_banks[0]
            
            # --- ОТКРЫВАЕМ ШАБЛОН И ИЩЕМ ЛИСТ ANNEX ---
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

            # --- ОПРЕДЕЛЯЕМ ТИП И ЗАПОЛНЯЕМ ---
            tpl_type = "UNKNOWN"
            
            if "ISS" in tpl_filename.upper():
                h17_val = sheet_tpl["H17"].value
                h17_text = str(h17_val).strip() if h17_val else ""
                
                if "Good Faith" in h17_text:
                    tpl_type = "ISS_2"
                else:
                    tpl_type = "ISS_1"
                    
                b_iss = iss_data.get(bank_name, {"txn": 0, "cards_total": 0, "cards_period": 0})
                b_auth = auth_data.get(bank_name, {"auth_iss": 0})

                if tpl_type == "ISS_1":
                    log_text.insert("end", f"  [+] ISS-1 (H17 пусто): {bank_name}\n")
                    sheet_tpl["I14"] = b_iss["txn"]
                    sheet_tpl["I11"] = b_iss["cards_total"]
                    sheet_tpl["I10"] = b_iss["cards_period"]
                    sheet_tpl["I12"] = b_auth["auth_iss"]

                elif tpl_type == "ISS_2":
                    log_text.insert("end", f"  [+] ISS-2 (H17 Good Faith): {bank_name}\n")
                    sheet_tpl["I14"] = b_iss["txn"]
                    sheet_tpl["I11"] = b_iss["cards_total"]
                    sheet_tpl["I10"] = b_iss["cards_period"]
                    sheet_tpl["I12"] = b_auth["auth_iss"]

            elif "ACQ" in tpl_filename.upper():
                h13_val = sheet_tpl["H13"].value
                h13_text = str(h13_val).strip() if h13_val else ""
                
                if "RDR" in h13_text:
                    tpl_type = "ACQ_2"
                else:
                    tpl_type = "ACQ_1"
                    
                b_acq = acq_data.get(bank_name, {"txn": 0, "chargebacks": 0})
                b_auth = auth_data.get(bank_name, {"auth_acq": 0})

                if tpl_type == "ACQ_1":
                    log_text.insert("end", f"  [+] ACQ-1 (H13 пусто): {bank_name}\n")
                    sheet_tpl["I11"] = b_acq["txn"]
                    sheet_tpl["I12"] = b_acq["chargebacks"]
                    sheet_tpl["I10"] = b_auth["auth_acq"]

                elif tpl_type == "ACQ_2":
                    log_text.insert("end", f"  [+] ACQ-2 (H13 RDR): {bank_name}\n")
                    sheet_tpl["I11"] = b_acq["txn"]
                    sheet_tpl["I12"] = b_acq["chargebacks"]
                    sheet_tpl["I10"] = b_auth["auth_acq"]

            # Сохраняем готовый файл
            new_filename = f"READY_{tpl_filename}"
            wb_tpl.save(os.path.join(save_folder, new_filename))
            log_text.insert("end", f"      -> Сохранен как {new_filename}\n")
            app.update()

        log_text.insert("end", "\n[УСПЕХ] Все доступные счета сгенерированы!\n\n")
        log_text.see("end")

    except Exception as e:
        log_text.insert("end", f"\n[КРИТИЧЕСКАЯ ОШИБКА] Что-то пошло не так: {e}\n\n")
        log_text.see("end")

# --- ВИЗУАЛ (КНОПКИ И ТЕКСТ) ---
label_title = ctk.CTkLabel(app, text="Массовый Генератор Счетов", font=("Arial", 20, "bold"))
label_title.pack(pady=(20, 10))

btn_bi = ctk.CTkButton(app, text="1. Выбрать 3 BI-отчета", command=load_bi_reports, width=300)
btn_bi.pack(pady=10)

btn_tpl = ctk.CTkButton(app, text="2. Выбрать ВСЕ Шаблоны", command=load_templates, width=300)
btn_tpl.pack(pady=10)

btn_gen = ctk.CTkButton(app, text="3. СГЕНЕРИРОВАТЬ СЧЕТА", command=generate_invoices, fg_color="#28a745", hover_color="#218838", width=300)
btn_gen.pack(pady=(20, 20))

log_text = ctk.CTkTextbox(app, width=600, height=220)
log_text.pack(pady=10)
log_text.insert("end", "Программа готова! Вы можете выделить сразу несколько файлов в окне загрузки.\n\n")

app.mainloop()