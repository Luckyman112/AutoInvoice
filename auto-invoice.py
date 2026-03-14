import customtkinter as ctk
from tkinter import filedialog
import openpyxl
import os

# --- НАСТРОЙКИ ОКНА ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.geometry("600x500")
app.title("Генератор Счетов")

# --- ПЕРЕМЕННЫЕ ДЛЯ ПУТЕЙ ---
bi_report_path = ""
template_path = ""

# --- ЛОГИКА КНОПОК ИНТЕРФЕЙСА ---
def load_bi_report():
    global bi_report_path
    bi_report_path = filedialog.askopenfilename(
        title="Выберите BI-отчет",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if bi_report_path:
        log_text.insert("end", f"[OK] BI-отчет загружен:\n{os.path.basename(bi_report_path)}\n\n")
        log_text.see("end")

def load_template():
    global template_path
    template_path = filedialog.askopenfilename(
        title="Выберите Шаблон",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if template_path:
        log_text.insert("end", f"[OK] Шаблон загружен:\n{os.path.basename(template_path)}\n\n")
        log_text.see("end")

def generate_invoices():
    if not bi_report_path or not template_path:
        log_text.insert("end", "[ОШИБКА] Сначала загрузите BI-отчет и Шаблон!\n\n")
        log_text.see("end")
        return
    
    save_folder = filedialog.askdirectory(title="Выберите папку для сохранения счетов")
    if not save_folder:
        log_text.insert("end", "[ОТМЕНА] Папка не выбрана.\n\n")
        return
    
    log_text.insert("end", "[*] Начинаем генерацию...\n")
    log_text.see("end")
    app.update() # Обновляем окно, чтобы текст появился сразу

    try:
        # ==========================================
        # 1: ЧИТАЕМ БИ-ОТЧЕТ И СОБИРАЕМ БАНКИ
        # ==========================================
        wb_bi = openpyxl.load_workbook(bi_report_path, data_only=True)
        sheet_bi = wb_bi.active
        
        # Собираем все названия банков из колонки A в список
        all_banks = []
        for row in sheet_bi.iter_rows(min_row=2, values_only=True):
            if row[0]: # Если колонка A не пустая
                all_banks.append(str(row[0]).strip())
        
        # ==========================================
        # 2: ОПРЕДЕЛЯЕМ БАНК ПО НАЗВАНИЮ ШАБЛОНА
        # ==========================================
        template_filename = os.path.basename(template_path)
        bank_name_to_search = None
        
        # Собираем ВСЕ банки, которые встречаются в названии файла
        possible_banks = []
        for bank in all_banks:
            if bank.lower() in template_filename.lower():
                possible_banks.append(bank)
                
        # Если нашлось несколько совпадений (например, ['DECTA', 'Teslapay'])
        if len(possible_banks) > 1:
            # Удаляем 'DECTA' из списка кандидатов, так как это компания-отправитель
            possible_banks = [b for b in possible_banks if b.lower() != 'decta']
            
        # На всякий случай сортируем по длине (чтобы брать самое длинное и точное название)
        possible_banks.sort(key=len, reverse=True)
                
        if not possible_banks:
            log_text.insert("end", f"[ОШИБКА] Не удалось найти название клиента в имени файла: {template_filename}\n\n")
            log_text.see("end")
            return
            
        # Берем правильного клиента
        bank_name_to_search = possible_banks[0]
            
        log_text.insert("end", f"[*] Распознан банк-клиент: {bank_name_to_search}\n")
        app.update()
        
        # ==========================================
        # 3: ИЩЕМ ДАННЫЕ В БИ-ОТЧЕТЕ
        # ==========================================
        actual_number_of_transactions_issuing = 0
        count_cards = 0
        card_count_in_period = 0

        for row in sheet_bi.iter_rows(values_only=True):
            col_a = row[0]
            col_f = row[5]
            
            if col_a and str(col_a).strip().lower() == bank_name_to_search.lower():
                actual_number_of_transactions_issuing = row[1] if row[1] is not None else 0
                
            if col_f and str(col_f).strip().lower() == bank_name_to_search.lower():
                count_cards = row[6] if row[6] is not None else 0
                card_count_in_period = row[7] if row[7] is not None else 0

        # ==========================================
        # 4: ЗАПИСЫВАЕМ В ШАБЛОН
        # ==========================================
        wb_template = openpyxl.load_workbook(template_path)
        target_sheet_name = None
        for name in wb_template.sheetnames:
            if "annex" in name.lower():
                target_sheet_name = name
                break 

        if target_sheet_name:
            sheet_template = wb_template[target_sheet_name]
            sheet_template["I12"] = actual_number_of_transactions_issuing
            sheet_template["I11"] = count_cards
            sheet_template["I10"] = card_count_in_period
            
            # Сохраняем с понятным именем
            new_filename = f"READY_{template_filename}"
            new_file_path = os.path.join(save_folder, new_filename)
            wb_template.save(new_file_path)
            
            log_text.insert("end", f"[УСПЕХ] Готово! Сохранено как: {new_filename}\n\n")
        else:
            log_text.insert("end", "[ОШИБКА] Лист 'annex' не найден в шаблоне!\n\n")

    except Exception as e:
        log_text.insert("end", f"[КРИТИЧЕСКАЯ ОШИБКА] Что-то пошло не так: {e}\n\n")
    
    log_text.see("end")

# --- ВИЗУАЛ (КНОПКИ И ТЕКСТ) ---
label_title = ctk.CTkLabel(app, text="Генератор Счетов по BI-отчету", font=("Arial", 20, "bold"))
label_title.pack(pady=(20, 10))

btn_bi = ctk.CTkButton(app, text="1. Загрузить BI-отчет", command=load_bi_report, width=250)
btn_bi.pack(pady=10)

btn_tpl = ctk.CTkButton(app, text="2. Загрузить Шаблон счета", command=load_template, width=250)
btn_tpl.pack(pady=10)

btn_gen = ctk.CTkButton(app, text="3. Сгенерировать счет", command=generate_invoices, fg_color="#28a745", hover_color="#218838", width=250)
btn_gen.pack(pady=(20, 20))

log_text = ctk.CTkTextbox(app, width=550, height=180)
log_text.pack(pady=10)
log_text.insert("end", "Добро пожаловать! Загрузите файлы для начала работы.\n\n")

app.mainloop()