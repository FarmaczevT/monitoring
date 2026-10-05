import datetime
import io
import math
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import urllib.parse
from PIL import Image, ImageTk
import openpyxl
import requests
from bs4 import BeautifulSoup


def resource_path(relative_path):
    """Получает абсолютный путь к ресурсу, работает для dev и для PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ModernButton(tk.Button):
    """Стилизованная кнопка с эффектом наведения"""

    def __init__(
        self,
        master,
        bg_color="#43C639",
        hover_color="#36A02E",
        text_color="#FFFFFF",
        *args,
        **kwargs,
    ):
        super().__init__(master, *args, **kwargs)
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.text_color = text_color

        self.config(
            bg=self.bg_color,
            fg=self.text_color,
            activebackground=self.hover_color,
            activeforeground=self.text_color,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
        )

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        if self["state"] != tk.DISABLED:
            self.config(bg=self.hover_color)

    def on_leave(self, e):
        if self["state"] != tk.DISABLED:
            self.config(bg=self.bg_color)


class SpimexParserApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Organic Trade — Мониторинг СПбМТСБ")
        self.root.geometry("780x640")
        self.root.minsize(700, 580)
        self.root.configure(bg="#1E1E24")

        # --- Установка иконки окна (.ico) ---
        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception as e:
                print(f"Не удалось установить иконку: {e}")

        self.file_path = ""
        self.bulletin_data = {}  # {ticker: {date_str: price_rounded}}
        self.excel_tickers = set()

        # Список слов/заголовков для исключения
        self.excluded_names = {
            "ортоксилол", "тариф", "толуол", "нефрас с2", "керосин", 
            "ацетон", "бутанол", "изобутанол", "бгст", "мэг", "фау", 
            "этилбензол", "наименование", "продукт", "инструмент"
        }

        self.setup_ui()

        # HTTP Сессия
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "Chrome/128.0.0.0 Safari/537.36"
            )
        })

    def setup_ui(self):
        # Header Container
        header_frame = tk.Frame(self.root, bg="#2D2D37", height=100)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        # Logo Loading
        logo_path = resource_path("logo.png")
        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                img.thumbnail((220, 70), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                lbl_logo = tk.Label(
                    header_frame, image=self.logo_img, bg="#2D2D37"
                )
                lbl_logo.pack(side=tk.LEFT, padx=25, pady=15)
            except Exception as e:
                print(f"Ошибка загрузки логотипа: {e}")

        # Title Label in Header
        title_frame = tk.Frame(header_frame, bg="#2D2D37")
        title_frame.pack(side=tk.RIGHT, padx=25, pady=15)

        lbl_title = tk.Label(
            title_frame,
            text="Мониторинг СПбМТСБ",
            font=("Segoe UI", 16, "bold"),
            fg="#FFFFFF",
            bg="#2D2D37",
        )
        lbl_title.pack(anchor="e")

        lbl_subtitle = tk.Label(
            title_frame,
            text="Автоматический парсер цен инструментов",
            font=("Segoe UI", 9),
            fg="#9E9EAE",
            bg="#2D2D37",
        )
        lbl_subtitle.pack(anchor="e")

        # Main Content Area
        main_frame = tk.Frame(self.root, bg="#1E1E24")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=(15, 5))

        # File Selection Card
        card_frame = tk.Frame(
            main_frame,
            bg="#2D2D37",
            highlightbackground="#3E3E4D",
            highlightthickness=1,
        )
        card_frame.pack(fill=tk.X, pady=(0, 15))

        card_inner = tk.Frame(card_frame, bg="#2D2D37", padx=15, pady=15)
        card_inner.pack(fill=tk.X)

        self.btn_select_file = ModernButton(
            card_inner,
            text="📁 Выбрать Excel файл",
            bg_color="#3E3E4D",
            hover_color="#4E4E5E",
            command=self.select_file,
        )
        self.btn_select_file.pack(side=tk.LEFT)

        self.lbl_file = tk.Label(
            card_inner,
            text="Файл не выбран",
            fg="#888899",
            bg="#2D2D37",
            font=("Segoe UI", 9, "italic"),
            anchor="w",
        )
        self.lbl_file.pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(15, 0)
        )

        # Action Button
        self.btn_run = ModernButton(
            main_frame,
            text="🚀 Запустить мониторинг",
            bg_color="#43C639",
            hover_color="#36A02E",
            command=self.run_monitoring,
            state=tk.DISABLED,
        )
        self.btn_run.config(bg="#2D4D32", fg="#779977", cursor="arrow")
        self.btn_run.pack(fill=tk.X, pady=(0, 15))

        # Console / Log Card
        log_card = tk.Frame(
            main_frame,
            bg="#18181C",
            highlightbackground="#3E3E4D",
            highlightthickness=1,
        )
        log_card.pack(fill=tk.BOTH, expand=True)

        log_title = tk.Label(
            log_card,
            text="ЖУРНАЛ СОБЫТИЙ",
            font=("Segoe UI", 8, "bold"),
            fg="#6E6E7E",
            bg="#18181C",
            padx=10,
            pady=5,
        )
        log_title.pack(anchor="w")

        # Scrollbar + Text Log
        log_inner = tk.Frame(log_card, bg="#18181C")
        log_inner.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        scrollbar = ttk.Scrollbar(log_inner)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(
            log_inner,
            height=10,
            bg="#18181C",
            fg="#45FF75",
            insertbackground="white",
            relief=tk.FLAT,
            font=("Consolas", 9),
            yscrollcommand=scrollbar.set,
            state=tk.DISABLED,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        # Tags for log styling
        self.log_text.tag_config("error", foreground="#FF5555")
        self.log_text.tag_config("success", foreground="#43C639")
        self.log_text.tag_config("info", foreground="#888899")

        # Footer / Copyright
        footer_frame = tk.Frame(self.root, bg="#1E1E24")
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(0, 2))

        lbl_author = tk.Label(
            footer_frame,
            text="Developed by Denis Gorn. © All Rights Reserved",
            font=("Segoe UI", 6, "italic"),
            fg="#6E6E7E",
            bg="#1E1E24",
        )
        lbl_author.pack(side=tk.RIGHT, padx=15)

        self.status_bar = tk.Label(
            footer_frame,
            text="Готов к работе",
            bd=0,
            relief=tk.FLAT,
            bg="#1E1E24",
            fg="#777788",
            anchor=tk.W,
            font=("Segoe UI", 8),
            padx=15,
        )
        self.status_bar.pack(side=tk.LEFT)

    def log(self, message, tag=None):
        self.log_text.config(state=tk.NORMAL)
        if tag:
            self.log_text.insert(tk.END, message + "\n", tag)
        else:
            self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def update_status(self, text):
        self.status_bar.config(text=text)
        self.root.update_idletasks()

    def select_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if filename:
            self.file_path = filename
            basename = os.path.basename(filename)
            self.lbl_file.config(
                text=f"Выбран: {basename}", fg="#FFFFFF", font=("Segoe UI", 9)
            )

            self.btn_run.config(
                state=tk.NORMAL,
                bg="#43C639",
                fg="#FFFFFF",
                cursor="hand2",
            )
            self.log(f"Загружен файл: {filename}", "info")
            self.update_status(f"Загружен файл {basename}")

    def is_valid_ticker(self, val_str):
        """Проверяет, является ли строка валидным тикером СПбМТСБ (поддерживает дефисы)"""
        if not val_str:
            return False
        
        t_clean = str(val_str).strip()
        t_lower = t_clean.lower()

        if t_lower in self.excluded_names:
            return False

        # Добавлен дефис в класс символов: [A-Z0-9-]
        if re.match(r"^[A-Z0-9-]{7,15}$", t_clean, re.IGNORECASE):
            return True

        return False

    def load_excel_tickers(self, ws):
        """Загружает тикеры из указанного листа Excel"""
        self.excel_tickers.clear()

        for row in range(4, ws.max_row + 1):
            val = ws.cell(row=row, column=1).value
            if val and self.is_valid_ticker(val):
                self.excel_tickers.add(str(val).strip())

        self.log(
            f"Найдено действительных тикеров в Excel: {len(self.excel_tickers)}", "info"
        )

    def clean_number(self, val_str):
        if not val_str:
            return None
        cleaned = (
            str(val_str)
            .replace(" ", "")
            .replace("\xa0", "")
            .replace(",", ".")
            .strip()
        )
        if cleaned in ["-", "—", "–", "", "0"]:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None

    def round_up_thousands(self, price_num):
        """
        Делит цену на 1000 и округляет до 1 знака после запятой.
        Пример: 180011 -> 180.0 (или 180.011 с округлением)
        """
        thousands = price_num / 1000.0
        # Округляем до 1 знака после запятой (стандартное математическое округление)
        return round(thousands, 1)

    def fetch_ticker_data_for_month(self, ticker, target_month, target_year):
        """
        Скачивает страницу и ищет данные строго внутри tbody таблицы результатов торгов.
        Возвращает словарь {date_str: rounded_price}
        """
        url = f"https://spimex.com/markets/oil_products/instruments/list/detail.php?code={ticker}"
        results = {}
        try:
            res = self.session.get(url, timeout=12)
            if res.status_code != 200:
                self.log(f"[!] {ticker}: ошибка HTTP {res.status_code}", "error")
                return results

            soup = BeautifulSoup(res.text, "html.parser")
            
            # Находим таблицу по ее точному классу со скриншота
            table = soup.find("table", class_=lambda x: x and "results_table" in x)
            if not table:
                # Запасной вариант поиска через заголовок
                target_header = soup.find(
                    lambda tag: tag.name in ["h2", "h3", "div"] and 
                    "Результаты последних 10 торговых сессий" in tag.text
                )
                if target_header:
                    container = target_header.find_parent("div")
                    if container:
                        table = container.find("table")
            
            if not table:
                self.log(f"[-] {ticker}: таблица результатов торгов не найдена", "info")
                return results

            # Ищем блок <tbody>, чтобы исключить шапку (<thead>) с её строками-заголовками
            tbody = table.find("tbody")
            target_rows = tbody.find_all("tr") if tbody else table.find_all("tr")

            found_rows_count = 0

            for row in target_rows:
                cols = [c.text.strip() for c in row.find_all(["td", "th"])]
                if not cols or len(cols) < 4:
                    continue
                
                # Проверяем первую колонку на наличие даты формата ДД.ММ.ГГГГ (например, 01.10.2026)
                date_match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", cols[0])
                if date_match:
                    found_rows_count += 1
                    day_str, month_str, year_str = date_match.groups()
                    
                    if int(month_str) == target_month and int(year_str) == target_year:
                        trade_date = f"{day_str}.{month_str}.{year_str}"
                        
                        # 4-я колонка (индекс 3) — «Последняя» цена со скриншота
                        last_price_raw = cols[3]
                        price_num = self.clean_number(last_price_raw)
                        
                        if price_num:
                            rounded_price = self.round_up_thousands(price_num)
                            results[trade_date] = rounded_price

            if found_rows_count > 0:
                self.log(f"[i] {ticker}: найдено строк в tbody: {found_rows_count}, подошло за месяц: {len(results)}", "info")
            else:
                self.log(f"[-] {ticker}: в таблице не найдены строки с датами торгов", "info")

            return results

        except Exception as e:
            self.log(f"[!] Ошибка парсинга {ticker}: {e}", "error")
            return results

    def fetch_all_tickers_data(self, target_month, target_year):
        """Опрашивает страницы всех тикеров за указанный месяц"""
        self.bulletin_data.clear()
        total = len(self.excel_tickers)
        
        for idx, ticker in enumerate(sorted(self.excel_tickers), 1):
            self.update_status(f"Парсинг {idx}/{total}: {ticker}...")
            month_data = self.fetch_ticker_data_for_month(ticker, target_month, target_year)
            
            if month_data:
                self.bulletin_data[ticker] = month_data
                for d_str, price in month_data.items():
                    self.log(f"[✓] {ticker} | Дата: {d_str} | Последняя: {price}", "success")
            else:
                self.log(f"[-] {ticker}: данные за {target_month:02d}.{target_year} не найдены", "info")

    def get_or_create_sheet_for_current_month(self, wb):
        """Определяет имя листа для текущего месяца (например T10) и делает его активным"""
        now = datetime.datetime.now()
        sheet_name = f"Т{now.month}"
        
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(title=sheet_name)
            self.log(f"[i] Создан новый лист: {sheet_name}", "info")
            
        wb.active = ws
        self.log(f"Активный лист в Excel: {sheet_name}", "info")
        return ws, now.month, now.year

    def run_monitoring(self):
        self.btn_run.config(
            state=tk.DISABLED, bg="#2D4D32", fg="#779977", cursor="arrow"
        )
        self.log("\n================ ЗАПУСК МОНИТОРИНГА ================")

        try:
            wb = openpyxl.load_workbook(self.file_path)
            ws, target_month, target_year = self.get_or_create_sheet_for_current_month(wb)

            self.load_excel_tickers(ws)

            if not self.excel_tickers:
                self.log("[!] Ошибка: тикеры в листе Excel не найдены.", "error")
                self.update_status("Ошибка выполнения")
                return

            self.fetch_all_tickers_data(target_month, target_year)

            self.update_status("Запись данных в Excel...")

            # Считываем имеющиеся даты во 2-й строке Excel (с поддержкой объектов дат и текста)
            date_columns = {}
            for col in range(3, ws.max_column + 1):
                val = ws.cell(row=2, column=col).value
                if val is not None:
                    if isinstance(val, (datetime.datetime, datetime.date)):
                        date_str = val.strftime("%d.%m.%Y")
                    else:
                        d_raw = str(val).strip()
                        d_parts = re.findall(r"\d+", d_raw)
                        if len(d_parts) >= 2:
                            day = int(d_parts[0])
                            month = int(d_parts[1])
                            year = int(d_parts[2]) if len(d_parts) > 2 else target_year
                            date_str = f"{day:02d}.{month:02d}.{year}"
                        else:
                            date_str = d_raw
                    date_columns[date_str] = col

            added_count = 0
            skipped_count = 0

            for row in range(4, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=1).value

                if cell_val and self.is_valid_ticker(cell_val):
                    ticker = str(cell_val).strip()
                    tariff_row = row - 1

                    if ticker in self.bulletin_data:
                        ticker_month_data = self.bulletin_data[ticker]

                        for trade_date_str, price in ticker_month_data.items():
                            col_idx = date_columns.get(trade_date_str)
                            if not col_idx:
                                alt_date = f"{int(trade_date_str[:2])}.{int(trade_date_str[3:5])}.{trade_date_str[6:]}"
                                col_idx = date_columns.get(alt_date)

                            if col_idx:
                                # ПРОВЕРКА: Если ячейка с тарифом пустая, записываем! Если уже есть — НЕ перезаписываем.
                                current_tariff_val = ws.cell(row=tariff_row, column=col_idx).value
                                
                                if current_tariff_val is None or str(current_tariff_val).strip() == "":
                                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                                    ws.cell(row=tariff_row, column=col_idx).value = price
                                    
                                    formula = f"={col_letter}{tariff_row}+$B${tariff_row}"
                                    ws.cell(row=row, column=col_idx).value = formula

                                    added_count += 1
                                else:
                                    skipped_count += 1
                            else:
                                self.log(
                                    f"[!] Дата {trade_date_str} для {ticker} отсутствует во 2-й строке Excel.",
                                    "error",
                                )

            output_file = self.file_path.replace(".xlsx", "_updated.xlsx")
            wb.save(output_file)

            self.log(
                f"\nУСПЕШНО ЗАВЕРШЕНО!\nЗаписано новых ячеек: {added_count}\nПропущено (уже заполнены): {skipped_count}\nРезультат сохранен: {output_file}",
                "success",
            )
            self.update_status(f"Завершено. Добавлено: {added_count}, Пропущено: {skipped_count}")
            messagebox.showinfo(
                "Успех",
                f"Мониторинг успешно завершен!\n\nЗаписано новых цен: {added_count}\nПропущено (уже есть данные): {skipped_count}\n\nФайл сохранен рядом с исходным.",
            )

        except PermissionError:
            output_file = self.file_path.replace(".xlsx", "_updated.xlsx")
            self.log(
                f"[!] Файл {output_file} открыт в Excel. Закройте его и повторите попытку.",
                "error",
            )
            messagebox.showerror(
                "Ошибка доступа",
                f"Не удалось сохранить файл:\n{output_file}\n\nПожалуйста, закройте его в Excel и запустите снова.",
            )
        except Exception as e:
            self.log(f"Ошибка: {e}", "error")
            messagebox.showerror("Ошибка", str(e))
            self.update_status("Произошла ошибка")
        finally:
            if self.file_path:
                self.btn_run.config(
                    state=tk.NORMAL,
                    bg="#43C639",
                    fg="#FFFFFF",
                    cursor="hand2",
                )


if __name__ == "__main__":
    root = tk.Tk()
    app = SpimexParserApp(root)
    root.mainloop()