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
import pdfplumber
import requests


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
        self.bulletin_data = {}  # {ticker: price}
        self.bulletin_date_str = None
        self.excel_tickers = set()

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
            text="Автоматический парсер PDF бюллетеней",
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

        # Tag for errors/warnings in log
        self.log_text.tag_config("error", foreground="#FF5555")
        self.log_text.tag_config("success", foreground="#43C639")
        self.log_text.tag_config("info", foreground="#888899")

        # --- Блок авторских прав (Копирайт) ---
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

        # Status Bar
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

    def load_excel_tickers(self):
        self.excel_tickers.clear()
        wb = openpyxl.load_workbook(self.file_path, data_only=True)
        ws = wb.active

        for row in range(4, ws.max_row + 1):
            val = ws.cell(row=row, column=1).value
            if val:
                t_str = str(val).strip()
                if t_str != "Тариф" and len(t_str) >= 6:
                    self.excel_tickers.add(t_str)

        self.log(
            f"Загружено тикеров из Excel: {len(self.excel_tickers)}", "info"
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
        """Переводит сумму в тысячи и округляет до 1 знака в большую сторону"""
        thousands = price_num / 1000.0
        rounded = math.ceil(thousands * 10) / 10.0
        if rounded.is_integer():
            return int(rounded)
        return rounded

    def parse_pdf_bytes(self, pdf_bytes):
        self.bulletin_data.clear()

        if not self.excel_tickers:
            self.log("[!] Ошибка: список тикеров из Excel пуст.", "error")
            return

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                if not words:
                    continue

                rows_dict = {}
                for w in words:
                    top_key = round(w["top"] / 3.0) * 3
                    if top_key not in rows_dict:
                        rows_dict[top_key] = []
                    rows_dict[top_key].append(w)

                sorted_tops = sorted(rows_dict.keys())

                for top_key in sorted_tops:
                    line_words = sorted(
                        rows_dict[top_key], key=lambda x: x["x0"]
                    )
                    line_text = " ".join([w["text"] for w in line_words])

                    found_ticker = None
                    for ticker in self.excel_tickers:
                        if ticker in line_text:
                            found_ticker = ticker
                            break

                    if not found_ticker:
                        continue

                    ticker_idx = line_text.find(found_ticker)
                    after_ticker = line_text[ticker_idx + len(found_ticker) :]

                    tokens = re.findall(
                        r"(?:(?:\d{1,3}(?:\s\d{3})+|\d+)(?:[,\.]\d+)?|[\-—–])",
                        after_ticker,
                    )

                    if len(tokens) >= 5:
                        max_price_token = tokens[-5]
                        price_num = self.clean_number(max_price_token)

                        if price_num and price_num >= 1000:
                            self.bulletin_data[found_ticker] = (
                                self.round_up_thousands(price_num)
                            )

    def download_latest_bulletin(self):
        self.log("Запрос свежего бюллетеня с spimex.com...", "info")
        self.update_status("Подключение к spimex.com...")
        url = "https://spimex.com/markets/oil_products/trades/results/"

        try:
            res = self.session.get(url, timeout=15)
            if res.status_code != 200:
                self.log(
                    f"[!] Ошибка загрузки страницы: HTTP {res.status_code}",
                    "error",
                )
                return False

            from bs4 import BeautifulSoup

            soup = BeautifulSoup(res.text, "html.parser")

            download_link = None
            for a in soup.find_all("a", href=True):
                href = a["href"].lower()
                if ".pdf" in href and (
                    "upload" in href or "bulletin" in href or "oil" in href
                ):
                    download_link = urllib.parse.urljoin(url, a["href"])
                    break

            if not download_link:
                self.log("[!] Ссылка на PDF-бюллетень не найдена.", "error")
                return False

            self.log(f"Загрузка PDF файла: {download_link}", "info")
            self.update_status("Скачивание PDF бюллетеня...")
            file_res = self.session.get(download_link, timeout=30)

            if file_res.status_code != 200 or not file_res.content.startswith(
                b"%PDF"
            ):
                self.log(
                    "[!] Скачанный файл не является корректным PDF.", "error"
                )
                return False

            date_match = re.search(
                r"(\d{8}|\d{2}\.\d{2}\.\d{4})", download_link
            )
            if date_match:
                d_str = date_match.group(1)
                if len(d_str) == 8 and d_str.isdigit():
                    self.bulletin_date_str = (
                        f"{d_str[6:8]}.{d_str[4:6]}.{d_str[:4]}"
                    )
                else:
                    self.bulletin_date_str = d_str
            else:
                self.bulletin_date_str = datetime.date.today().strftime(
                    "%d.%m.%Y"
                )

            try:
                dt = datetime.datetime.strptime(
                    self.bulletin_date_str, "%d.%m.%Y"
                )
                self.bulletin_date_str = dt.strftime("%d.%m.%Y")
            except ValueError:
                pass

            self.log(
                f"PDF успешно загружен. Дата бюллетеня: {self.bulletin_date_str}",
                "success",
            )
            self.update_status("Парсинг PDF...")

            self.parse_pdf_bytes(file_res.content)
            self.log(
                f"Извлечено позиций: {len(self.bulletin_data)} из {len(self.excel_tickers)}",
                "success",
            )
            return True

        except Exception as e:
            self.log(f"[!] Ошибка обработки PDF: {e}", "error")
            return False

    def run_monitoring(self):
        self.btn_run.config(
            state=tk.DISABLED, bg="#2D4D32", fg="#779977", cursor="arrow"
        )
        self.log("\n================ ЗАПУСК МОНИТОРИНГА ================")

        try:
            self.load_excel_tickers()

            if not self.download_latest_bulletin():
                self.log("[!] Процесс остановлен.", "error")
                self.update_status("Ошибка выполнения")
                return

            self.update_status("Запись данных в Excel...")
            wb = openpyxl.load_workbook(self.file_path)
            ws = wb.active

            date_columns = {}
            for col in range(3, ws.max_column + 1):
                val = ws.cell(row=2, column=col).value
                if val is not None:
                    if isinstance(val, (datetime.datetime, datetime.date)):
                        date_str = val.strftime("%d.%m.%Y")
                    else:
                        d_raw = str(val).strip()
                        d_parts = re.findall(r"\d+", d_raw)
                        if len(d_parts) == 3:
                            date_str = f"{int(d_parts[0]):02d}.{int(d_parts[1]):02d}.{d_parts[2]}"
                        else:
                            date_str = d_raw
                    date_columns[date_str] = col

            processed_count = 0
            for row in range(4, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=1).value

                if (
                    cell_val
                    and str(cell_val).strip() != "Тариф"
                    and len(str(cell_val).strip()) >= 6
                ):
                    ticker = str(cell_val).strip()
                    tariff_row = row - 1

                    if ticker in self.bulletin_data:
                        price = self.bulletin_data[ticker]
                        trade_date_str = self.bulletin_date_str

                        col_idx = date_columns.get(trade_date_str)
                        if not col_idx:
                            alt_date = f"{int(trade_date_str[:2])}.{int(trade_date_str[3:5])}.{trade_date_str[6:]}"
                            col_idx = date_columns.get(alt_date)

                        if col_idx:
                            col_letter = openpyxl.utils.get_column_letter(
                                col_idx
                            )
                            ws.cell(
                                row=tariff_row, column=col_idx
                            ).value = price
                            formula = (
                                f"={col_letter}{tariff_row}+$B${tariff_row}"
                            )
                            ws.cell(row=row, column=col_idx).value = formula

                            self.log(
                                f"[✓] {ticker} | Дата: {trade_date_str} | Макс. сумма: {price}",
                                "success",
                            )
                            processed_count += 1
                        else:
                            self.log(
                                f"[!] Дата {trade_date_str} отсутствует во 2-й строке Excel.",
                                "error",
                            )
                    else:
                        self.log(
                            f"[-] {ticker}: нет торгов на СПбМТСБ.", "info"
                        )

            output_file = self.file_path.replace(".xlsx", "_updated.xlsx")
            wb.save(output_file)

            self.log(
                f"\nУСПЕШНО ЗАВЕРШЕНО!\nРезультат сохранен: {output_file}",
                "success",
            )
            self.update_status(f"Завершено. Обновлено строк: {processed_count}")
            messagebox.showinfo(
                "Успех",
                f"Мониторинг успешно завершен!\n\nОбновлено позиций: {processed_count}\nФайл сохранен рядом с исходным.",
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