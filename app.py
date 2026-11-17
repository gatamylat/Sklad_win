import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from collections import Counter
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from uuid import uuid4

from stock_check import check_files, item_text, locations_text, show_number, write_report


APP_DIR = Path(os.getenv('LOCALAPPDATA') or Path.home() / '.local' / 'share') / 'Sklad_win'
SETTINGS = APP_DIR / 'settings.json'
LOGS = APP_DIR / 'logs'
EXCEL = [('Книга Excel', '*.xlsx *.xlsm')]
REGISTER = [('Реестр склада', '*.json')]


def key_vault():
    if sys.platform == 'win32':
        from keyring.backends.Windows import WinVaultKeyring

        return WinVaultKeyring()
    return None


def read_settings():
    if SETTINGS.exists():
        data = json.loads(SETTINGS.read_text(encoding='utf-8'))
        return str(data['purchases']), str(data['registry'])
    return '', ''


def save_settings(purchases, registry):
    APP_DIR.mkdir(parents=True, exist_ok=True)
    temporary = SETTINGS.with_suffix(f'.{uuid4().hex}.tmp')
    try:
        temporary.write_text(json.dumps({'purchases': purchases, 'registry': registry}, ensure_ascii=False), encoding='utf-8')
        temporary.replace(SETTINGS)
    finally:
        temporary.unlink(missing_ok=True)


class App:
    def __init__(self, root):
        self.root = root
        self.report = None
        self.input_files = ()
        self.events = queue.Queue()
        root.title('Склад Массивбург · Проверка заявки')
        root.geometry('1200x820')
        root.minsize(790, 620)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        panel = ttk.Frame(root, padding=20)
        panel.grid(sticky='nsew')
        panel.columnconfigure(1, weight=1)
        panel.rowconfigure(8, weight=3)
        panel.rowconfigure(9, weight=2)

        ttk.Label(panel, text='Проверка заявки', font=('Segoe UI', 22, 'bold')).grid(row=0, columnspan=3, sticky='w', pady=(0, 14))
        self.request = tk.StringVar()
        self.purchases = tk.StringVar()
        self.registry = tk.StringVar()
        self.key = tk.StringVar(value=os.getenv('DEEPSEEK_API_KEY', ''))
        self.path_input(panel, 1, 'Заявка конструктора', self.request, EXCEL)
        self.path_input(panel, 2, 'Закупки', self.purchases, EXCEL)
        self.path_input(panel, 3, 'Реестр сверки склад', self.registry, REGISTER)
        ttk.Label(panel, text='Ключ DeepSeek').grid(row=4, column=0, sticky='w', pady=4, padx=(0, 12))
        ttk.Entry(panel, textvariable=self.key, show='•').grid(row=4, column=1, sticky='ew', pady=4)
        ttk.Button(panel, text='Запомнить ключ', command=self.remember_key).grid(row=4, column=2, padx=(8, 0))

        controls = ttk.Frame(panel)
        controls.grid(row=5, columnspan=3, sticky='ew', pady=(12, 8))
        self.run_button = ttk.Button(controls, text='Проверить заявку', command=self.run)
        self.run_button.pack(side='left')
        self.export_button = ttk.Button(controls, text='Сохранить результат в Excel', command=self.export, state='disabled')
        self.export_button.pack(side='left', padx=8)
        ttk.Button(controls, text='Открыть папку журналов', command=self.open_logs).pack(side='left')

        self.status = tk.StringVar(value='Выберите файлы для проверки.')
        self.sources = tk.StringVar()
        ttk.Label(panel, textvariable=self.status).grid(row=6, columnspan=3, sticky='w', pady=(0, 4))
        ttk.Label(panel, textvariable=self.sources).grid(row=7, columnspan=3, sticky='w', pady=(0, 7))
        self.build_table(panel)
        detail_panel = ttk.LabelFrame(panel, text='Данные по выбранной строке', padding=8)
        detail_panel.grid(row=9, columnspan=3, sticky='nsew', pady=(10, 0))
        self.detail = tk.Text(detail_panel, height=7, wrap='word', font=('Segoe UI', 10), state='disabled')
        self.detail.pack(side='left', fill='both', expand=True)
        details_scroll = ttk.Scrollbar(detail_panel, command=self.detail.yview)
        details_scroll.pack(side='right', fill='y')
        self.detail.configure(yscrollcommand=details_scroll.set)

        try:
            purchases, registry = read_settings()
            self.purchases.set(purchases)
            self.registry.set(registry)
        except (ValueError, KeyError, OSError):
            self.status.set('Не удалось прочитать последние пути к источникам. Выберите файлы заново.')
        if not self.key.get():
            try:
                vault = key_vault()
                if vault:
                    self.key.set(vault.get_password('Sklad_win', 'DeepSeek') or '')
            except Exception:
                self.status.set('Ключ из хранилища Windows недоступен. Его можно ввести для этой проверки.')

    def path_input(self, panel, row, label, variable, filetypes):
        ttk.Label(panel, text=label).grid(row=row, column=0, sticky='w', pady=4, padx=(0, 12))
        ttk.Entry(panel, textvariable=variable).grid(row=row, column=1, sticky='ew', pady=4)

        def choose():
            name = filedialog.askopenfilename(parent=self.root, title=label, filetypes=filetypes)
            if name:
                variable.set(name)

        ttk.Button(panel, text='Выбрать…', command=choose).grid(row=row, column=2, padx=(8, 0))

    def build_table(self, panel):
        table_frame = ttk.Frame(panel)
        table_frame.grid(row=8, columnspan=3, sticky='nsew')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        columns = ('row', 'article', 'name', 'unit', 'need', 'take', 'buy', 'status', 'place')
        titles = ('Строка', 'Артикул', 'Позиция', 'Ед.', 'Нужно', 'Взять', 'Докупить', 'Результат', 'Полка')
        widths = (52, 110, 270, 55, 62, 62, 62, 155, 130)
        self.table = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        self.table.grid(row=0, column=0, sticky='nsew')
        for column, title, width in zip(columns, titles, widths):
            self.table.heading(column, text=title)
            self.table.column(column, width=width, minwidth=width, stretch=column == 'name',
                              anchor='e' if column in ('need', 'take', 'buy') else 'w')
        vertical = ttk.Scrollbar(table_frame, command=self.table.yview)
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal = ttk.Scrollbar(table_frame, orient='horizontal', command=self.table.xview)
        horizontal.grid(row=1, column=0, sticky='ew')
        self.table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.table.tag_configure('confirmed', background='#e6f3ee')
        self.table.tag_configure('partial', background='#eaf2f9')
        self.table.tag_configure('manual', background='#fff2d9')
        self.table.bind('<<TreeviewSelect>>', self.show_details)

    def remember_key(self):
        key = self.key.get().strip()
        try:
            vault = key_vault()
            if vault is None:
                self.status.set('Ключ можно использовать на время запуска. Хранилище Windows недоступно.')
                return
            if key:
                vault.set_password('Sklad_win', 'DeepSeek', key)
                self.status.set('Ключ сохранён в диспетчере учётных данных Windows.')
            elif vault.get_password('Sklad_win', 'DeepSeek'):
                vault.delete_password('Sklad_win', 'DeepSeek')
                self.status.set('Сохранённый ключ удалён.')
            else:
                self.status.set('Введите ключ DeepSeek, чтобы запомнить его.')
        except Exception:
            self.status.set('Не удалось сохранить ключ в Windows. Введённый ключ доступен до закрытия окна.')

    def run(self):
        paths = tuple(Path(value.get().strip().strip('"')).expanduser() for value in (self.request, self.purchases, self.registry))
        if any(not path.is_file() for path in paths):
            messagebox.showerror('Файлы для проверки', 'Выберите заявку Excel, выгрузку «Закупки» и реестр JSON.', parent=self.root)
            return
        if len({path.resolve() for path in paths}) != 3:
            messagebox.showerror('Файлы для проверки', 'Укажите отдельный файл для каждого из трёх полей.', parent=self.root)
            return
        self.status.set('Идёт проверка заявки…')
        self.sources.set('')
        self.fill_details('')
        self.table.delete(*self.table.get_children())
        self.run_button.configure(state='disabled')
        self.export_button.configure(state='disabled')
        self.report = None
        self.input_files = paths
        thread = threading.Thread(target=self.worker, args=(paths, self.key.get().strip()), daemon=True)
        thread.start()
        self.root.after(100, self.finish)

    def worker(self, paths, key):
        try:
            try:
                save_settings(str(paths[1].resolve()), str(paths[2].resolve()))
                settings_note = ''
            except OSError:
                settings_note = ' Пути к источникам не удалось запомнить.'
            report, _ = check_files(*paths, key, LOGS)
            self.events.put((report, settings_note))
        except Exception as error:
            self.events.put((None, str(error) or type(error).__name__))

    def finish(self):
        try:
            report, note = self.events.get_nowait()
        except queue.Empty:
            self.root.after(100, self.finish)
            return
        self.run_button.configure(state='normal')
        if report is None:
            self.status.set('Проверка прервана.')
            messagebox.showerror('Ошибка проверки', note, parent=self.root)
            return
        self.report = report
        counts = Counter(row['status'] for row in report['rows'])
        for index, row in enumerate(report['rows']):
            tag = {'Взять со склада': 'confirmed', 'Частично со склада': 'partial', 'Докупить': 'buy'}.get(row['status'], 'manual')
            values = (row['row'], row['article'], row['name'], row['unit'], *(show_number(row[key]) for key in ('need', 'take', 'buy')),
                      row['status'], locations_text(row))
            self.table.insert('', 'end', iid=str(index), values=values, tags=(tag,))
        self.status.set(' · '.join(f'{status}: {count}' for status, count in counts.items()))
        dates = ' · '.join(f'{role}: {(info.get("exported") or info["modified"])[:16].replace("T", " ")}'
                           for role, info in report['files'].items())
        self.sources.set(f'{dates}\n{report["ai"]}{note}')
        self.export_button.configure(state='normal')
        self.table.selection_set('0')

    def fill_details(self, text):
        self.detail.configure(state='normal')
        self.detail.delete('1.0', 'end')
        self.detail.insert('1.0', text)
        self.detail.configure(state='disabled')

    def show_details(self, _event):
        selection = self.table.selection()
        if not selection or not self.report:
            return
        row = self.report['rows'][int(selection[0])]
        lines = [row['reason']]
        for key, title in (('purchases', 'Закупки'), ('registry', 'Реестр'),
                           ('suggestions', 'Подсказки DeepSeek'), ('candidates', 'Кандидаты')):
            if row.get(key):
                lines.append(f'\n{title}')
                for item in row[key]:
                    relation = {'same': 'Та же позиция: ', 'alternative': 'Аналог: '}.get(item.get('relation'), '')
                    lines.append(relation + item_text(item))
        self.fill_details('\n'.join(lines))

    def export(self):
        name = filedialog.asksaveasfilename(parent=self.root, title='Сохранить результат в Excel',
                                           initialfile=f'{self.input_files[0].stem} — проверка.xlsx',
                                           filetypes=[('Excel', '*.xlsx')], defaultextension='.xlsx')
        if not name:
            return
        path = Path(name)
        try:
            if path.suffix.lower() != '.xlsx' or any(path.resolve() == source.resolve() or
                    (path.exists() and path.samefile(source)) for source in self.input_files):
                raise ValueError('Выберите отдельный файл с расширением .xlsx.')
            write_report(self.report, path)
        except (OSError, ValueError) as error:
            messagebox.showerror('Ошибка сохранения', str(error), parent=self.root)
            return
        messagebox.showinfo('Результат сохранён', str(path), parent=self.root)

    def open_logs(self):
        try:
            LOGS.mkdir(parents=True, exist_ok=True)
            if sys.platform == 'win32':
                os.startfile(LOGS)
            else:
                subprocess.Popen(['xdg-open', str(LOGS)])
        except OSError as error:
            messagebox.showerror('Папка журналов', str(error), parent=self.root)


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
