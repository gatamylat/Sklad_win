import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font, messagebox, ttk
from uuid import uuid4

from stock_check import article_key, check_files, inventory_text, item_text, locations_text, show_number, write_report


APP_DIR = Path(os.getenv('LOCALAPPDATA') or Path.home() / '.local' / 'share') / 'Sklad_win'
SETTINGS = APP_DIR / 'settings.json'
LOGS = APP_DIR / 'logs'
EXCEL = [('Книга Excel', '*.xlsx *.xlsm')]
REGISTER = [('Реестр склада', '*.json')]
ZOOMS = ('Авто', '100%', '115%', '130%', '150%', '175%', '200%')
FILTERS = ('Все позиции', 'Найдено на складе', 'Требуется уточнение')
CANVAS, PAPER, SIDEBAR = '#e5ecef', '#fbfcfd', '#eef3f5'
INK, MUTED, LINE = '#172c36', '#425b68', '#c2d1d8'
MAST, ACCENT, SHELF_MARK = '#203b49', '#235568', '#e3bf70'
TYPE_SCALE = {
    'body': (11, 'normal'), 'small': (10, 'normal'), 'action': (10, 'bold'),
    'heading': (15, 'bold'), 'title': (27, 'bold'), 'number': (17, 'bold'),
}
COLUMNS = (
    ('row', 'Стр.', 49), ('article', 'Артикул', 156), ('name', 'Номенклатура', 260),
    ('need', 'Нужно', 83), ('stock', 'По реестру', 125), ('free', 'Свободно', 86),
    ('place', 'Полка', 105), ('take', 'Взять', 62), ('buy', 'Докупить', 81), ('status', 'Статус', 175),
)


def enable_windows_dpi():
    if sys.platform == 'win32':
        from ctypes import c_void_p, windll

        try:
            if windll.user32.SetProcessDpiAwarenessContext(c_void_p(-4)):
                return
        except AttributeError:
            pass
        windll.user32.SetProcessDPIAware()


def key_vault():
    if sys.platform == 'win32':
        from keyring.backends.Windows import WinVaultKeyring

        return WinVaultKeyring()
    return None


def read_settings():
    try:
        data = json.loads(SETTINGS.read_text(encoding='utf-8'))
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, ValueError, OSError):
        return {}


def save_settings(changes):
    APP_DIR.mkdir(parents=True, exist_ok=True)
    temporary = SETTINGS.with_suffix(f'.{uuid4().hex}.tmp')
    try:
        data = {**read_settings(), **changes}
        temporary.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
        temporary.replace(SETTINGS)
    finally:
        temporary.unlink(missing_ok=True)


def unit_quantity(quantity, unit):
    return f'{show_number(quantity)} {unit}'.strip() if quantity is not None else '—'


def zoom_factor(choice, screen_width, dpi):
    if choice != 'Авто':
        return int(choice[:-1]) / 100
    return 1.15 if screen_width >= 2600 and dpi < 1.2 else 1


class App:
    def __init__(self, root):
        self.root = root
        self.report = None
        self.input_files = ()
        self.events = queue.Queue()
        self.started = False
        self.dpi = float(root.tk.call('tk', 'scaling')) * 0.75
        self.screen = (root.winfo_screenwidth(), root.winfo_screenheight())
        self.style = ttk.Style(root)
        available = set(font.families(root))
        face = next((name for name in ('Source Sans 3', 'Inter', 'Segoe UI', 'DejaVu Sans') if name in available),
                    font.nametofont('TkDefaultFont', root=root).actual('family'))
        self.fonts = {key: font.Font(root, family=face, size=size, weight=weight)
                      for key, (size, weight) in TYPE_SCALE.items()}
        self.configure_styles()
        prefs = read_settings()
        self.path_values = {
            'Заявка': tk.StringVar(), 'Закупки': tk.StringVar(value=str(prefs.get('purchases') or '')),
            'Реестр': tk.StringVar(value=str(prefs.get('registry') or '')),
        }
        selected_zoom = prefs.get('zoom')
        self.zoom = tk.StringVar(value=selected_zoom if selected_zoom in ZOOMS else ZOOMS[0])
        self.key = tk.StringVar(value=os.getenv('DEEPSEEK_API_KEY', ''))
        self.key_note = tk.StringVar(value='Вставьте ключ из буфера обмена.')
        self.status = tk.StringVar(value='Выберите файлы для проверки.')
        self.sources = tk.StringVar()
        self.filter = tk.StringVar(value=FILTERS[0])
        self.query = tk.StringVar()
        self.stats = [tk.StringVar(value='—') for _ in range(3)]

        root.title('Склад Массивбург | Проверка заявки')
        root.configure(background=CANVAS)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)
        self.build_header()
        body = ttk.Frame(root, style='Canvas.TFrame', padding=(self.px(20), self.px(16)))
        body.grid(row=1, column=0, sticky='nsew')
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        self.build_sidebar(body)
        self.build_workspace(body)
        ttk.Label(root, textvariable=self.status, style='Status.TLabel', anchor='w', padding=(self.px(20), self.px(8))).grid(
            row=2, column=0, sticky='ew')

        self.set_zoom()
        width, height = self.screen
        root.geometry(f'{round(width * 0.91)}x{round(height * 0.87)}')
        root.minsize(min(self.px(1030), round(width * .96)), min(self.px(660), round(height * .90)))
        if not self.key.get():
            try:
                vault = key_vault()
                if vault:
                    self.key.set(vault.get_password('Sklad_win', 'DeepSeek') or '')
            except Exception:
                self.key_note.set('Хранилище ключей Windows недоступно.')
        if self.key.get():
            self.key_note.set('Ключ установлен.')
        self.started = True

    def px(self, size):
        return max(1, round(size * self.dpi))

    def configure_styles(self):
        self.style.theme_use('clam')
        f = self.fonts
        self.style.configure('Canvas.TFrame', background=CANVAS)
        self.style.configure('Paper.TFrame', background=PAPER)
        self.style.configure('Sidebar.TFrame', background=SIDEBAR)
        self.style.configure('Mast.TFrame', background=MAST)
        self.style.configure('Mast.TLabel', background=MAST, foreground=PAPER, font=f['title'])
        self.style.configure('Brand.TLabel', background=MAST, foreground='#d9e6ec', font=f['small'])
        self.style.configure('ShelfMark.TFrame', background=SHELF_MARK)
        self.style.configure('Rule.TFrame', background=LINE)
        self.style.configure('SideHeading.TLabel', background=SIDEBAR, foreground=INK, font=f['heading'])
        self.style.configure('SideField.TLabel', background=SIDEBAR, foreground=INK, font=f['action'])
        self.style.configure('SideNote.TLabel', background=SIDEBAR, foreground=MUTED, font=f['small'])
        self.style.configure('Paper.TLabel', background=PAPER, foreground=INK, font=f['action'])
        self.style.configure('Section.TLabel', background=PAPER, foreground=INK, font=f['heading'])
        self.style.configure('Metric.TLabel', background=PAPER, foreground=MAST, font=f['number'])
        self.style.configure('Muted.TLabel', background=PAPER, foreground=MUTED, font=f['small'])
        self.style.configure('Status.TLabel', background=SIDEBAR, foreground=MUTED, font=f['small'])
        self.style.configure('TButton', font=f['action'], padding=(self.px(11), self.px(6)), borderwidth=1, relief='flat',
                             background=PAPER, foreground=INK, bordercolor=LINE, lightcolor=PAPER, darkcolor=PAPER,
                             focusthickness=2, focuscolor=ACCENT)
        self.style.map('TButton', background=[('pressed', '#d1e1e8'), ('active', '#e0ebf0')],
                       bordercolor=[('focus', ACCENT)], foreground=[('disabled', '#556a74')])
        self.style.configure('Primary.TButton', foreground=PAPER, background=ACCENT, bordercolor=ACCENT,
                             lightcolor=ACCENT, darkcolor=ACCENT, padding=(self.px(12), self.px(8)), focuscolor=PAPER)
        self.style.map('Primary.TButton', background=[('disabled', '#dce6eb'), ('pressed', '#214555'), ('active', MAST)],
                       bordercolor=[('disabled', LINE), ('focus', ACCENT)], foreground=[('disabled', MUTED)])
        self.style.configure('TEntry', padding=(self.px(8), self.px(5)), bordercolor=LINE, fieldbackground=PAPER,
                             foreground=INK, insertcolor=INK, lightcolor=PAPER, darkcolor=PAPER)
        self.style.map('TEntry', bordercolor=[('focus', ACCENT)], lightcolor=[('focus', ACCENT)])
        self.style.configure('TCombobox', padding=(self.px(8), self.px(4)), bordercolor=LINE, fieldbackground=PAPER,
                             foreground=INK, arrowcolor=MAST)
        self.style.map('TCombobox', fieldbackground=[('readonly', PAPER)], bordercolor=[('focus', ACCENT)],
                       lightcolor=[('focus', ACCENT)])
        self.style.configure('Treeview', font=f['body'], background=PAPER, fieldbackground=PAPER,
                             foreground=INK, borderwidth=0)
        self.style.configure('Treeview.Heading', font=f['action'], background='#d7e5ec',
                             foreground=MAST, padding=(self.px(7), self.px(8)), relief='flat', borderwidth=1,
                             bordercolor=LINE, lightcolor='#d7e5ec', darkcolor='#d7e5ec')
        self.style.map('Treeview.Heading', background=[('active', '#d7e5ec')])
        self.style.map('Treeview', background=[('selected', '#b9d9e4')], foreground=[('selected', INK)])
        self.style.configure('TScrollbar', arrowsize=self.px(13), background='#cfdee5', troughcolor=SIDEBAR,
                             bordercolor=LINE, arrowcolor=MUTED)

    def build_header(self):
        top = ttk.Frame(self.root, style='Mast.TFrame', padding=(self.px(20), self.px(12)))
        top.grid(row=0, column=0, sticky='ew')
        top.columnconfigure(1, weight=1)
        ttk.Frame(top, width=self.px(5), style='ShelfMark.TFrame').grid(
            row=0, column=0, rowspan=2, sticky='ns', padx=(0, self.px(12)), pady=self.px(3))
        ttk.Label(top, text='Склад Массивбург', style='Brand.TLabel').grid(row=0, column=1, sticky='w')
        ttk.Label(top, text='Проверка заявки', style='Mast.TLabel').grid(row=1, column=1, sticky='w')
        ttk.Label(top, text='Масштаб', style='Brand.TLabel').grid(row=0, column=2, rowspan=2, padx=(0, self.px(9)))
        ttk.Combobox(top, textvariable=self.zoom, values=ZOOMS, state='readonly', width=8,
                     font=self.fonts['body']).grid(row=0, column=3, rowspan=2, sticky='e')
        self.zoom.trace_add('write', lambda *_: self.set_zoom())

    def build_sidebar(self, parent):
        panel = ttk.Frame(parent, style='Sidebar.TFrame', padding=self.px(16))
        panel.grid(row=0, column=0, sticky='ns', padx=(0, self.px(15)))
        panel.columnconfigure(0, weight=1)
        ttk.Label(panel, text='Файлы проверки', style='SideHeading.TLabel').grid(
            row=0, columnspan=2, sticky='w', pady=(0, self.px(15)))
        for row, (role, label, types) in enumerate((
            ('Заявка', 'Заявка конструктора', EXCEL),
            ('Закупки', 'Закупки', EXCEL),
            ('Реестр', 'Реестр сверки склада', REGISTER),
        )):
            self.path_input(panel, 1 + 2 * row, role, label, types)
        ttk.Frame(panel, height=self.px(1), style='Rule.TFrame').grid(row=7, columnspan=2, sticky='ew', pady=self.px(13))
        ttk.Label(panel, text='DeepSeek', style='SideHeading.TLabel').grid(row=8, columnspan=2, sticky='w')
        self.key_entry = ttk.Entry(panel, textvariable=self.key, show='•', font=self.fonts['body'], width=25)
        self.key_entry.grid(row=9, columnspan=2, sticky='ew', pady=(self.px(8), self.px(5)))
        self.key_entry.bind('<Control-v>', self.paste_key)
        self.key_entry.bind('<Control-V>', self.paste_key)
        self.key_entry.bind('<Control-a>', self.select_key)
        menu = tk.Menu(self.key_entry, tearoff=False)
        menu.add_command(label='Вставить', command=self.paste_key)
        menu.add_command(label='Выделить всё', command=self.select_key)
        self.key_entry.bind('<Button-3>', lambda event: menu.tk_popup(event.x_root, event.y_root))
        ttk.Button(panel, text='Вставить', command=self.paste_key).grid(row=10, column=0, sticky='ew',
                                                                       padx=(0, self.px(5)))
        ttk.Button(panel, text='Запомнить ключ', command=self.remember_key).grid(row=10, column=1, sticky='ew')
        self.key_feedback = ttk.Label(panel, textvariable=self.key_note, style='SideNote.TLabel')
        self.key_feedback.grid(row=11, columnspan=2, sticky='w', pady=(self.px(6), self.px(14)))
        self.run_button = ttk.Button(panel, text='Проверить заявку', style='Primary.TButton', command=self.run)
        self.run_button.grid(row=12, columnspan=2, sticky='ew')
        ttk.Button(panel, text='Открыть журналы', command=self.open_logs).grid(row=13, columnspan=2, sticky='ew',
                                                                               pady=(self.px(8), 0))

    def path_input(self, panel, row, role, label, filetypes):
        ttk.Label(panel, text=label, style='SideField.TLabel').grid(row=row, column=0, columnspan=2, sticky='w')
        variable = self.path_values[role]
        entry = ttk.Entry(panel, textvariable=variable, font=self.fonts['body'], width=23)
        entry.grid(row=row + 1, column=0, sticky='ew', pady=(self.px(4), self.px(12)), padx=(0, self.px(5)))

        def choose():
            path = filedialog.askopenfilename(parent=self.root, title=label, filetypes=filetypes)
            if path:
                variable.set(path)

        def reveal(*_):
            entry.xview_moveto(1)

        variable.trace_add('write', reveal)
        entry.after_idle(reveal)
        ttk.Button(panel, text='Обзор…', command=choose).grid(row=row + 1, column=1, sticky='ew',
                                                               pady=(self.px(4), self.px(12)))

    def build_workspace(self, parent):
        panel = ttk.Frame(parent, style='Paper.TFrame', padding=self.px(16))
        panel.grid(row=0, column=1, sticky='nsew')
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(3, weight=4)
        panel.rowconfigure(4, weight=1)

        header = ttk.Frame(panel, style='Paper.TFrame')
        header.grid(row=0, column=0, sticky='ew', pady=(0, self.px(10)))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text='Результат', style='Section.TLabel').grid(row=0, column=0, sticky='w')
        self.export_full = ttk.Button(header, text='Полный отчёт', state='disabled', command=self.export)
        self.export_full.grid(row=0, column=1, padx=(0, self.px(6)))
        self.export_inventory = ttk.Button(header, text='Выгрузить наличие', state='disabled',
                                           style='Primary.TButton', command=lambda: self.export(inventory_only=True))
        self.export_inventory.grid(row=0, column=2)

        metrics = ttk.Frame(panel, style='Paper.TFrame', padding=(0, self.px(5), 0, self.px(9)))
        metrics.grid(row=1, column=0, sticky='ew', pady=(0, self.px(8)))
        for col, title in enumerate(('Позиций в заявке', 'Найдено на складе', 'Есть с количеством')):
            column = col * 2
            metrics.columnconfigure(column, weight=1, uniform='metrics')
            inset = self.px(14) if col else 0
            if col:
                ttk.Frame(metrics, width=self.px(1), style='Rule.TFrame').grid(
                    row=0, column=column - 1, rowspan=2, sticky='ns', pady=self.px(3))
            ttk.Label(metrics, text=title, style='Muted.TLabel').grid(row=0, column=column, sticky='w', padx=(inset, 0))
            ttk.Label(metrics, textvariable=self.stats[col], style='Metric.TLabel').grid(
                row=1, column=column, sticky='w', padx=(inset, 0))
        tools = ttk.Frame(panel, style='Paper.TFrame')
        tools.grid(row=2, column=0, sticky='ew', pady=(0, self.px(7)))
        tools.columnconfigure(2, weight=1)
        ttk.Combobox(tools, textvariable=self.filter, values=FILTERS, state='readonly', width=21,
                     font=self.fonts['body']).grid(row=0, column=0, sticky='w')
        ttk.Label(tools, text='Поиск', style='Muted.TLabel').grid(row=0, column=1, padx=self.px(9))
        ttk.Entry(tools, textvariable=self.query, font=self.fonts['body']).grid(row=0, column=2, sticky='ew')
        self.filter.trace_add('write', lambda *_: self.fill_table())
        self.query.trace_add('write', lambda *_: self.fill_table())
        self.build_table(panel)

        details = ttk.Frame(panel, style='Paper.TFrame', padding=(0, self.px(13), 0, 0))
        details.grid(row=4, column=0, sticky='nsew')
        details.columnconfigure(0, weight=1)
        details.rowconfigure(1, weight=1)
        ttk.Label(details, text='Данные по позиции', style='Paper.TLabel').grid(row=0, column=0, sticky='w',
                                                                               pady=(0, self.px(4)))
        self.detail = tk.Text(details, height=4, wrap='word', font=self.fonts['body'], state='disabled',
                              foreground=INK, background=SIDEBAR, borderwidth=0, spacing1=self.px(2),
                              padx=self.px(12), pady=self.px(8), highlightthickness=self.px(1),
                              highlightbackground=LINE, highlightcolor=ACCENT)
        self.detail.grid(row=1, column=0, sticky='nsew')
        scroll = ttk.Scrollbar(details, command=self.detail.yview)
        scroll.grid(row=1, column=1, sticky='ns')
        self.detail.configure(yscrollcommand=scroll.set)
        ttk.Label(panel, textvariable=self.sources, style='Muted.TLabel', padding=(0, self.px(7), 0, 0)).grid(
            row=5, column=0, sticky='ew')

    def build_table(self, parent):
        panel = ttk.Frame(parent, style='Paper.TFrame')
        panel.grid(row=3, column=0, sticky='nsew')
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(0, weight=1)
        self.table = ttk.Treeview(panel, columns=[col for col, _, _ in COLUMNS], show='headings', selectmode='browse')
        self.table.grid(row=0, column=0, sticky='nsew')
        for col, title, _ in COLUMNS:
            self.table.heading(col, text=title, anchor='w')
        vertical = ttk.Scrollbar(panel, command=self.table.yview)
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal = ttk.Scrollbar(panel, orient='horizontal', command=self.table.xview)
        horizontal.grid(row=1, column=0, sticky='ew')
        self.table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.table.tag_configure('confirmed', background='#dcece8')
        self.table.tag_configure('found', background='#f7edd8')
        self.table.tag_configure('stripe', background='#f0f5f7')
        self.table.bind('<<TreeviewSelect>>', self.show_details)
        self.empty = ttk.Label(panel, text='Загрузите файлы и запустите проверку.', style='Muted.TLabel')
        self.empty.place(relx=.5, rely=.5, anchor='center')

    def set_zoom(self):
        choice = self.zoom.get()
        factor = zoom_factor(choice, self.screen[0], self.dpi)
        for key, (size, _) in TYPE_SCALE.items():
            self.fonts[key].configure(size=round(size * factor))
        self.style.configure('Treeview', rowheight=self.px(31 * factor))
        for col, _, width in COLUMNS:
            size = self.px(width * factor)
            self.table.column(col, width=size, minwidth=size, stretch=col == 'name',
                              anchor='e' if col in ('need', 'free', 'take', 'buy') else 'w')
        self.key_feedback.configure(wraplength=self.px(240 * factor))
        if self.started:
            try:
                save_settings({'zoom': choice})
            except OSError:
                pass

    def select_key(self, _event=None):
        self.key_entry.focus_set()
        self.key_entry.selection_range(0, 'end')
        return 'break'

    def paste_key(self, _event=None):
        try:
            clipboard = self.root.clipboard_get().strip()
        except tk.TclError:
            clipboard = ''
        if not clipboard or any(char.isspace() for char in clipboard):
            self.key_note.set('Скопируйте один ключ API и нажмите «Вставить».')
        else:
            self.key.set(clipboard)
            self.key_entry.icursor('end')
            self.key_note.set('Ключ вставлен.')
        self.key_entry.focus_set()
        return 'break'

    def remember_key(self):
        key = self.key.get().strip()
        try:
            vault = key_vault()
            if vault is None:
                self.key_note.set('Ключ действует до закрытия окна.')
            elif key:
                vault.set_password('Sklad_win', 'DeepSeek', key)
                self.key_note.set('Ключ сохранён в Windows.')
            elif vault.get_password('Sklad_win', 'DeepSeek'):
                vault.delete_password('Sklad_win', 'DeepSeek')
                self.key_note.set('Сохранённый ключ удалён.')
            else:
                self.key_note.set('Вставьте ключ, чтобы запомнить его.')
        except Exception:
            self.key_note.set('Сохранение недоступно. Ключ действует до закрытия окна.')

    def run(self):
        paths = tuple(Path(value.get().strip().strip('"')).expanduser() for value in self.path_values.values())
        if any(not path.is_file() for path in paths):
            self.status.set('Выберите заявку Excel, выгрузку «Закупки» и реестр JSON.')
            return
        if len({path.resolve() for path in paths}) != 3:
            self.status.set('Укажите отдельный файл для каждого поля.')
            return
        try:
            save_settings({'purchases': str(paths[1].resolve()), 'registry': str(paths[2].resolve())})
        except OSError:
            pass
        self.status.set('Читаю файлы и сверяю остатки…')
        self.sources.set('')
        self.fill_details('')
        self.table.delete(*self.table.get_children())
        self.empty.place_forget()
        self.run_button.configure(state='disabled')
        self.export_full.configure(state='disabled')
        self.export_inventory.configure(state='disabled')
        self.report = None
        self.input_files = paths
        for stat in self.stats:
            stat.set('—')
        threading.Thread(target=self.worker, args=(paths, self.key.get().strip()), daemon=True).start()
        self.root.after(100, self.finish)

    def worker(self, paths, key):
        try:
            report, _ = check_files(*paths, key, LOGS)
            self.events.put((report, ''))
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
            self.empty.place(relx=.5, rely=.5, anchor='center')
            messagebox.showerror('Ошибка проверки', note, parent=self.root)
            return
        self.report = report
        rows = report['rows']
        found = sum(row['stock_found'] for row in rows)
        counted = sum(row['stock_found'] and ((row['stock_qty'] or 0) > 0 or (row['free_qty'] or 0) > 0) for row in rows)
        for stat, value in zip(self.stats, (len(rows), found, counted)):
            stat.set(str(value))
        self.query.set('')
        self.filter.set(FILTERS[1] if found else FILTERS[0])
        self.status.set(report['ai'])
        dates = [f'{role}   {(info.get("exported") or info["modified"])[:16].replace("T", " ")}'
                 for role, info in report['files'].items()]
        self.sources.set('     '.join(dates))
        self.export_full.configure(state='normal')
        self.export_inventory.configure(state='normal')

    def fill_table(self):
        if self.report is None:
            return
        self.table.delete(*self.table.get_children())
        query = self.query.get().strip().casefold()
        code_query = article_key(query)
        view = self.filter.get()
        for row in self.report['rows']:
            if (view == FILTERS[1] and not row['stock_found']) or (view == FILTERS[2] and row['take'] is not None):
                continue
            article = row['article'] or row['matched_article']
            shelf = locations_text(row) or ', '.join(row['stock_places'])
            row_articles = [article, *(code for item in row['registry'] for code in item['articles'])]
            searchable = ' '.join([row['name'], shelf, *row_articles]).casefold()
            if query and query not in searchable and not (code_query and any(code_query in article_key(code) for code in row_articles)):
                continue
            tag = 'confirmed' if (row['take'] or 0) > 0 else 'found' if row['stock_found'] else 'stripe' if row['row'] % 2 else ''
            values = (row['row'], article, row['name'], unit_quantity(row['need'], row['unit']), inventory_text(row),
                      unit_quantity(row['free_qty'], row['free_unit']), shelf,
                      show_number(row['take']), show_number(row['buy']), row['status'])
            self.table.insert('', 'end', iid=str(row['row']), values=values, tags=(tag,))
        items = self.table.get_children()
        if items:
            self.empty.place_forget()
            self.table.selection_set(items[0])
            self.table.focus(items[0])
            self.table.see(items[0])
        else:
            self.empty.configure(text='По этому фильтру позиций нет.')
            self.empty.place(relx=.5, rely=.5, anchor='center')
            self.fill_details('')

    def fill_details(self, text):
        self.detail.configure(state='normal')
        self.detail.delete('1.0', 'end')
        self.detail.insert('1.0', text)
        self.detail.configure(state='disabled')

    def show_details(self, _event):
        selection = self.table.selection()
        if not selection or not self.report:
            return
        row = next(row for row in self.report['rows'] if row['row'] == int(selection[0]))
        lines = [row['name'], row['reason']]
        for key, title in (('purchases', 'Закупки'), ('registry', 'Реестр'),
                           ('suggestions', 'Подсказки DeepSeek'), ('candidates', 'Кандидаты')):
            if row.get(key):
                lines.append(f'\n{title}')
                for item in row[key]:
                    relation = {'same': 'Та же позиция: ', 'alternative': 'Аналог: '}.get(item.get('relation'), '')
                    lines.append(relation + item_text(item))
        self.fill_details('\n'.join(lines))

    def export(self, inventory_only=False):
        if self.report is None:
            return
        label = 'Выгрузить наличие' if inventory_only else 'Сохранить полный отчёт'
        suffix = 'наличие' if inventory_only else 'проверка'
        name = filedialog.asksaveasfilename(parent=self.root, title=label,
                                           initialfile=f'{self.input_files[0].stem} — {suffix}.xlsx',
                                           filetypes=[('Excel', '*.xlsx')], defaultextension='.xlsx')
        if not name:
            return
        path = Path(name)
        try:
            if path.suffix.lower() != '.xlsx' or any(path.resolve() == source.resolve() or
                    (path.exists() and source.exists() and path.samefile(source)) for source in self.input_files):
                raise ValueError('Выберите отдельный файл с расширением .xlsx.')
            write_report(self.report, path, inventory_only=inventory_only)
        except (OSError, ValueError) as error:
            messagebox.showerror('Ошибка сохранения', str(error), parent=self.root)
            return
        self.status.set(f'Сохранено: {path.name}')

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
    enable_windows_dpi()
    root = tk.Tk()
    App(root)
    root.mainloop()
