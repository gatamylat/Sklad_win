import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from concurrent.futures import CancelledError
from math import hypot, sqrt
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
THEMES = {
    'light': dict(bg='#FAF9F5', panel='#FFFFFF', fg='#1C1C1E', muted='#625D55', line='#E2DED5',
                  selected='#F0E9DA', accent='#795222', on_accent='#FFFFFF'),
    'dark': dict(bg='#1A1A1D', panel='#2C2C2E', fg='#F0EFEA', muted='#A5A6A7', line='#343537',
                 selected='#34302B', accent='#DDB57D', on_accent='#1A1A1D'),
}
COLUMNS = (
    ('article', 'Артикул', 127), ('name', 'Номенклатура', 204), ('need', 'Нужно всего', 105),
    ('stock', 'По реестру', 112), ('free', 'Свободно', 86), ('place', 'Полка', 85),
    ('take', 'Взять', 60), ('buy', 'Докупить', 80), ('status', 'Статус', 150),
)
NUMERIC = ('need', 'stock', 'free', 'take', 'buy')


def enable_windows_dpi():
    if sys.platform == 'win32':
        from ctypes import c_void_p, windll

        try:
            if windll.user32.SetProcessDpiAwarenessContext(c_void_p(-4)):
                return
        except AttributeError:
            pass
        windll.user32.SetProcessDPIAware()


def load_windows_fonts():
    if sys.platform == 'win32':
        from ctypes import c_wchar_p, windll

        directory = Path(__file__).resolve().parent / 'fonts'
        for name in ('SkladDisplay-Medium.ttf', 'SkladSans-Regular.ttf'):
            windll.gdi32.AddFontResourceExW(c_wchar_p(str(directory / name)), 0x10, None)


def key_vault():
    if sys.platform == 'win32':
        from keyring.backends.Windows import WinVaultKeyring

        return WinVaultKeyring()
    return None


def read_settings():
    try:
        data = json.loads(SETTINGS.read_text(encoding='utf-8'))
        return data if isinstance(data, dict) else {}
    except (ValueError, OSError):
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


def rounded_tile(root, radius, fill, border):
    size = radius * 2 + 3
    image = tk.PhotoImage(master=root, width=size, height=size)

    def paint(color, inset):
        curve = radius - inset
        for y in range(inset, size - inset):
            distance = max(0, curve - min(y - inset, size - 1 - inset - y))
            margin = inset + round(curve - sqrt(curve ** 2 - distance ** 2))
            image.put(color, to=(margin, y, size - margin, y + 1))

    paint(border, 0)
    paint(fill, 1)
    return image


class App:
    def __init__(self, root):
        self.root = root
        self.report = None
        self.input_files = ()
        self.events = queue.Queue()
        self.revision = 0
        self.timer = None
        self.running = None
        self.cancellation = None
        self.theme_images = {}
        self.tooltip_timer = None
        self.tooltip_window = None
        self.dpi = float(root.tk.call('tk', 'scaling')) * .75
        self.screen = (root.winfo_screenwidth(), root.winfo_screenheight())
        self.style = ttk.Style(root)
        self.style.theme_use('clam')
        prefs = read_settings()
        self.theme = prefs.get('theme') if prefs.get('theme') in THEMES else 'light'
        self.zoom = tk.StringVar(value=prefs.get('zoom') if prefs.get('zoom') in ZOOMS else 'Авто')
        self.requests = []
        self.sources = {role: Path(prefs[role]) if prefs.get(role) else None for role in ('purchases', 'registry')}
        self.query = tk.StringVar()
        self.filter = tk.StringVar(value=FILTERS[0])
        self.summary = tk.StringVar(value='—')
        self.status = tk.StringVar()
        self.key = tk.StringVar(value=os.getenv('DEEPSEEK_API_KEY', ''))
        self.key_note = tk.StringVar(value='Вставьте ключ из буфера обмена.')
        self.key_window = None
        self.key_entry = None
        if not self.key.get():
            try:
                vault = key_vault()
                if vault:
                    self.key.set(vault.get_password('Sklad_win', 'DeepSeek') or '')
            except Exception:
                self.key_note.set('Хранилище ключей Windows недоступно.')
        if self.key.get():
            self.key_note.set('Ключ установлен.')

        available = set(font.families(root))
        fallback = font.nametofont('TkDefaultFont', root=root).actual('family')
        body = next((name for name in ('Segoe UI', 'DejaVu Sans') if name in available), fallback)
        latin = next((name for name in ('Sklad Sans', 'DM Sans') if name in available), body)
        heading = next((name for name in ('Sklad Display', 'Playfair Display', 'Georgia', 'DejaVu Serif')
                        if name in available), body)
        self.fonts = {name: font.Font(root, family=family, size=-size, weight=weight)
                      for name, (family, size, weight) in {
                          'title': (heading, 22, 'normal'), 'section': (heading, 18, 'normal'),
                          'body': (body, 14, 'normal'), 'strong': (body, 14, 'bold'),
                          'button': (body, 13, 'normal'), 'small': (body, 12, 'normal'),
                          'latin': (latin, 13, 'normal'), 'latin_data': (latin, 14, 'normal'),
                      }.items()}
        self.sizes = dict(title=22, section=18, body=14, strong=14, button=13, small=12, latin=13, latin_data=14)

        root.title('Склад Массивбург')
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)
        self.build_header()
        self.build_files()
        self.build_workspace()
        self.status_label = ttk.Label(root, textvariable=self.status, style='Small.TLabel', anchor='w',
                                      padding=(self.px(24), self.px(9)))
        self.status_label.grid(row=3, column=0, sticky='ew')
        self.status_label.bind('<Configure>', lambda event: self.status_label.configure(
            wraplength=max(self.px(100), event.width - self.px(48))))
        self.zoom.trace_add('write', lambda *_: self.apply_zoom())
        self.query.trace_add('write', lambda *_: self.fill_table())
        self.filter.trace_add('write', lambda *_: self.fill_table())
        root.bind('<Control-r>', self.manual_run)
        root.bind('<Control-o>', lambda _: self.add_requests())
        self.apply_zoom(remember=False)
        self.refresh_files()
        self.status.set(self.missing())
        width, height = self.screen
        root.geometry(f'{round(width * .92)}x{round(height * .89)}')
        root.minsize(min(self.px(980), round(width * .96)), min(self.px(570), round(height * .91)))

    def px(self, size):
        return max(1, round(size * self.dpi))

    def save_preferences(self, **changes):
        try:
            save_settings(changes)
        except OSError:
            self.status.set('Не удалось сохранить настройки.')

    def make_icon(self, kind):
        size = self.px(18 * self.factor)
        icon = tk.PhotoImage(master=self.root, width=size, height=size)
        color = THEMES[self.theme]['fg']
        for y in range(size):
            for x in range(size):
                dx, dy = (x + .5) * 18 / size - 9, (y + .5) * 18 / size - 9
                distance = hypot(dx, dy)
                if kind == 'menu':
                    filled = abs(dx) < 7 and any(abs(dy - line) < .8 for line in (-4, 0, 4))
                elif kind == 'sun':
                    filled = 3.4 < distance < 4.8 or (5.9 < distance < 8.4 and
                        (min(abs(dx), abs(dy)) < .85 or abs(abs(dx) - abs(dy)) < 1.05))
                else:
                    filled = hypot(dx + 1, dy - .7) < 7.2 and hypot(dx - 2.7, dy + 2.8) > 6.4
                if filled:
                    icon.put(color, (x, y))
        return icon

    def configure_buttons(self, colors):
        radius = max(2, self.px(7 * self.factor))
        key = self.theme, radius
        if key not in self.theme_images:
            images = {}
            for variant, fill, border in (
                ('quiet', colors['bg'], colors['bg']), ('field', colors['panel'], colors['line']),
                ('action', colors['accent'], colors['accent']),
            ):
                prefix = f'Surf.{self.theme}.{radius}.{variant}'
                states = [rounded_tile(self.root, radius, fill, border),
                          rounded_tile(self.root, radius, colors['selected'], colors['accent']),
                          rounded_tile(self.root, radius, colors['selected'] if variant == 'action' else fill, colors['accent']),
                          rounded_tile(self.root, radius, colors['panel'], colors['line'])]
                self.style.element_create(prefix, 'image', states[0], ('active', states[1]), ('focus', states[2]),
                                          ('disabled', states[3]), border=radius + 1, padding=1, sticky='nsew')
                images[variant] = (prefix, states)
            self.theme_images[key] = images
        height = self.px(34 * self.factor)
        vertical = max(2, (height - self.fonts['button'].metrics('linespace') - 2) // 2)
        for style, variant in (('TButton', 'field'), ('Quiet.TButton', 'quiet'), ('LatinQuiet.TButton', 'quiet'), ('Action.TButton', 'action'),
                               ('Icon.TButton', 'quiet'), ('File.TButton', 'field')):
            element = self.theme_images[key][variant][0]
            self.style.layout(style, [(element, {'sticky': 'nsew', 'children': [
                ('Button.padding', {'sticky': 'nsew', 'children': [('Button.label', {'sticky': 'nsew'})]})
            ]})])
            padding = (self.px(10), vertical)
            if style == 'Icon.TButton':
                padding = (self.px(7 * self.factor), self.px(7 * self.factor))
            role = 'latin' if style == 'LatinQuiet.TButton' else 'button'
            self.style.configure(style, background=colors['bg'], font=self.fonts[role], padding=padding,
                                 foreground=colors['on_accent'] if variant == 'action' else colors['fg'],
                                 anchor='center')
            self.style.map(style, foreground=[('disabled', colors['muted']), ('active', colors['fg']), ('focus', colors['fg'])])

    def apply_theme(self):
        c = THEMES[self.theme]
        self.root.configure(bg=c['bg'])
        self.style.configure('.', background=c['bg'], foreground=c['fg'], font=self.fonts['body'])
        self.style.configure('TFrame', background=c['bg'])
        self.style.configure('Rule.TFrame', background=c['line'])
        for style, role, color in (('Title', 'title', 'fg'), ('Section', 'section', 'fg'), ('Small', 'small', 'muted'),
                                   ('Body', 'body', 'fg')):
            self.style.configure(f'{style}.TLabel', background=c['bg'], foreground=c[color], font=self.fonts[role])
        self.style.configure('SearchHint.TLabel', background=c['panel'], foreground=c['muted'], font=self.fonts['body'])
        self.configure_buttons(c)
        for style in ('TEntry', 'TCombobox'):
            self.style.configure(style, padding=(self.px(9), self.px(6)), bordercolor=c['line'],
                                 fieldbackground=c['panel'], foreground=c['fg'], insertcolor=c['fg'],
                                 arrowcolor=c['fg'], background=c['panel'], lightcolor=c['panel'], darkcolor=c['panel'],
                                 selectbackground=c['selected'], selectforeground=c['fg'])
            self.style.map(style, fieldbackground=[('readonly', c['panel'])], foreground=[('readonly', c['fg'])],
                           bordercolor=[('focus', c['accent'])], lightcolor=[('focus', c['accent'])])
        self.style.configure('Treeview', background=c['bg'], foreground=c['fg'], fieldbackground=c['bg'],
                             font=self.fonts['body'], rowheight=self.px(34 * self.factor), borderwidth=0)
        self.style.configure('Treeview.Heading', background=c['bg'], foreground=c['muted'],
                             padding=(self.px(7), self.px(7)), font=self.fonts['button'], relief='flat',
                             lightcolor=c['bg'], darkcolor=c['bg'], bordercolor=c['line'])
        self.style.map('Treeview.Heading', background=[('active', c['bg'])])
        self.style.map('Treeview', background=[('selected', c['selected'])], foreground=[('selected', c['fg'])])
        self.style.configure('TScrollbar', background=c['line'], troughcolor=c['bg'], arrowcolor=c['muted'],
                             bordercolor=c['bg'], arrowsize=self.px(12))
        self.root.option_add('*TCombobox*Listbox.background', c['panel'])
        self.root.option_add('*TCombobox*Listbox.foreground', c['fg'])
        self.root.option_add('*TCombobox*Listbox.selectBackground', c['selected'])
        self.root.option_add('*TCombobox*Listbox.selectForeground', c['fg'])
        for combo in (self.zoom_control, self.filter_control):
            try:
                popdown = combo.tk.call('ttk::combobox::PopdownWindow', str(combo))
                combo.tk.call(f'{popdown}.f.l', 'configure', '-background', c['panel'], '-foreground', c['fg'],
                              '-selectbackground', c['selected'], '-selectforeground', c['fg'],
                              '-font', str(self.fonts['body']))
            except tk.TclError:
                pass
        self.file_canvas.configure(bg=c['bg'])
        self.detail.configure(bg=c['bg'], fg=c['fg'], selectbackground=c['selected'], selectforeground=c['fg'])
        self.detail.tag_configure('accent', foreground=c['accent'], font=self.fonts['strong'])
        self.detail.tag_configure('strong', font=self.fonts['strong'])
        self.detail.tag_configure('small', foreground=c['muted'], font=self.fonts['small'])
        self.detail.tag_configure('latin', font=self.fonts['latin_data'])
        self.theme_button.configure(text=self.theme_action(), takefocus=True)
        self.menu_icon = self.make_icon('menu')
        self.theme_icon = self.make_icon('sun' if self.theme == 'dark' else 'moon')
        self.menu_button.configure(image=self.menu_icon, compound='image')
        self.theme_button.configure(image=self.theme_icon, compound='image')
        for dialog in (self.key_window, self.tooltip_window):
            if dialog is not None and dialog.winfo_exists():
                dialog.configure(bg=c['bg'])

    def theme_action(self):
        return 'Включить светлую тему' if self.theme == 'dark' else 'Включить тёмную тему'

    def toggle_theme(self):
        self.theme = 'dark' if self.theme == 'light' else 'light'
        self.hide_tooltip()
        self.apply_theme()
        self.save_preferences(theme=self.theme)

    def apply_zoom(self, remember=True):
        choice = self.zoom.get()
        self.factor = zoom_factor(choice, self.screen[0], self.dpi)
        for name, text_font in self.fonts.items():
            text_font.configure(size=-self.px(self.sizes[name] * self.factor))
        for column, _, width in COLUMNS:
            size = self.px(width * self.factor)
            self.table.column(column, width=size, minwidth=size, stretch=column == 'name',
                              anchor='e' if column in NUMERIC else 'w')
        self.apply_theme()
        if remember:
            self.save_preferences(zoom=choice)

    def build_header(self):
        header = ttk.Frame(self.root, padding=(self.px(24), self.px(12), self.px(24), self.px(8)))
        header.grid(row=0, column=0, sticky='ew')
        header.columnconfigure(1, weight=1)
        self.menu_button = ttk.Button(header, text='Меню', style='Icon.TButton', command=self.show_menu)
        self.menu_button.grid(row=0, column=0, padx=(0, self.px(10)))
        ttk.Label(header, text='Проверка заявок', style='Title.TLabel').grid(row=0, column=1, sticky='w')
        ttk.Button(header, text='DeepSeek', command=self.open_key, style='LatinQuiet.TButton').grid(
            row=0, column=2, padx=(self.px(7), self.px(12)))
        ttk.Label(header, text='Масштаб', style='Small.TLabel').grid(row=0, column=3)
        self.zoom_control = ttk.Combobox(header, textvariable=self.zoom, values=ZOOMS, state='readonly', width=8,
                                         font=self.fonts['button'])
        self.zoom_control.grid(row=0, column=4, padx=(self.px(7), self.px(12)))
        ttk.Button(header, text='Открыть журналы', command=self.open_logs, style='Quiet.TButton').grid(row=0, column=5)
        self.theme_button = ttk.Button(header, command=self.toggle_theme, style='Icon.TButton')
        self.theme_button.grid(row=0, column=6, padx=(self.px(8), 0))
        self.tooltip(self.theme_button, self.theme_action)
        self.tooltip(self.menu_button, lambda: 'Меню')

    def tooltip(self, widget, text):
        widget.bind('<Enter>', lambda _: self.queue_tooltip(widget, text), add='+')
        widget.bind('<Leave>', lambda _: self.hide_tooltip(), add='+')
        widget.bind('<Destroy>', lambda _: self.hide_tooltip(), add='+')

    def queue_tooltip(self, widget, text):
        self.hide_tooltip()
        self.tooltip_timer = self.root.after(550, lambda: self.show_tooltip(widget, text()))

    def show_tooltip(self, widget, text):
        self.tooltip_timer = None
        if not widget.winfo_exists():
            return
        self.tooltip_window = window = tk.Toplevel(self.root)
        window.overrideredirect(True)
        ttk.Label(window, text=text, style='Small.TLabel', padding=self.px(6)).pack()
        window.update_idletasks()
        x = min(widget.winfo_rootx(), self.screen[0] - window.winfo_reqwidth())
        window.geometry(f'+{max(0, x)}+{widget.winfo_rooty() + widget.winfo_height()}')

    def hide_tooltip(self):
        if self.tooltip_timer is not None:
            self.root.after_cancel(self.tooltip_timer)
            self.tooltip_timer = None
        if self.tooltip_window is not None:
            self.tooltip_window.destroy()
            self.tooltip_window = None

    def popup(self, button, items, event=None):
        self.hide_tooltip()
        c = THEMES[self.theme]
        menu = tk.Menu(self.root, tearoff=False, background=c['panel'], foreground=c['fg'], font=self.fonts['body'],
                       activebackground=c['selected'], activeforeground=c['fg'])
        for name, action in items:
            menu.add_command(label=name, command=action)
        x = event.x_root if event else button.winfo_rootx()
        y = event.y_root if event else button.winfo_rooty() + button.winfo_height()
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()
        menu.bind('<Unmap>', lambda _: self.root.after_idle(menu.destroy), add='+')

    def show_menu(self):
        self.popup(self.menu_button, [('Добавить заявки…', self.add_requests),
                                    ('Заменить «Закупки»…', lambda: self.choose_source('purchases')),
                                    ('Заменить реестр…', lambda: self.choose_source('registry')),
                                    ('Проверить заявки', self.manual_run)])

    def build_files(self):
        files = ttk.Frame(self.root, padding=(self.px(24), 0, self.px(24), self.px(10)))
        files.grid(row=1, column=0, sticky='ew')
        files.columnconfigure(0, weight=1)
        self.file_canvas = tk.Canvas(files, highlightthickness=0, borderwidth=0)
        self.file_canvas.grid(row=0, column=0, rowspan=2, sticky='ew')
        self.file_strip = ttk.Frame(self.file_canvas)
        self.file_canvas.create_window(0, 0, window=self.file_strip, anchor='nw')
        self.file_strip.bind('<Configure>', lambda _: self.resize_file_strip())
        self.file_canvas.bind('<Configure>', lambda _: self.resize_file_strip())
        self.file_scroll = ttk.Scrollbar(files, orient='horizontal', command=self.file_canvas.xview)
        self.file_scroll.grid(row=2, column=0, sticky='ew')
        self.file_canvas.configure(xscrollcommand=self.file_scroll.set)
        self.add_button = ttk.Button(files, text='+', style='Icon.TButton', command=self.add_requests)
        self.add_button.grid(row=0, column=1, padx=(self.px(6), self.px(12)))
        self.tooltip(self.add_button, lambda: 'Добавить заявки Excel')
        self.source_buttons = {}
        for col, role in enumerate(('purchases', 'registry'), 2):
            button = ttk.Button(files, width=20, style='File.TButton', command=lambda r=role: self.choose_source(r))
            button.grid(row=0, column=col, padx=(self.px(8), 0))
            button.bind('<Button-3>', lambda event, r=role: self.source_menu(r, event))
            self.tooltip(button, lambda r=role: str(self.sources[r] or self.role_label(r)))
            self.source_buttons[role] = button

    def resize_file_strip(self):
        width = self.file_strip.winfo_reqwidth()
        self.file_canvas.configure(height=max(self.px(34), self.file_strip.winfo_reqheight()), scrollregion=(0, 0, width, 1))
        if width > self.file_canvas.winfo_width():
            self.file_scroll.grid()
        else:
            self.file_scroll.grid_remove()

    @staticmethod
    def role_label(role):
        return 'Закупки' if role == 'purchases' else 'Реестр'

    def refresh_files(self):
        for child in self.file_strip.winfo_children():
            child.destroy()
        if not self.requests:
            ttk.Button(self.file_strip, text='Добавить заявки Excel', style='File.TButton',
                       command=self.add_requests).pack(side='left')
        for path in self.requests:
            text = path.name if len(path.name) <= 38 else f'{path.name[:32]}…{path.suffix}'
            button = ttk.Button(self.file_strip, text=text, style='File.TButton')
            button.configure(command=lambda p=path, b=button: self.request_menu(b, p))
            button.bind('<Button-3>', lambda event, p=path, b=button: self.request_menu(b, p, event))
            button.pack(side='left', padx=(0, self.px(8)))
            self.tooltip(button, lambda p=path: str(p))
        for role, button in self.source_buttons.items():
            path = self.sources[role]
            text = path.name if path else self.role_label(role)
            button.configure(text=text if len(text) <= 23 else f'{text[:17]}…{path.suffix}')
        self.root.after_idle(self.resize_file_strip)

    def add_requests(self):
        names = filedialog.askopenfilenames(parent=self.root, title='Выберите заявки Excel', filetypes=EXCEL)
        if names:
            existing = set(self.requests)
            self.requests.extend(path for path in dict.fromkeys(Path(name).resolve() for name in names) if path not in existing)
            self.refresh_files()
            self.inputs_changed()

    def request_menu(self, button, path, event=None):
        self.popup(button, [('Удалить заявку', lambda: self.remove_request(path))], event)

    def remove_request(self, path):
        if path in self.requests:
            self.requests.remove(path)
            self.refresh_files()
            self.inputs_changed()

    def choose_source(self, role):
        file = filedialog.askopenfilename(parent=self.root, title=self.role_label(role),
                                           filetypes=EXCEL if role == 'purchases' else REGISTER)
        if file:
            self.set_source(role, Path(file).resolve())

    def source_menu(self, role, event):
        items = [('Заменить файл…', lambda: self.choose_source(role))]
        if self.sources[role]:
            items.append(('Убрать файл', lambda: self.set_source(role, None)))
        self.popup(self.source_buttons[role], items, event)

    def set_source(self, role, path):
        self.sources[role] = path
        self.save_preferences(**{role: str(path) if path else ''})
        self.refresh_files()
        self.inputs_changed()

    def build_workspace(self):
        outer = ttk.Frame(self.root)
        outer.grid(row=2, column=0, sticky='nsew')
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)
        ttk.Frame(outer, height=self.px(1), style='Rule.TFrame').grid(row=0, column=0, sticky='ew')
        body = ttk.Frame(outer, padding=(self.px(24), self.px(16), self.px(24), self.px(12)))
        body.grid(row=1, column=0, sticky='nsew')
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=3, uniform='workspace')
        body.columnconfigure(2, weight=1, uniform='workspace')

        panel = ttk.Frame(body, padding=(0, 0, self.px(16), 0))
        panel.grid(row=0, column=0, sticky='nsew')
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(3, weight=1)
        header = ttk.Frame(panel)
        header.grid(row=0, column=0, sticky='ew')
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text='Общий результат', style='Section.TLabel').grid(row=0, column=0, sticky='nw')
        self.export_full = ttk.Button(header, text='Полный отчёт', command=self.export, state='disabled',
                                      style='Quiet.TButton')
        self.export_full.grid(row=0, column=1, padx=(self.px(8), self.px(8)))
        self.export_inventory = ttk.Button(header, text='Выгрузить наличие', state='disabled', style='Action.TButton',
                                           command=lambda: self.export(inventory_only=True))
        self.export_inventory.grid(row=0, column=2)
        ttk.Label(panel, textvariable=self.summary, style='Small.TLabel').grid(
            row=1, column=0, sticky='w', pady=(self.px(3), self.px(15)))
        toolbar = ttk.Frame(panel)
        toolbar.grid(row=2, column=0, sticky='ew', pady=(0, self.px(11)))
        toolbar.columnconfigure(1, weight=1)
        self.filter_control = ttk.Combobox(toolbar, textvariable=self.filter, values=FILTERS, state='readonly',
                                           width=22, font=self.fonts['button'])
        self.filter_control.grid(row=0, column=0)
        self.search = ttk.Entry(toolbar, textvariable=self.query, font=self.fonts['body'])
        self.search.grid(row=0, column=1, sticky='ew', padx=(self.px(10), 0))
        self.search_hint = ttk.Label(self.search, text='Поиск по полке, названию или артикулу', style='SearchHint.TLabel')
        self.search_hint.bind('<Button-1>', lambda _: self.search.focus_set())
        self.search.bind('<FocusIn>', self.update_search_hint)
        self.search.bind('<FocusOut>', self.update_search_hint)
        self.query.trace_add('write', self.update_search_hint)
        self.update_search_hint()
        self.build_table(panel)

        ttk.Frame(body, width=self.px(1), style='Rule.TFrame').grid(row=0, column=1, sticky='ns')
        details = ttk.Frame(body, padding=(self.px(16), 0, 0, 0))
        details.grid(row=0, column=2, sticky='nsew')
        details.columnconfigure(0, weight=1)
        details.rowconfigure(2, weight=1)
        title = ttk.Label(details, text='Данные по позиции', style='Section.TLabel')
        title.grid(row=0, column=0, columnspan=2, sticky='nw')
        header.bind('<Configure>', lambda _: details.rowconfigure(0, minsize=header.winfo_height()))
        title.bind('<Configure>', lambda _: title.configure(wraplength=details.winfo_width()))
        ttk.Frame(details, height=self.px(1), style='Rule.TFrame').grid(
            row=1, column=0, columnspan=2, sticky='ew', pady=(self.px(9), self.px(8)))
        self.detail = tk.Text(details, width=1, height=1, wrap='word', font=self.fonts['body'], state='disabled',
                              borderwidth=0, highlightthickness=0, spacing1=self.px(2), spacing3=self.px(2))
        self.detail.grid(row=2, column=0, sticky='nsew')
        self.detail.bind('<Configure>', lambda event: self.detail.configure(tabs=(max(self.px(100), event.width - self.px(8)), 'right')))
        scroll = ttk.Scrollbar(details, orient='vertical', command=self.detail.yview)
        scroll.grid(row=2, column=1, sticky='ns')
        self.detail.configure(yscrollcommand=scroll.set)

    def update_search_hint(self, *_):
        if self.query.get() or self.root.focus_get() is self.search:
            self.search_hint.place_forget()
        else:
            self.search_hint.place(x=self.px(9), rely=.5, anchor='w', relwidth=.9)

    def build_table(self, parent):
        panel = ttk.Frame(parent)
        panel.grid(row=3, column=0, sticky='nsew')
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(0, weight=1)
        self.table = ttk.Treeview(panel, columns=[col for col, _, _ in COLUMNS], show='headings',
                                  selectmode='browse', takefocus=True)
        self.table.grid(row=0, column=0, sticky='nsew')
        for col, title, _ in COLUMNS:
            self.table.heading(col, text=title, anchor='e' if col in NUMERIC else 'w')
        scroll_y = ttk.Scrollbar(panel, command=self.table.yview)
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x = ttk.Scrollbar(panel, orient='horizontal', command=self.table.xview)
        scroll_x.grid(row=1, column=0, sticky='ew')
        self.table.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.table.bind('<<TreeviewSelect>>', self.show_details)
        self.empty = ttk.Label(panel, text='Выберите файлы для проверки.', style='Small.TLabel')
        self.empty.place(relx=.5, rely=.44, anchor='center')

    def open_key(self):
        if self.key_window is not None and self.key_window.winfo_exists():
            self.key_window.lift()
            self.key_entry.focus_set()
            return
        self.key_window = dialog = tk.Toplevel(self.root)
        dialog.title('Ключ DeepSeek')
        dialog.transient(self.root)
        dialog.configure(bg=THEMES[self.theme]['bg'])
        frame = ttk.Frame(dialog, padding=self.px(20))
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(0, weight=1)
        self.key_entry = ttk.Entry(frame, textvariable=self.key, show='•', font=self.fonts['body'], width=38)
        self.key_entry.grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0, self.px(9)))
        self.key_entry.bind('<Control-v>', self.paste_key)
        self.key_entry.bind('<Control-V>', self.paste_key)
        self.key_entry.bind('<Control-a>', self.select_key)
        self.key_entry.bind('<Button-3>', lambda event: self.popup(self.key_entry, [
            ('Вставить', self.paste_key), ('Выделить всё', self.select_key)], event))
        ttk.Button(frame, text='Вставить', command=self.paste_key).grid(row=1, column=0, sticky='ew')
        ttk.Button(frame, text='Запомнить ключ', command=self.remember_key).grid(row=1, column=1, padx=self.px(8))
        ttk.Button(frame, text='Готово', command=self.close_key).grid(row=1, column=2)
        ttk.Label(frame, textvariable=self.key_note, style='Small.TLabel').grid(
            row=2, column=0, columnspan=3, sticky='w', pady=(self.px(8), 0))
        self.initial_key = self.key.get().strip()
        dialog.protocol('WM_DELETE_WINDOW', self.close_key)
        dialog.bind('<Escape>', lambda _: self.close_key())
        self.key_entry.focus_set()

    def close_key(self):
        changed = self.initial_key != self.key.get().strip()
        self.key_window.destroy()
        self.key_window = self.key_entry = None
        if changed and not self.missing():
            self.inputs_changed()

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

    def missing(self):
        files = []
        if not self.requests:
            files.append('заявку Excel')
        if not self.sources['purchases']:
            files.append('«Закупки»')
        if not self.sources['registry']:
            files.append('реестр')
        if files:
            return f'Добавьте {", ".join(files)}.'
        paths = [*self.requests, *self.sources.values()]
        if len(set(paths)) != len(paths):
            return 'Укажите отдельные файлы для заявок, «Закупок» и реестра.'
        return ''

    def manual_run(self, _event=None):
        self.inputs_changed()
        return 'break'

    def inputs_changed(self):
        self.revision += 1
        if self.cancellation is not None:
            self.cancellation.set()
        if self.timer is not None:
            self.root.after_cancel(self.timer)
            self.timer = None
        self.report = None
        self.input_files = ()
        self.summary.set('—')
        self.table.delete(*self.table.get_children())
        self.clear_details()
        self.export_full.state(['disabled'])
        self.export_inventory.state(['disabled'])
        missing = self.missing()
        self.status.set(missing or f'Проверяю заявки: {len(self.requests)}…')
        self.empty.configure(text=missing or 'Читаю файлы и сверяю остатки…')
        self.empty.place(relx=.5, rely=.44, anchor='center')
        if not missing:
            self.timer = self.root.after(160, self.launch)

    def launch(self):
        self.timer = None
        if self.running is not None or self.missing():
            return
        paths = tuple(self.requests), self.sources['purchases'], self.sources['registry']
        self.running = self.revision
        self.cancellation = threading.Event()
        args = self.running, paths, self.key.get().strip(), self.cancellation.is_set
        threading.Thread(target=self.worker, args=args, daemon=True).start()
        self.root.after(100, self.finish)

    def worker(self, revision, paths, key, cancelled):
        report, error = None, ''
        try:
            report, _ = check_files(*paths, key, LOGS, cancel=cancelled)
        except CancelledError:
            pass
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        self.events.put((revision, paths, report, error))

    def finish(self):
        try:
            revision, paths, report, error = self.events.get_nowait()
        except queue.Empty:
            self.root.after(100, self.finish)
            return
        self.running = self.cancellation = None
        if revision != self.revision:
            if self.timer is None and not self.missing():
                self.launch()
            return
        if error:
            self.status.set(error)
            self.empty.configure(text='Замените файл, указанный в строке состояния.')
            return
        if report is None:
            return
        self.report = report
        self.input_files = (*paths[0], paths[1], paths[2])
        rows = report['rows']
        found = sum(row['stock_found'] for row in rows)
        counted = sum((row['stock_qty'] or 0) > 0 or (row['free_qty'] or 0) > 0 for row in rows)
        self.summary.set(f'Заявок: {report["request_count"]}    Позиций: {len(rows)}    Найдено: {found}    С количеством: {counted}')
        if self.filter.get() == FILTERS[0] and found:
            self.filter.set(FILTERS[1])
        else:
            self.fill_table()
        self.status.set(f'Проверено заявок: {report["request_count"]}.  {report["ai"]}')
        self.export_full.state(['!disabled'])
        self.export_inventory.state(['!disabled'])

    def fill_table(self):
        if self.report is None:
            return
        selected = self.table.selection()
        self.table.delete(*self.table.get_children())
        query = self.query.get().strip().casefold()
        code_query = article_key(query)
        view = self.filter.get()
        for row in self.report['rows']:
            if (view == FILTERS[1] and not row['stock_found']) or (view == FILTERS[2] and row['take'] is not None):
                continue
            article = row['article'] or row['matched_article']
            assigned = dict.fromkeys(item['place'] for item in row['locations'])
            shelf = ', '.join(assigned or row['stock_places'])
            codes = [article, *(line['article'] for line in row.get('request_lines', [])),
                     *(code for item in row['registry'] for code in item['articles'])]
            names = [line['name'] for line in row.get('request_lines', [])]
            searchable = ' '.join([row['name'], shelf, *codes, *names]).casefold()
            if query and query not in searchable and not (code_query and any(code_query in article_key(code) for code in codes)):
                continue
            values = (article, row['name'], unit_quantity(row['need'], row['unit']), inventory_text(row),
                      unit_quantity(row['free_qty'], row['free_unit']), shelf,
                      show_number(row['take']), show_number(row['buy']), row['status'])
            self.table.insert('', 'end', iid=str(row['row']), values=values)
        items = self.table.get_children()
        if items:
            self.empty.place_forget()
            target = selected[0] if selected and selected[0] in items else items[0]
            self.table.selection_set(target)
            self.table.focus(target)
            self.table.see(target)
            self.show_details()
        else:
            self.empty.configure(text='По этому фильтру позиций нет.')
            self.empty.place(relx=.5, rely=.44, anchor='center')
            self.clear_details()

    def clear_details(self):
        self.detail.configure(state='normal')
        self.detail.delete('1.0', 'end')
        self.detail.configure(state='disabled')

    def show_details(self, _event=None):
        selection = self.table.selection()
        if not selection or not self.report:
            return
        row = next(row for row in self.report['rows'] if row['row'] == int(selection[0]))
        self.detail.configure(state='normal')
        self.detail.delete('1.0', 'end')

        def line(text, tag=''):
            self.detail.insert('end', f'{text}\n', tag)

        line(row['name'] or row['article'], 'strong')
        code = row['article'] or row['matched_article']
        if code:
            line(code, 'latin' if code.isascii() else 'small')
        line('')
        for item in row.get('request_lines', []):
            line(f'{item["file"]}\t{unit_quantity(item["need"], item["unit"])}')
            line(f'{item["sheet"]}, строка {item["row"]}', 'small')
        line('')
        if row['stock_places']:
            self.detail.insert('end', 'Полка\t')
            line(locations_text(row) if len(row['locations']) > 1 else
                 ', '.join(item['place'] for item in row['locations']) or ', '.join(row['stock_places']), 'accent')
        line(f'Взять\t{unit_quantity(row["take"], row["unit"])}', 'strong')
        line(f'Докупить\t{unit_quantity(row["buy"], row["unit"])}', 'strong')
        line(row['reason'], 'small')
        for key, title in (('purchases', 'Закупки'), ('registry', 'Реестр'),
                           ('suggestions', 'Подсказки DeepSeek'), ('candidates', 'Кандидаты')):
            if row.get(key):
                line('')
                line(title, 'strong')
                for item in row[key]:
                    relation = {'same': 'Та же позиция: ', 'alternative': 'Аналог: '}.get(item.get('relation'), '')
                    line(relation + item_text(item))
        line('')
        for role in ('Закупки', 'Реестр'):
            info = self.report['files'][role]
            date = (info.get('exported') or info['modified'])[:16].replace('T', ' ')
            line(f'{role}: {date}', 'small')
        self.detail.configure(state='disabled')
        self.detail.see('1.0')

    def export(self, inventory_only=False):
        if self.report is None:
            return
        label = 'Выгрузить наличие' if inventory_only else 'Сохранить полный отчёт'
        suffix = 'наличие' if inventory_only else 'проверка'
        stem = self.input_files[0].stem if self.report['request_count'] == 1 else 'Склад Массивбург'
        name = filedialog.asksaveasfilename(parent=self.root, title=label, initialfile=f'{stem} — {suffix}.xlsx',
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
    load_windows_fonts()
    root = tk.Tk()
    App(root)
    root.mainloop()
