"""Tk interface for mixed Office batch processing."""
from dataclasses import asdict
import json
import os
import queue
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import ppt_batch
import office_batch


def resizable_panes(parent):
    return tk.PanedWindow(parent, orient=tk.VERTICAL, background='#c7d9e3', borderwidth=0,
                          sashwidth=9, sashcursor='sb_v_double_arrow', showhandle=True,
                          handlesize=9, handlepad=16, opaqueresize=True)


class BatchView(ttk.Frame):
    def __init__(self, parent, *, add, remove, clear, destination, keyword_jobs, other_busy, update):
        super().__init__(parent, padding=(14, 12))
        self.destination = destination
        self.keyword_jobs = keyword_jobs
        self.other_busy = other_busy
        self.update = update
        self.busy = False
        self.events = queue.Queue()
        self.paths = []
        self.result = None
        self._drag_origin = None
        self._drag_paths = ()
        self._drag_y = None
        self._drop_index = None
        self._drag_timer = None
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        self.intro = ttk.Label(self, text='可拖入 Office 文件。按住清單中的一列上下拖曳可排序，分隔線可調整清單高度。', style='Hint.TLabel')
        self.intro.grid(row=0, column=0, sticky='w', pady=(0, 8))
        toolbar = ttk.Frame(self)
        toolbar.grid(row=1, column=0, sticky='ew', pady=(0, 7))
        self.buttons = []
        actions = [('加入檔案', add), ('清除清單', clear), ('移除選取檔案', lambda: remove(self.selected())),
                   ('上移', lambda: self.move(-1)), ('下移', lambda: self.move(1)), ('依檔名排序', self.sort)]
        for label, action in actions:
            button = ttk.Button(toolbar, text=label, command=action)
            button.pack(side='left', padx=(0, 7))
            self.buttons.append(button)
        self.panes = resizable_panes(self)
        self.panes.grid(row=2, column=0, sticky='nsew')
        table = ttk.Frame(self.panes)
        self.files = ttk.Treeview(table, columns=('name', 'format', 'path'), show='tree headings', height=4, selectmode='extended')
        self.files.heading('#0', text='順序')
        self.files.column('#0', width=66, minwidth=50, stretch=False, anchor='center')
        self.files.heading('name', text='文件')
        self.files.column('name', width=270, minwidth=120)
        self.files.heading('format', text='格式')
        self.files.column('format', width=90, stretch=False)
        self.files.heading('path', text='檔案位置')
        self.files.column('path', width=500, minwidth=100)
        self.files.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(table, orient='vertical', command=self.files.yview)
        scroll.pack(side='right', fill='y')
        self.files.configure(yscrollcommand=scroll.set)
        self.files.bind('<<TreeviewSelect>>', lambda event: self.update_controls())
        self._drop_line = tk.Frame(self.files, borderwidth=0, background='#096778')
        self.files.bind('<ButtonPress-1>', self._start_row_drag)
        self.files.bind('<B1-Motion>', self._move_row_drag)
        self.files.bind('<ButtonRelease-1>', self._finish_row_drag)
        self.files.bind('<Escape>', self._cancel_row_drag)
        self.files.bind('<Unmap>', self._cancel_row_drag)
        self.panes.add(table, minsize=56, stretch='never')

        lower = ttk.Frame(self.panes)
        lower.columnconfigure(0, weight=1)
        lower.rowconfigure(3, weight=1, minsize=38)
        self.panes.add(lower, stretch='always')

        options = ttk.LabelFrame(lower, text='這次要處理的項目', padding=(10, 5))
        options.grid(row=0, column=0, sticky='ew', pady=(7, 4))
        self.notes = tk.BooleanVar()
        self.hidden = tk.BooleanVar()
        self.comments = tk.BooleanVar()
        self.keywords = tk.BooleanVar()
        self.merge = tk.BooleanVar()
        self.word_comments = tk.BooleanVar()
        self.word_accept = tk.BooleanVar()
        self.word_merge = tk.BooleanVar()
        self.excel_comments = tk.BooleanVar()
        self.excel_notes = tk.BooleanVar()
        self.excel_merge = tk.BooleanVar()
        self.checks = []
        groups = [
            ('PowerPoint', [('刪除備忘稿', self.notes), ('刪除隱藏投影片', self.hidden),
                            ('刪除註解', self.comments), ('依清單順序合併', self.merge)]),
            ('Word', [('刪除註解', self.word_comments), ('接受所有修訂並停止追蹤', self.word_accept),
                      ('依清單順序合併', self.word_merge)]),
            ('Excel', [('刪除串列註解', self.excel_comments), ('刪除附註', self.excel_notes),
                       ('依清單順序合併工作表', self.excel_merge)]),
            ('共用', [('套用已勾選的關鍵字項目', self.keywords)]),
        ]
        self.format_rows = []
        for index, (family, choices) in enumerate(groups):
            ttk.Label(options, text=family).grid(row=index, column=0, sticky='w', padx=(0, 14))
            row = ttk.Frame(options)
            row.grid(row=index, column=1, sticky='ew', pady=2)
            self.format_rows.append(row)
            for label, var in choices:
                check = ttk.Checkbutton(row, text=label, variable=var, command=self.update_controls)
                check.pack(side='left', padx=(0, 11))
                self.checks.append(check)
        self.cleanup_options = [(self.notes, self.hidden, self.comments), (self.word_comments, self.word_accept),
                                (self.excel_comments, self.excel_notes)]
        all_cleanup = tuple(option for group in self.cleanup_options for option in group)
        self.cleanup_buttons = []
        for index, choices in enumerate([*self.cleanup_options, all_cleanup]):
            name = '全選 Office 清理' if index == 3 else '全選清理'
            button = ttk.Button(self.format_rows[index], text=name, command=lambda values=choices: self.toggle_cleanup(values))
            button.pack(side='left', padx=(0, 5))
            self.cleanup_buttons.append((button, choices, name))
        self.preview_button = ttk.Button(self.format_rows[2], text='預覽隱藏工作表', command=self.preview)
        self.preview_button.pack(side='left')
        self.keyword_hint = tk.StringVar()
        ttk.Label(lower, textvariable=self.keyword_hint, style='Hint.TLabel').grid(row=1, column=0, sticky='w', pady=(0, 4))
        controls = ttk.Frame(lower)
        controls.grid(row=2, column=0, sticky='ew', pady=(0, 6))
        self.run_button = ttk.Button(controls, text='執行並另存', style='Accent.TButton', command=self.run)
        self.run_button.pack(side='left', padx=(0, 8))
        self.open_button = ttk.Button(controls, text='開啟輸出資料夾', command=self.open_result)
        self.open_button.pack(side='left')
        self.status = tk.StringVar(value='清理會另存副本。接受修訂、合併、刪除隱藏頁和舊檔轉換需桌面版 Office。')
        self.status_label = ttk.Label(controls, textvariable=self.status, style='Hint.TLabel')
        self.status_label.pack(side='left', padx=12)
        log_box = ttk.Frame(lower)
        log_box.grid(row=3, column=0, sticky='nsew')
        self.log = tk.Text(log_box, height=1, wrap='word', state='disabled', background='#ffffff',
                           foreground='#193449', font=('Microsoft JhengHei', 10), relief='flat', padx=8, pady=5)
        self.log.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(log_box, orient='vertical', command=self.log.yview)
        scroll.pack(side='right', fill='y')
        self.log.configure(yscrollcommand=scroll.set)
        self.after_idle(lambda: self.panes.paneconfigure(lower, minsize=lower.winfo_reqheight()))

    def selected(self):
        return [self.paths[int(i)] for i in self.files.selection()]

    def select_all_cleanup(self):
        self.toggle_cleanup(self.cleanup_options[0])

    def toggle_cleanup(self, choices):
        value = not all(option.get() for option in choices)
        for option in choices:
            option.set(value)
        self.update_controls()

    def sync(self, files):
        supported = [path for path in files if Path(path).suffix.lower() in office_batch.FORMATS]
        if set(supported) != set(self.paths):
            self.paths = [path for path in self.paths if path in supported]
            self.paths.extend(path for path in supported if path not in self.paths)
            self.draw()
        self.update_controls()

    def draw(self, selected=()):
        self._cancel_row_drag()
        self.files.delete(*self.files.get_children())
        for index, path in enumerate(self.paths):
            label = 'even' if index % 2 == 0 else 'odd'
            self.files.insert('', 'end', iid=str(index), text=str(index + 1),
                              values=(Path(path).name, office_batch.family(path), str(Path(path).parent)), tags=(label,))
            if path in selected:
                self.files.selection_add(str(index))
        self.files.tag_configure('odd', background='#edf4f7')
        self.update_controls()

    def _start_row_drag(self, event):
        self._cancel_row_drag()
        # Shift/Ctrl clicks retain the tree's standard multi-selection behavior.
        if self.busy or self.other_busy() or len(self.paths) < 2 or event.state & 0x0005:
            return
        if self.files.identify_region(event.x, event.y) not in {'tree', 'cell'}:
            return
        row = self.files.identify_row(event.y)
        if not row:
            return
        anchor = self.paths[int(row)]
        selected = set(self.selected())
        self._drag_paths = tuple(path for path in self.paths if path in selected) if anchor in selected else (anchor,)
        self._drag_origin = event.y

    def _move_row_drag(self, event):
        if self._drag_origin is None:
            return
        if self._drag_y is None and abs(event.y - self._drag_origin) < 5:
            return
        self._drag_y = event.y
        self.files.configure(cursor='sb_v_double_arrow')
        self._show_drop_line()
        if self._drag_timer is None:
            self._drag_timer = self.after(65, self._scroll_drag)
        return 'break'

    def _first_visible_row(self):
        return min(len(self.paths) - 1, round(self.files.yview()[0] * len(self.paths)))

    def _show_drop_line(self):
        first = self._first_visible_row()
        first_box = self.files.bbox(str(first))
        self._drop_index = None
        if not first_box:
            self._drop_line.place_forget()
            return
        height = self.files.winfo_height()
        y = max(first_box[1], min(self._drag_y, height-2))
        row = self.files.identify_row(y)
        if row:
            _, top, _, size = self.files.bbox(row)
            after = y >= top + size / 2
            self._drop_index = int(row) + int(after)
            position = top + (size if after else 0)
        else:
            self._drop_index = len(self.paths)
            last_box = self.files.bbox(str(len(self.paths)-1))
            position = last_box[1]+last_box[3] if last_box else height-2
        self._drop_line.place(x=1, y=min(position, height-2), relwidth=1, width=-2, height=2)

    def _scroll_drag(self):
        self._drag_timer = None
        if self._drag_y is None:
            return
        first, last = self.files.yview()
        box = self.files.bbox(str(self._first_visible_row()))
        if box and self._drag_y < box[1]+12 and first > 0:
            self.files.yview_scroll(-1, 'units')
        elif self._drag_y > self.files.winfo_height()-12 and last < 1:
            self.files.yview_scroll(1, 'units')
        self._show_drop_line()
        self._drag_timer = self.after(65, self._scroll_drag)

    def _finish_row_drag(self, event):
        if self._drag_y is None:
            self._cancel_row_drag()
            return
        self._drag_y = event.y
        self._show_drop_line()
        boundary = self._drop_index
        moving = self._drag_paths
        self._cancel_row_drag()
        if boundary is None or not (0 <= event.x < self.files.winfo_width() and 0 <= event.y < self.files.winfo_height()):
            return 'break'
        selected = set(moving)
        remaining = [path for path in self.paths if path not in selected]
        insertion = sum(path not in selected for path in self.paths[:boundary])
        reordered = remaining[:insertion] + list(moving) + remaining[insertion:]
        if reordered != self.paths:
            first = self.files.yview()[0]
            self.paths[:] = reordered
            self.draw(selected)
            self.files.yview_moveto(first)
            self.files.focus(str(insertion))
        return 'break'

    def _cancel_row_drag(self, event=None):
        if self._drag_timer is not None:
            self.after_cancel(self._drag_timer)
            self._drag_timer = None
        if self._drag_y is not None:
            self.files.configure(cursor='')
        self._drag_origin = self._drag_y = self._drop_index = None
        self._drag_paths = ()
        self._drop_line.place_forget()

    def move(self, direction):
        selected = set(self.selected())
        indices = range(1, len(self.paths)) if direction == -1 else range(len(self.paths) - 2, -1, -1)
        for index in indices:
            other = index + direction
            if self.paths[index] in selected and self.paths[other] not in selected:
                self.paths[index], self.paths[other] = self.paths[other], self.paths[index]
        self.draw(selected)

    def sort(self):
        selected = self.selected()
        self.paths.sort(key=lambda path: [int(word) if word.isdecimal() else word.lower()
                                         for word in re.split(r'(\d+)', Path(path).name)])
        self.draw(selected)

    def update_controls(self):
        blocked = self.busy or self.other_busy()
        if blocked:
            self._cancel_row_drag()
        selected = self.files.selection()
        available = self.keyword_jobs(self.paths)
        self.keyword_hint.set(f'目前有 {sum(len(job.ids) for job in available.values())} 個已勾選的關鍵字項目。'
                              '可到「關鍵字掃描」確認或調整勾選。')
        for index, button in enumerate(self.buttons):
            enabled = bool(not blocked and (index < 2 or (selected if index < 5 else len(self.paths) > 1)))
            button.configure(state='normal' if enabled else 'disabled')
        for check in self.checks:
            check.configure(state='disabled' if blocked else 'normal')
        for button, choices, name in self.cleanup_buttons:
            button.configure(state='disabled' if blocked else 'normal',
                             text='清除清理勾選' if all(option.get() for option in choices) else name)
        can_preview = any(office_batch.family(path)=='Excel' for path in self.paths)
        self.preview_button.configure(state='normal' if not blocked and can_preview else 'disabled')
        tasks, _ = office_batch.plan(self.paths, self.actions(), available if self.keywords.get() else {})
        enabled = not blocked and bool(tasks)
        self.run_button.configure(state='normal' if enabled else 'disabled')
        self.open_button.configure(state='normal' if self.result and not blocked else 'disabled')

    def write_log(self, value, clear=False):
        self.log.configure(state='normal')
        if clear:
            self.log.delete('1.0', 'end')
        self.log.insert('end', value)
        self.log.see('end')
        self.log.configure(state='disabled')

    def open_result(self):
        try:
            os.startfile(str(self.result))
        except OSError as error:
            messagebox.showerror('無法開啟輸出資料夾', str(error), parent=self)

    def actions(self):
        return office_batch.Actions(ppt_batch.Options(self.notes.get(), self.hidden.get(), self.comments.get()),
                                    self.merge.get(), self.word_comments.get(), self.word_accept.get(),
                                    self.excel_comments.get(), self.excel_notes.get(),
                                    self.word_merge.get(), self.excel_merge.get())

    def preview(self):
        if self.busy or self.other_busy():
            return
        paths = [path for path in self.paths if office_batch.family(path)=='Excel']
        self.busy = True
        self.update()
        self.write_log('', clear=True)
        self.status.set('正在讀取 Excel 的隱藏工作表……')

        def worker():
            for path in paths:
                try:
                    sheets = office_batch.preview_excel(path)['sheets']
                    findings = ', '.join(f"{sheet['name']}（{'深度隱藏' if sheet['state']=='veryHidden' else '隱藏'}）"
                                         for sheet in sheets)
                    self.events.put(('log', f"{Path(path).name}：{findings or '沒有隱藏工作表'}\n"))
                except Exception as error:
                    self.events.put(('log', f'{Path(path).name}：無法預覽。{error}\n'))
            self.events.put(('preview_done', None))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self.poll)

    def run(self):
        if self.busy or self.other_busy():
            return
        actions = self.actions()
        keywords = self.keyword_jobs(self.paths) if self.keywords.get() else {}
        tasks, skipped = office_batch.plan(self.paths, actions, keywords)
        if not tasks:
            return
        target = self.destination('Office_Batch')
        if not target:
            return
        self.result = None
        self.busy = True
        self.update()
        self.write_log('', clear=True)
        self.status.set('正在處理 Office 文件……')

        def worker():
            results = []
            for kind, enabled in office_batch.merge_options(actions).items():
                if enabled and not any(task.merge and task.kind == kind for task in tasks):
                    self.events.put(('log', f'{kind} 合併至少需要兩份檔案。其他已勾選的項目照常處理。\n'))
            for index, task in enumerate(tasks, 1):
                try:
                    result = office_batch.process(task, target if task.merge else target / f'{index:03d}', actions, keywords)
                    result.update(status='success', format=task.kind)
                    results.append(result)
                    self.events.put(('log', f"完成 {Path(result['output']).name}\n"))
                except Exception as error:
                    results.append({'source': list(task.paths), 'status': 'failed', 'format': task.kind, 'error': str(error)})
                    self.events.put(('log', f"未輸出 {', '.join(Path(p).name for p in task.paths)}：{error}\n"))
            try:
                target.mkdir(parents=True, exist_ok=True)
                with (target / 'batch.report.json').open('x', encoding='utf-8') as stream:
                    json.dump({'operations': asdict(actions), 'results': results + skipped}, stream, ensure_ascii=False, indent=2)
            except OSError as error:
                self.events.put(('log', f'未能寫入批次報告：{error}\n'))
            success = sum('output' in item for item in results)
            self.events.put(('done', (success, len(tasks) - success, len(skipped), target)))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self.poll)

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == 'log':
                    self.write_log(value)
                elif kind == 'preview_done':
                    self.busy = False
                    self.status.set('Excel 隱藏工作表預覽完成。')
                    self.update()
                else:
                    success, errors, skipped, target = value
                    self.result = target if target.is_dir() else None
                    self.busy = False
                    self.status.set(f'處理完成。成功 {success}、失敗 {errors}、跳過 {skipped}。')
                    self.update()
        except queue.Empty:
            if self.busy:
                self.after(100, self.poll)
