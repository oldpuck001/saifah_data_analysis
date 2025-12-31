# create_select_database_sqlite.py

import tkinter as tk
from tkinter import filedialog
from . import fill_area_text
from ..sql_tool_sqlite import sql_sqlite_tool_class

sql_qlite_tool_example = sql_sqlite_tool_class()

class create_select_database_sqlite_modular:

    def create_select_database_sqlite_frame(self, root, control_frame_config, text_area):
        frame_result = tk.Frame(root)
        frame_result.pack(side=tk.TOP, fill=tk.BOTH, padx=5, pady=(10, 5))

        tk.Label(frame_result, text=control_frame_config['button_name'], width=15, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        entry_widget = tk.Entry(frame_result, state='readonly', readonlybackground='white')         # 创建Entry并保存引用
        entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        tk.Button(frame_result, text=control_frame_config['button_name'], command=lambda e=entry_widget: self.create_select_database_sqlite_path(e,text_area), width=20).pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.entry_widget = entry_widget

        return frame_result

    def create_select_database_sqlite_path(self, entry_widget, text_area):
        fill_text = ''

        path = filedialog.asksaveasfilename(defaultextension='.sqlite',
                                            filetypes=[('SQLite Database Files', '*.sqlite'),
                                                       ('SQLite Database Files', '*.db'),
                                                       ('SQLite Database Files', '*.sqlite3'),
                                                       ('SQLite Database Files', '*.db3'),
                                                       ('All Files', '*.*')],
                                            confirmoverwrite=False)

        if path:
            fill_text += f'Selected: {path}\n'
            result_info = sql_qlite_tool_example.create_database_sqlite(path)
            if result_info[0]:
                fill_text += f'Database created/select successfully!\n'
            else:
                fill_text += 'Database creation/select failed!\n'
                fill_text += result_info[1]
        else:
            fill_text += f'Database creation/select failed!\n'

        entry_widget.config(state='normal')
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, path)
        entry_widget.config(state='readonly')

        fill_area_text.text_area_fill(text_area, fill_text)