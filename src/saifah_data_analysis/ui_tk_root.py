# ui_tk_root.py

import sys
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from .ui_tk_modular import create_select_database_sqlite
from .ui_tk_modular import data_import_clean_file
from .ui_tk_modular import sql_sqlite_win_modular

create_select_database_sqlite_modular_class = create_select_database_sqlite.create_select_database_sqlite_modular()
data_import_clean_file_modular_class = data_import_clean_file.data_import_clean_file_modular()
sql_sqlite_win_modular_class = sql_sqlite_win_modular.sql_sqlite_win_modular_class()

class App:
    def __init__(self, title='My Application', geometry='1024x768+140+130', minsize_x=640, minsize_y=360, maxsize_x=1920, maxsize_y=1080,
                    resizable_x=True, resizable_y=True, control_frame_n=0, control_frame_config=[]):

        self.root = tk.Tk()                                             # 创建tk实例

        self.root.title(title)                                          # 设置窗口标题

        self.root.geometry(geometry)                                    # 设置窗口的大小和位置

        self.root.minsize(minsize_x, minsize_y)                         # 设置窗口的最小大小

        self.root.maxsize(maxsize_x, maxsize_y)                         # 设置窗口的最大大小

        self.root.resizable(resizable_x, resizable_y)                   # 设置窗口是否可以调整大小

        if sys.platform == 'darwin':
            self.root.after(200, self.bring_to_front)                   # macOS workaround: mainloop開始後再將視窗浮前

        self.control_frame_n =control_frame_n                           # 自定义UI模块个数

        self.control_frame_config = control_frame_config                # 自定义UI模块参数                                           

        self.root.control_frame_list = []                               # 自定义模块框架列表

        # 操作记录区
        self.frame_text_area = tk.Frame(self.root)
        self.frame_text_area.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, padx=5, pady=5)
        tk.Label(self.frame_text_area, text='Operation Log', anchor='w').pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        try:
            height_value = self.control_frame_config[self.control_frame_n]['text_area_hight']
            self.text_area = ScrolledText(self.frame_text_area, height=height_value)
        except:
            self.text_area = ScrolledText(self.frame_text_area)
        self.text_area.pack(side=tk.TOP, expand=True, fill=tk.X)
        self.text_area.config(state='disabled')

        # 选择选取控件区
        for n in range(control_frame_n):

            if self.control_frame_config[n]['name'] in ['create_select_sqlite_database']:
                self.root.control_frame_list.append(create_select_database_sqlite_modular_class.
                                                create_select_database_sqlite_frame(root=self.root,
                                                                                    control_frame_config=self.control_frame_config[n],
                                                                                    text_area=self.text_area))
            elif self.control_frame_config[n]['name'] in ['data_import_clean_file']:
                self.root.control_frame_list.append(data_import_clean_file_modular_class.
                                                    data_import_clean_file_frame(root=self.root,
                                                                                 control_frame_config=self.control_frame_config[n],
                                                                                 text_area=self.text_area))
            elif self.control_frame_config[n]['name'] in ['sql_sqlite_win']:
                self.root.control_frame_list.append(sql_sqlite_win_modular_class.
                                                    sql_sqlite_win_modular_frame(root=self.root,
                                                                                 control_frame_config=self.control_frame_config[n],
                                                                                 text_area=self.text_area))

    def bring_to_front(self):
        self.root.lift()
        self.root.focus_force()
        self.root.call('wm', 'attributes', '.', '-topmost', '1')
        self.root.call('wm', 'attributes', '.', '-topmost', '0')