# data_import_clean_file.py

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from . import fill_area_text
from ..pd_DataFrame_tool import pd_DataFrame_tool_class
from ..sql_qlite_tool import sql_sqlite_tool_class

pd_DataFrame_tool_example = pd_DataFrame_tool_class()
sql_table_import = sql_sqlite_tool_class()

class data_import_clean_file_modular:

    data_df = pd.DataFrame()

    def data_import_clean_file_frame(self, root, control_frame_config, text_area):
        frame_result = tk.Frame(root)
        frame_result.pack(side=tk.TOP, fill=tk.BOTH)

        # 地址栏行
        frame_1 = tk.Frame(frame_result)
        frame_1.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_1, text=control_frame_config['button_name'], width=15, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        entry_widget = tk.Entry(frame_1, state='readonly', readonlybackground='white')                                    # 创建Entry并保存引用
        entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        tk.Button(frame_1, text=control_frame_config['button_name'], command=lambda e=entry_widget: self.select_open_file_path(e,text_area), width=20).pack(side=tk.LEFT, padx=5, pady=5)
        frame_1.entry_widget = entry_widget
        frame_result.frame_1 = frame_1

        # 选择工作表与按钮列
        frame_2 = tk.Frame(frame_result)
        frame_2.pack(side=tk.TOP, fill=tk.BOTH, padx=5, pady=5)
        tk.Label(frame_2, text='Excel sheet', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options = []                                                                                        # 创建选项列表
        frame_2.combobox = ttk.Combobox(frame_2, values=options, height=10, width=55, state='readonly')     # 使用 ttk.Combobox（现代下拉选框）
        frame_2.combobox.pack(side=tk.LEFT, padx=(5,40), pady=5)
        frame_result.frame_2 = frame_2                                                                      # 将frame_1保存到control_frame_list[n]中，以便后续访问

        tk.Button(frame_2, text='读取工作表清单', command=lambda: self.button_sheets_fun(frame_1, frame_2, text_area), width=27).pack(side=tk.LEFT, padx=(5, 40), pady=5)
        tk.Button(frame_2, text='读取数据', command=lambda: self.button_load_fun(frame_1, frame_2, frame_3, frame_4, frame_5, frame_6, frame_7, text_area), width=27).pack(side=tk.LEFT, padx=(5, 5), pady=5)

        # 读取文件参数区

        frame_3 = tk.Frame(frame_result)
        frame_3.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_3, text='skiprows', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.entry_skiprows = tk.Entry(frame_3, width=20)
        frame_3.entry_skiprows.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_3.entry_skiprows.insert(0, 'None')

        tk.Label(frame_3, text='usecols', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.entry_usecols = tk.Entry(frame_3, width=20)
        frame_3.entry_usecols.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_3.entry_usecols.insert(0, 'None')

        tk.Label(frame_3, text='nrows', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.entry_nrows = tk.Entry(frame_3, width=20)
        frame_3.entry_nrows.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_3.entry_nrows.insert(0, 'None')

        tk.Label(frame_3, text='index_col', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.entry_index_col = tk.Entry(frame_3, width=20)
        frame_3.entry_index_col.pack(side=tk.LEFT, padx=(5, 5), pady=5)
        frame_3.entry_index_col.insert(0, 'None')

        frame_result.frame_3 = frame_3

        frame_4 = tk.Frame(frame_result)
        frame_4.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_4, text='header', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_4.entry_header = tk.Entry(frame_4, width=20)
        frame_4.entry_header.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_4.entry_header.insert(0, 0)

        tk.Label(frame_4, text='names', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_4.entry_names = tk.Entry(frame_4, width=20)
        frame_4.entry_names.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_4.entry_names.insert(0, 'None')

        tk.Label(frame_4, text='na_values', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_4.entry_na_values = tk.Entry(frame_4, width=20)
        frame_4.entry_na_values.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_4.entry_na_values.insert(0, 'None')

        tk.Label(frame_4, text='keep_default_na', width=12, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options = [True, False]
        frame_4.combobox_keep_default_na = ttk.Combobox(frame_4, values=options, height=2, width=15, state='readonly')
        frame_4.combobox_keep_default_na.set('True')
        frame_4.combobox_keep_default_na.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        frame_result.frame_4 = frame_4

        frame_5 = tk.Frame(frame_result)
        frame_5.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_5, text='dtype', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.entry_dtype = tk.Entry(frame_5, width=56)
        frame_5.entry_dtype.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_5.entry_dtype.insert(0, 'None')

        tk.Label(frame_5, text='sep', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.entry_sep = tk.Entry(frame_5, width=20)
        frame_5.entry_sep.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_5.entry_sep.insert(0, ',')

        tk.Label(frame_5, text='encoding', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.entry_encoding = tk.Entry(frame_5, width=20)
        frame_5.entry_encoding.pack(side=tk.LEFT, padx=(5, 5), pady=5)
        frame_5.entry_encoding.insert(0, 'utf-8')

        frame_result.frame_5 = frame_5

        frame_6 = tk.Frame(frame_result)
        frame_6.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_6, text='converters', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_6.entry_converters = tk.Entry(frame_6, width=92)
        frame_6.entry_converters.pack(side=tk.LEFT, padx=(5, 40), pady=5)
        frame_6.entry_converters.insert(0, 'None')

        tk.Label(frame_6, text='engine_csv', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_6.entry_engine_csv = tk.Entry(frame_6, width=20)
        frame_6.entry_engine_csv.pack(side=tk.LEFT, padx=(5, 5), pady=5)
        frame_6.entry_engine_csv.insert(0, 'c')

        frame_result.frame_6 = frame_6

        frame_7 = tk.Frame(frame_result)
        frame_7.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_7, text='DataFrame column', width=17, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_col = []
        frame_7.combobox_col = ttk.Combobox(frame_7, values=options_col, height=10, width=33, state='readonly')
        frame_7.combobox_col.pack(side=tk.LEFT, padx=(5, 40), pady=5)

        tk.Label(frame_7, text='数据清洗选项', width=12, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_clean = ['删除指定列的重复行',
                         '填充缺失值为 0',
                         '填充缺失值为 <空白>',
                         '填充缺失值为 重复上一行',
                         '标准化文本（去除首尾空格及转换为小写英文字母）',
                         '将数据类型转换为字符型',
                         '转换为整数，支持缺失值',
                         '转换为浮点数，支持缺失值',
                         '将数据类型转换为时间日期类型']
        frame_7.combobox_clean = ttk.Combobox(frame_7, values=options_clean, height=10, width=38, state='readonly')
        frame_7.combobox_clean.set(options_clean[7])
        frame_7.combobox_clean.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Button(frame_7, text='数据预览', command=lambda: self.button_clean_fun(self.data_df), width=20).pack(side=tk.LEFT, padx=(5, 5), pady=5)
        # tk.Button(frame_2, text='数据清洗', command=lambda n=n: self.button_clean_fun(n), width=20).pack(side=tk.LEFT, padx=(5, 20), pady=5)


        frame_result.frame_7 = 7
        '''
        frame_7 = tk.Frame(self.control_frame_list[n])
        frame_7.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_7, text='SQL table name', width=14, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_7.entry_sql_table_name = tk.Entry(frame_7, width=20)
        frame_7.entry_sql_table_name.pack(side=tk.LEFT, padx=(5, 25), pady=5)

        tk.Label(frame_7, text='if_exists', width=7, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_sql_if_exists = ['fail', 'replace', 'append']
        frame_7.combobox_sql_if_exists = ttk.Combobox(frame_7, values=options_sql_if_exists, height=10, width=17, state='readonly')
        frame_7.combobox_sql_if_exists.set(options_sql_if_exists[1])
        frame_7.combobox_sql_if_exists.pack(side=tk.LEFT, padx=(5, 25), pady=5)

        tk.Label(frame_7, text='index', width=6, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_sql_index = [True, False]
        frame_7.combobox_sql_index = ttk.Combobox(frame_7, values=options_sql_index, height=10, width=17, state='readonly')
        frame_7.combobox_sql_index.set('False')
        frame_7.combobox_sql_index.pack(side=tk.LEFT, padx=(5, 25), pady=5)

        tk.Label(frame_7, text='index_label', width=11, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_7.entry_sql_index_label = tk.Entry(frame_7, width=20)
        frame_7.entry_sql_index_label.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        # tk.Button(frame_2, text='导出至.xlsx', command=self.button_export_fun, width=20).pack(side=tk.LEFT, padx=(5, 20), pady=5)
        # tk.Button(frame_1, text='导入数据库', command=lambda n=n: self.button_sql_fun(n), width=20).pack(side=tk.LEFT, padx=(5, 5), pady=5)

        self.control_frame_list[n].frame_7 = frame_7
        '''
        return frame_result


    # 选择文件按钮函数
    def select_open_file_path(self, entry_widget, text_area):
        fill_text = ''

        path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx'),
                                                     ('Excel Files', '*.xls'),
                                                     ('CSV Files', '*.csv'),
                                                     ('Text Files', '*.txt'),
                                                     ('All Files', '*.*')])

        if path:
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            entry_widget.config(state='readonly')
            fill_text += f'Selected: {path}\n'
            fill_text += 'File selection successful!\n'
        else:
            fill_text += f'File selection failed!\n'

        entry_widget.config(state='normal')
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, path)
        entry_widget.config(state='readonly')

        fill_area_text.text_area_fill(text_area, fill_text)


    # 读取工作表列表按钮函数
    def button_sheets_fun(self, frame_path, frame_sheets, text_area):

        path = frame_path.entry_widget.get()

        if path:
            result_info = pd_DataFrame_tool_example.sheetnames_import(path)
            if result_info[0]:
                frame_sheets.combobox['values'] = result_info[2]
                frame_sheets.combobox.set(result_info[2][0])                                     # 设置默认选择第一个
                frame_sheets.combobox.config(state='readonly')
            else:
                frame_sheets.combobox.set('')
                frame_sheets.combobox.config(state='disabled')

        fill_text = result_info[1]
        fill_area_text.text_area_fill(text_area, fill_text)


    # 读取数据按钮函数
    def button_load_fun(self, frame_path, frame_sheets, frame_3, frame_4, frame_5, frame_6, frame_7, text_area):

        path = frame_path.entry_widget.get()
        sheet_name = frame_sheets.combobox.get()
        skiprows = frame_3.entry_skiprows.get()
        usecols = frame_3.entry_usecols.get()
        nrows = frame_3.entry_nrows.get()
        index_col = frame_3.entry_index_col.get()
        header = frame_4.entry_header.get()
        names = frame_4.entry_names.get()
        na_values = frame_4.entry_na_values.get()
        keep_default_na = frame_4.combobox_keep_default_na.get()
        dtype = frame_5.entry_dtype.get()
        sep = frame_5.entry_sep.get()
        encoding = frame_5.entry_encoding.get()
        converters = frame_6.entry_converters.get()
        engine_csv = frame_6.entry_engine_csv.get()

        if path:
            result_info = pd_DataFrame_tool_example.read_xlsx_xls_csv_txt_fun(original_file_path=path,
                                                                              sheet_name=sheet_name,
                                                                              skiprows=skiprows,
                                                                              usecols=usecols,
                                                                              nrows=nrows,
                                                                              index_col=index_col,
                                                                              header=header,
                                                                              names=names,
                                                                              na_values=na_values,
                                                                              keep_default_na=keep_default_na,
                                                                              dtype=dtype,
                                                                              converters=converters,
                                                                              sep=sep,
                                                                              encoding=encoding,
                                                                              engine_csv=engine_csv)
            if result_info[0]:
                self.data_df = result_info[2]
                option_clean_cols = self.data_df.columns.tolist()
                frame_7.combobox_col['values'] = option_clean_cols
                frame_7.combobox_col.set(option_clean_cols[0])
                frame_7.combobox_col.config(state='readonly')
            fill_text = result_info[1]
        else:
            fill_text = f'Please select a file first!\n'

        fill_area_text.text_area_fill(text_area, fill_text)

    # 数据预览
    def button_clean_fun(self, df, frame_path, text_area):

        path = frame_path.entry_widget.get()
        

        print(df)


'''
    def button_export_fun(self):

        path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])

        if path:
            pd_DataFrame_tool_example = pd_DataFrame_tool.pd_DataFrame_tool_class()
            result_info = pd_DataFrame_tool_example.df_export_xlsx(self.data_df_class, path)

        fill_text = result_info[1]
        self.text_area_fill(fill_text)

    def convert_none(self, value):
        if str(value).strip().lower() == 'none':
            return None
        else:
            try:
                return int(value)
            except:
                return value 

    def button_clean_fun(self, n):

        clean_col = self.control_frame_list[n].frame_6.combobox_col.get()
        clean_option = self.control_frame_list[n].frame_6.combobox_clean.get()

        pd_DataFrame_tool_example = pd_DataFrame_tool.pd_DataFrame_tool_class()
        result_info = pd_DataFrame_tool_example.df_cleaning(self.data_df_class, clean_col, clean_option)

        if result_info[0]:
            self.data_df_class = result_info[2]

        fill_text = result_info[1]
        self.text_area_fill(fill_text)

    def button_sql_fun(self, n):
        fill_text = ''

        n_config = self.control_frame_config[3][n][0]
        sql_path = self.control_frame_list[n_config].entry_widget.get()
        table_name = self.control_frame_list[n].frame_7.entry_sql_table_name.get()
        if_exists = self.control_frame_list[n].frame_7.combobox_sql_if_exists.get()
        index = self.control_frame_list[n].frame_7.combobox_sql_index.get()
        index_label = self.control_frame_list[n].frame_7.entry_sql_index_label.get()

        if sql_path and table_name:

            result_info = sql_table_import.import_dataframe(sql_path, self.data_df_class, table_name, if_exists, index, index_label)
            if result_info[0]:
                fill_text += f'Database import successful!\n'
            else:
                fill_text += 'Database import failed!\n'
                fill_text += result_info[1]
        else:
            if not sql_path:
                fill_text += f'Please check that the database path is not empty!\n'
            if not table_name:
                fill_text += f'Please check that the database table name is not empty!\n'

        self.text_area_fill(fill_text)
'''