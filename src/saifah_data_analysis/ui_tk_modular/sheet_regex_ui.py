# sheet_regex_ui.py

import os
import re
import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk
import subprocess
import pandas as pd
import operator
from . import fill_area_text
from ..DataFrame_tools_pd import sheetnames_import_fun
from ..DataFrame_tools_pd import columns_title_fun
from ..DataFrame_tools_pd import read_xlsx_xls_csv_txt_fun
from ..DataFrame_tools_pd import df_export_xlsx_fun


class sheet_regex_ui_class:

    regex_list = []
    left_right = False
    data_df = pd.DataFrame()

    def sheet_regex_ui_frame(self, root, control_frame_config, text_area):

        frame_result = tk.Frame(root)
        frame_result.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        frame_result.frame_1 = tk.Frame(frame_result)
        frame_result.frame_1.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_result.frame_1, text='File Path', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_1.entry_widget = tk.Entry(frame_result.frame_1, state='readonly', readonlybackground='white')         # 创建Entry并保存引用
        frame_result.frame_1.entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

        options_banlk = []
        options_label = ['', '~']

        frame_result.frame_2 = tk.Frame(frame_result)
        frame_result.frame_2.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_result.frame_2, text='选择工作表', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_2.combobox_sheet = ttk.Combobox(frame_result.frame_2, values=options_banlk, state='readonly')
        frame_result.frame_2.combobox_sheet.pack(side=tk.LEFT, padx=5, pady=5)
        tk.Label(frame_result.frame_2, text='选择筛选列', width=8, anchor='w').pack(side=tk.LEFT, padx=(50, 5), pady=5)
        frame_result.frame_2.combobox_col = ttk.Combobox(frame_result.frame_2, values=options_banlk, state='readonly')
        frame_result.frame_2.combobox_col.pack(side=tk.LEFT, padx=5, pady=5)
        tk.Label(frame_result.frame_2, text='选择标识符', width=8, anchor='w').pack(side=tk.LEFT, padx=(50, 5), pady=5)
        frame_result.frame_2.combobox_label = ttk.Combobox(frame_result.frame_2, values=options_label, state='readonly')
        frame_result.frame_2.combobox_label.pack(side=tk.LEFT, padx=5, pady=5)

        frame_result.frame_3 = tk.Frame(frame_result)
        frame_result.frame_3.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_result.frame_3, text='正则表达式（Python环境）', width=17, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_3.entry_widget = tk.Entry(frame_result.frame_3)
        frame_result.frame_3.entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

        # 正则表达式指令区
        frame_result.frame_4 = tk.Frame(frame_result)
        frame_result.frame_4.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        tk.Label(frame_result.frame_4, text='正则表达式指令区', anchor='w').pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        frame_result.frame_4.regex_command_text_area = ScrolledText(frame_result.frame_4, height=15)
        frame_result.frame_4.regex_command_text_area.pack(side=tk.TOP, expand=True, fill=tk.BOTH)

        # 按钮行
        frame_result.frame_5 = tk.Frame(frame_result)
        frame_result.frame_5.pack(side=tk.TOP, fill=tk.BOTH)
        # 选择文件按钮
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][0],
                  command=lambda e=frame_result.frame_1.entry_widget,
                                 f=frame_result.frame_2.combobox_sheet,
                                 g=frame_result.frame_2.combobox_col: self.select_sheet(e,f,g,text_area),
                  width=10).pack(side=tk.LEFT, padx=5)
        # 导入数据按钮
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][1],
                  command=lambda e=frame_result.frame_1.entry_widget,
                                 f=frame_result.frame_2.combobox_sheet: self.import_sheet(e,f,text_area),
                  width=10).pack(side=tk.LEFT, padx=5)
        # 预览数据按钮
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][2],
                  command=lambda e=frame_result.frame_1.entry_widget: self.button_preview_fun(e,text_area),
                  width=10).pack(side=tk.LEFT, padx=5)
        # button_preview_fun
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][3],
                  command=lambda e=frame_result.frame_2.combobox_col,
                                 f=frame_result.frame_2.combobox_label,
                                 g=frame_result.frame_3.entry_widget,
                                 h=frame_result.frame_4.regex_command_text_area: self.add_regex(e,f,g,h),
                  width=10).pack(side=tk.LEFT, padx=5)
        
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][4],
                  command=lambda e=frame_result.frame_4.regex_command_text_area: self.add_brackets(e),
                  width=10).pack(side=tk.LEFT, padx=5)
        
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][5],
                  command=lambda e=frame_result.frame_4.regex_command_text_area: self.add_and(e),
                  width=10).pack(side=tk.LEFT, padx=5)
        
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][6],
                  command=lambda e=frame_result.frame_4.regex_command_text_area: self.add_or(e),
                  width=10).pack(side=tk.LEFT, padx=5)
        
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][7],
                  command=lambda e=frame_result.frame_4.regex_command_text_area: self.command_regex(e,text_area),
                  width=10).pack(side=tk.LEFT, padx=5)
        
        tk.Button(frame_result.frame_5, text=control_frame_config['button_name'][8],
                #   command=lambda e=frame_result.frame_1.entry_widget,
                #                  f=frame_result.frame_2.combobox_sheet,
                #                  g=frame_result.frame_3.combobox_row,
                #                  h=frame_result.frame_4.combobox_col,
                #                  i=frame_result.frame_5.combobox_value: self.input_sheet(e,f,g,h,i,text_area),
                  width=10).pack(side=tk.LEFT, padx=5)

        frame_result.frame_2.combobox_sheet.bind('<<ComboboxSelected>>',
                                                  lambda event: self.on_sheet_change(event,
                                                                                     frame_result.frame_1.entry_widget,
                                                                                     frame_result.frame_2.combobox_sheet,
                                                                                     frame_result.frame_2.combobox_col,))
        
        return frame_result


    # 选择文件按钮函数
    def select_sheet(self, path_widget, sheet_widget, combobox_col, text_area):

        fill_text = ''
        path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx'),
                                                     ('Excel Files', '*.xls')])

        if path:
            path_widget.config(state='normal')
            path_widget.delete(0, tk.END)
            path_widget.insert(0, path)
            path_widget.config(state='readonly')
            fill_text += f'Selected: {path}\n'
            fill_text += 'File selection successful!\n'

            sheets_name_result = sheetnames_import_fun.sheetnames_import(path)
            if sheets_name_result[0]:

                sheets_name_list = sheets_name_result[2]
                sheet_widget['values'] = sheets_name_list
                sheet_widget.set(sheets_name_list[0])               # 设置默认选择第一个
                sheet_widget.config(state='readonly')

            else:

                sheet_widget.set('')
                sheet_widget.config(state='disabled')

            fill_text += sheets_name_result[1]

        else:
            fill_text += f'File selection failed!\n'

        self.on_sheet_change(None, path_widget, sheet_widget, combobox_col)

        fill_area_text.text_area_fill(text_area, fill_text)


    # 自动更新列名列表
    def on_sheet_change(self, event, file_entry, combobox_sheet, combobox_col):

        file_path = file_entry.get()
        sheet_name = combobox_sheet.get()

        result = columns_title_fun.columns_title(file_path, sheet_name)

        if result[0]:
            new_options = result[1]

            combobox_col['values'] = new_options
            if new_options[0]:
                combobox_col.set(new_options[0])
            combobox_col.config(state='readonly')


    # 导入按钮函数
    def import_sheet(self, path_widget, sheet_widget, text_area):

        fill_text = ''
        path = path_widget.get()
        sheet_name = sheet_widget.get()

        if path:

            result_info = read_xlsx_xls_csv_txt_fun.read_xlsx_xls_csv_txt(original_file_path=path, sheet_name=sheet_name)
            if result_info[0]:
                self.data_df = result_info[2]
                self.data_df.reset_index(drop=True, inplace=True)
                fill_text += result_info[1]
            else:
                fill_text += result_info[1]
        else:
            fill_text += f'Please select a file first!\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # 数据预览
    def button_preview_fun(self, frame_path, text_area):

        path = frame_path.get()
        dir_path = os.path.dirname(path)
        file_path = os.path.join(dir_path, 'df_preview.xlsx')
        result_info = df_export_xlsx_fun.df_export_xlsx(self.data_df, file_path)
        fill_text = result_info[1]
        fill_area_text.text_area_fill(text_area, fill_text)
        subprocess.run(['open', file_path])


    # 添加正则表达式按钮函数
    def add_regex(self, combobox_col, combobox_label, entry_regex, text_area_regex):

        col = combobox_col.get()
        label = combobox_label.get()
        regex = entry_regex.get()

        if self.left_right:
            fill_regex = f"  ['{label}', '{col}', '{regex}'],\n"
        else:
            fill_regex = f"['{label}', '{col}', '{regex}'],\n"

        text_area_regex.insert(tk.INSERT, fill_regex)
        text_area_regex.see(tk.END)


    # []按钮函数
    def add_brackets(self, text_area_regex):

        if self.left_right:

            self.left_right = False
            fill_regex = '],\n'

            text_area_regex.insert(tk.INSERT, fill_regex)
            text_area_regex.see(tk.END)

        else:

            self.left_right = True
            fill_regex = '[\n'

            text_area_regex.insert(tk.INSERT, fill_regex)
            text_area_regex.see(tk.END)


    # &按钮函数
    def add_and(self, text_area_regex):

        if self.left_right:
            fill_regex = f"  ['AND'],\n"
            text_area_regex.insert(tk.INSERT, fill_regex)
            text_area_regex.see(tk.END)


    # &按钮函数
    def add_or(self, text_area_regex):

        if self.left_right:
            fill_regex = f"  ['OR'],\n"
            text_area_regex.insert(tk.INSERT, fill_regex)
            text_area_regex.see(tk.END)
    

    # 执行筛选按钮
    def command_regex(self, text_area_regex, text_area):

        fill_text = ''
        l_r = False
        regex_list = []

        text = text_area_regex.get('1.0', 'end-1c')

        text = text.strip()             # 去除首尾空白

        for line in text.splitlines():
            text_split = line.split(', ')
            if text_split[0] == "[''":

                regex_col = text_split[1].strip('"\'')
                regex_text = ', '.join(text_split[2:])
                regex_text = regex_text[1:-3]
                regex_append = ['', regex_col, regex_text]
                regex_list.append(regex_append)

            elif text_split[0] == "['~'":

                regex_col = text_split[1].strip('"\'')
                regex_text = ', '.join(text_split[2:])
                regex_text = regex_text[1:-3]
                regex_append = ['~', regex_col, regex_text]
                regex_list.append(regex_append)

            elif text_split[0] == '[':

                l_r = True
                regex_list.append([])

            elif text_split[0] == "  [''":

                regex_col = text_split[1].strip('"\'')
                regex_text = ', '.join(text_split[2:])
                regex_text = regex_text[1:-3]
                regex_append = ['', regex_col, regex_text]
                regex_list[-1].append(regex_append)

            elif text_split[0] == "  ['~'":

                regex_col = text_split[1].strip('"\'')
                regex_text = ', '.join(text_split[2:])
                regex_text = regex_text[1:-3]
                regex_append = ['~', regex_col, regex_text]
                regex_list[-1].append(regex_append)

            elif text_split[0] == "  ['AND'],":

                regex_list[-1].append(['AND'])

            elif text_split[0] == "  ['OR'],":

                regex_list[-1].append(['OR'])

            elif text_split[0] == '],':

                l_r = False

        for regex_l in regex_list:

            if_label = regex_l[0]
            column_index = regex_l[1]
            regex_pattern = regex_l[2]
            # 使用正則表達式篩選列並更新 DataFrame
            if if_label == '':
                self.data_df = self.data_df[self.data_df[column_index].str.contains(regex_pattern, regex=True, na=False)]
            elif if_label == '~':
                self.data_df = self.data_df[~self.data_df[column_index].str.contains(regex_pattern, regex=True, na=False)]
            else:
                mask = self.build_mask(self.data_df, regex_l)
                self.data_df = self.data_df[mask]

        fill_text += '筛选成功！\n'
        fill_area_text.text_area_fill(text_area, fill_text)


    # 多列筛选用函数
    def build_mask(self, df, conditions):

        final_mask = None
        pending_op = None   # 上一个 AND / OR

        for c in conditions:

            if len(c) == 3:
                regex_index = c[0]
                regex_col = c[1]
                regex_re = c[2]

                if regex_index == '~':
                    col_mask = ~df[regex_col].astype(str).str.contains(regex_re, regex=True, na=False)
                else:
                    col_mask = df[regex_col].astype(str).str.contains(regex_re, regex=True, na=False)

                if final_mask is None:
                    final_mask = col_mask
                else:
                    final_mask = pending_op(final_mask, col_mask)

            else:
                if c[0] == 'AND':
                    pending_op = operator.and_
                else:
                    pending_op = operator.or_

        return final_mask

        


'''
['~', '科目代码', '6711'],
[
  ['', '科目代码', '6001'],
  ['OR'],
  ['', '科目代码', '6401'],
  ['AND'],
  ['~', '科目名称', '主营业务收入'],
],
'''