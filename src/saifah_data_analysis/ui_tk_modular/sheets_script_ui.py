# sheets_script_ui.py

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from . import fill_area_text
from ..DataFrame_tools_pd import sheetnames_import_fun
from ..DataFrame_tools_pd import read_xlsx_xls_csv_txt_fun
from ..DataFrame_tools_pd import pd_concat_fun
from ..xlsx_tools_openxl import export_new_xlsx_fun

class sheets_script_ui_class:

    def sheets_script_ui_frame(self, root, control_frame_config, text_area):
        frame_result = tk.Frame(root)
        frame_result.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        frame_result.frames = []

        num = control_frame_config['button_name'][0]

        for n in range(num):

            frame = tk.Frame(frame_result)
            frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=3)

            tk.Label(frame, text=f'表格 {n+1}', width=5, anchor='w').pack(side=tk.LEFT, padx=5)

            frame.entry_widget = tk.Entry(frame, state='readonly', readonlybackground='white')                  # 创建Entry并保存引用
            frame.entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

            options = []                                                                                        # 创建选项列表
            frame.combobox_widget = ttk.Combobox(frame, values=options, height=10, width=25, state='readonly')         # 使用 ttk.Combobox（现代下拉选框）
            frame.combobox_widget.pack(side=tk.LEFT, padx=5)

            tk.Button(frame, text=control_frame_config['button_name'][1], command=lambda e=frame.entry_widget,f=frame.combobox_widget: self.input_sheet(e,f,text_area), width=10).pack(side=tk.LEFT, padx=5)
            tk.Button(frame, text=control_frame_config['button_name'][2], command=lambda e=frame.entry_widget,f=frame.combobox_widget: self.del_sheet(e,f,text_area), width=10).pack(side=tk.LEFT, padx=5)

            frame_result.frames.append(frame)

        frame_export = tk.Frame(frame_result)
        frame_export.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=(0,10))
        tk.Button(frame_export, text=control_frame_config['button_name'][3], command=lambda e=frame_result.frames: self.output_sheet(e,num,text_area), width=10).pack()

        return frame_result
    

    # 添加按钮函数
    def input_sheet(self, entry_widget, combobox_widget, text_area):
        fill_text = ''

        path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx'),
                                                     ('Excel Files', '*.xls')])

        if path:
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            entry_widget.config(state='readonly')
            fill_text += f'Selected: {path}\n'
            fill_text += 'File selection successful!\n'

            sheets_name_result = sheetnames_import_fun.sheetnames_import(path)
            if sheets_name_result[0]:

                sheets_name_list = sheets_name_result[2]
                sheets_name_list.insert(0, 'All worksheets')

                combobox_widget['values'] = sheets_name_list
                combobox_widget.set(sheets_name_list[0])               # 设置默认选择第一个
                combobox_widget.config(state='readonly')

            else:

                combobox_widget.set('')
                combobox_widget.config(state='disabled')

            fill_text += sheets_name_result[1]

        else:
            fill_text += f'File selection failed!\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # 删除按钮函数
    def del_sheet(self, entry_widget, combobox_widget, text_area):

        fill_text = ''

        entry_widget.config(state='normal')
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, '')
        entry_widget.config(state='readonly')
        fill_text += 'Del path successful!\n'

        combobox_widget['values'] = []
        combobox_widget.set('')
        fill_text += 'Del sheets name successful!\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # 导出按钮函数
    def output_sheet(self, frames, num, text_area):

        fill_text = ''
        path_list = [None] * num
        sheet_list = [None] * num
        df_list = []

        for n in range(num):
            path_list[n] = frames[n].entry_widget.get()
            sheet_list[n] = frames[n].combobox_widget.get()

        for n in range(num):
            if path_list[n]:
                if sheet_list[n] == 'All worksheets':
                    sheetnames = sheetnames_import_fun.sheetnames_import(path_list[n])[2]
                    for sheet in sheetnames:
                        df_result = read_xlsx_xls_csv_txt_fun.read_xlsx_xls_csv_txt(path_list[n], sheet_name=sheet)
                        if df_result[0]:
                            df_list.append(df_result[2])
                else:
                    df_result = read_xlsx_xls_csv_txt_fun.read_xlsx_xls_csv_txt(path_list[n], sheet_name=sheet_list[n])
                    if df_result[0]:
                        df_list.append(df_result[2])

        df_export = pd_concat_fun.pd_concat(df_list)

        fill_text += export_new_xlsx_fun.export_new_xlsx(df_export)

        fill_area_text.text_area_fill(text_area, fill_text)