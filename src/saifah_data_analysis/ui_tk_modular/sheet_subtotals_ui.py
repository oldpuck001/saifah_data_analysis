# sheet_subtotals_ui.py

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from . import fill_area_text
from ..DataFrame_tools_pd import sheetnames_import_fun

class sheet_subtotals_ui_class:

    def sheet_subtotals_ui_frame(self, root, control_frame_config, text_area):
        frame_result = tk.Frame(root)
        frame_result.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        frame_result.frame_1 = tk.Frame(frame_result)
        frame_result.frame_1.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_result.frame_1, text='File Path', width=10).pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_1.entry_widget = tk.Entry(frame_result.frame_1, state='readonly', readonlybackground='white')         # 创建Entry并保存引用
        frame_result.frame_1.entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

        options = []

        frame_result.frame_2 = tk.Frame(frame_result)
        frame_result.frame_2.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_result.frame_2, text='选择工作表', width=10).pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_2.combobox_sheet = ttk.Combobox(frame_result.frame_2, values=options, state='readonly')
        frame_result.frame_2.combobox_sheet.pack(side=tk.LEFT, padx=5, pady=5)

        frame_result.frame_3 = tk.Frame(frame_result)
        frame_result.frame_3.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_result.frame_3, text='选择行目录', width=10).pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_3.combobox_row = ttk.Combobox(frame_result.frame_3, values=options, state='readonly')
        frame_result.frame_3.combobox_row.pack(side=tk.LEFT, padx=5, pady=5)

        frame_result.frame_4 = tk.Frame(frame_result)
        frame_result.frame_4.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_result.frame_4, text='选择列标题', width=10).pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_4.combobox_col = ttk.Combobox(frame_result.frame_4, values=options, state='readonly')
        frame_result.frame_4.combobox_col.pack(side=tk.LEFT, padx=5, pady=5)

        frame_result.frame_5 = tk.Frame(frame_result)
        frame_result.frame_5.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Label(frame_result.frame_5, text='选择数值列', width=10).pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_5.combobox_value = ttk.Combobox(frame_result.frame_5, values=options, state='readonly')
        frame_result.frame_5.combobox_value.pack(side=tk.LEFT, padx=5, pady=5)

        frame_result.frame_6 = tk.Frame(frame_result)
        frame_result.frame_6.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Button(frame_result.frame_6, text=control_frame_config['button_name'][0], command=lambda e=frame_result.frame_1.entry_widget,f=frame_result.frame_2.combobox_sheet: self.input_sheet(e,f,text_area), width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_result.frame_6, text=control_frame_config['button_name'][1], width=10).pack(side=tk.LEFT, padx=5)

        frame_result.frame_2.combobox_sheet.bind('<<ComboboxSelected>>', self.on_sheet_change)

        return frame_result


    # 导入按钮函数
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


    def on_sheet_change(self, event=None):

        file_path = self.frame_result.frame_1.entry_widget.get()
        selected_sheet = self.frame_result.frame_2.combobox_sheet.get()





        new_values = sheet_to_items.get(selected_sheet, [])

        combobox_widget_2['values'] = new_values

        if new_values:
            combobox_widget_2.set(new_values[0])
        else:
            combobox_widget_2.set('')
    