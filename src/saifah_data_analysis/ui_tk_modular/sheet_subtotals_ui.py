# sheet_subtotals_ui.py

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from . import fill_area_text
from ..DataFrame_tools_pd import sheetnames_import_fun
from ..DataFrame_tools_pd import columns_title_fun
from ..DataFrame_tools_pd import sheet_pivot_table_fun

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
        tk.Button(frame_result.frame_6, text=control_frame_config['button_name'][0],
                  command=lambda e=frame_result.frame_1.entry_widget,
                                 f=frame_result.frame_2.combobox_sheet,
                                 g=frame_result.frame_3.combobox_row,
                                 h=frame_result.frame_4.combobox_col,
                                 i=frame_result.frame_5.combobox_value: self.input_sheet(e,f,g,h,i,text_area),
                  width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_result.frame_6, text=control_frame_config['button_name'][1],
                  command=lambda e=frame_result.frame_1.entry_widget,
                                 f=frame_result.frame_2.combobox_sheet,
                                 g=frame_result.frame_3.combobox_row,
                                 h=frame_result.frame_4.combobox_col,
                                 i=frame_result.frame_5.combobox_value: self.subtotals_generate(e,f,g,h,i,text_area),
                  width=10).pack(side=tk.LEFT, padx=5)

        frame_result.frame_2.combobox_sheet.bind('<<ComboboxSelected>>',
                                                  lambda event: self.on_sheet_change(event,
                                                                                     frame_result.frame_1.entry_widget,
                                                                                     frame_result.frame_2.combobox_sheet,
                                                                                     frame_result.frame_3.combobox_row,
                                                                                     frame_result.frame_4.combobox_col,
                                                                                     frame_result.frame_5.combobox_value))

        return frame_result


    # 导入按钮函数
    def input_sheet(self, entry_widget, combobox_widget, combobox_row, combobox_col, combobox_value, text_area):

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

        self.on_sheet_change(None, entry_widget, combobox_widget, combobox_row, combobox_col, combobox_value)

        fill_area_text.text_area_fill(text_area, fill_text)


    # 自动更新列名列表
    def on_sheet_change(self, event, file_entry, combobox_sheet, combobox_row, combobox_col, combobox_value):

        file_path = file_entry.get()
        sheet_name = combobox_sheet.get()

        result = columns_title_fun.columns_title(file_path, sheet_name)

        if result[0]:
            new_options = result[1]

            combobox_row['values'] = new_options
            if new_options[0]:
                combobox_row.set(new_options[0])
            combobox_row.config(state='readonly')

            combobox_col['values'] = new_options
            if new_options[0]:
                combobox_col.set(new_options[0])
            combobox_col.config(state='readonly')

            combobox_value['values'] = new_options
            if new_options[0]:
                combobox_value.set(new_options[0])
            combobox_value.config(state='readonly')


    # 导出按钮函数
    def subtotals_generate(self, file_path, sheet_name, combobox_row, combobox_col, combobox_value, text_area):

        fill_text = ''
        file_path = file_path.get()
        sheet_name = sheet_name.get()
        row_value = combobox_row.get()
        column_value = combobox_col.get()
        subtotal_value = combobox_value.get()

        if column_value == row_value or column_value == subtotal_value or row_value == subtotal_value:

            fill_text += '行目录分类列、列标题分类列、数值列必须为不同列，请重新选择！'

            fill_area_text.text_area_fill(text_area, fill_text)
        
        else:

            df_pivot_result = sheet_pivot_table_fun.sheet_pivot_table(file_path, sheet_name, row_value, column_value, subtotal_value)
            if df_pivot_result[0]:
                df_pivot_table = df_pivot_result[1]
            print(df_pivot_table)

# from openpyxl.utils import get_column_letter
#             # 创建一个Excel工作簿
#             wb = openpyxl.Workbook()
#             ws = wb.active

#             # 重設索引並將DataFrame的數據添加到Excel工作表
#             exportdf.reset_index(inplace=True)

#             # 修改DataFrame的第1列第1行（列标题）为空白
#             exportdf.columns.values[0] = ''
#             for r in openpyxl.utils.dataframe.dataframe_to_rows(exportdf, index=False, header=True):
#                 ws.append(r)

#             # 调整列宽
#             for column_cells in ws.columns:
#                 length = max(len(str(cell.value))+8 for cell in column_cells)
#                 ws.column_dimensions[column_cells[0].column_letter].width = length

#             # 添加合计公式到最后一列
#             colnumber = ws.max_column + 1
#             row_sum_title = ws.cell(row=1, column=colnumber)
#             row_sum_title.value = 'Total'
#             for row_index, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column)):
#                 row_sum_cell = ws.cell(row=row_index + 2, column=colnumber)
#                 row_sum_cell.value = f'=SUM({row[1].coordinate}:{row[-2].coordinate})'
#                 row_sum_cell.number_format = '#,##0.00'

#             # 添加合计公式到最后一行
#             rownumber = ws.max_row+1
#             col_sum_title = ws.cell(row=rownumber, column=1)
#             col_sum_title.value = 'Total'
#             for col_index in range(2, ws.max_column + 1):
#                 col_letter = get_column_letter(col_index)
#                 col_sum_cell = ws.cell(row=rownumber, column=col_index)
#                 col_sum_cell.value = f'=SUM({col_letter}2:{col_letter}{ws.max_row-1})'
#                 col_sum_cell.number_format = '#,##0.00'
                
#             # 设置数值单元格格式为两位小数且有千分位符
#             for row in ws.iter_rows(min_row=2, max_row=ws.max_row - 1, min_col=2, max_col=ws.max_column):
#                 for cell in row:
#                     if isinstance(cell.value, (int, float)):
#                         cell.number_format = '#,##0.00'  # 设置为带千分位和两位小数

#             # 保存Excel文件
#             wb.save(save_path)

#         result_text = {'result_message': '生成成功！'}

#         return ['subtotals_generate', result_text]