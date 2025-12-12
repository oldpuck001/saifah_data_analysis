# openxl_xlsx_tool.py

import openpyxl
import openpyxl.styles
from openpyxl import load_workbook

class openxl_xlsx_tool_class:

    def __init__(self):
        pass

    def in_out_result_new_xlsx(self, file_path):
        # 实例化工作簿
        wb = openpyxl.Workbook()

        # 获取第一张工作表并赋予它一个名称
        ws_1 = wb.active
        ws_1.title = '数值合计对比'

        # 设置列宽
        ws_1.column_dimensions['A'].width = 15
        ws_1.column_dimensions['B'].width = 15
        ws_1.column_dimensions['C'].width = 15
        ws_1.column_dimensions['D'].width = 15

        # 设置行高
        ws_1.row_dimensions[1].height = 25
        ws_1.row_dimensions[2].height = 20
        ws_1.row_dimensions[3].height = 20
        ws_1.row_dimensions[4].height = 20

        # 设置边框
        thin_border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'),
                                            right=openpyxl.styles.Side(style='thin'),
                                            top=openpyxl.styles.Side(style='thin'),
                                            bottom=openpyxl.styles.Side(style='thin'))

        for row in ws_1.iter_rows(min_row=1, min_col=1, max_row=4, max_col=4):
            for cell in row:
                cell.border = thin_border

        ws_1['A1'].value = 'table_name'
        ws_1['A1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

        ws_1['B1'].value = '笔数'
        ws_1['B1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

        ws_1['C1'].value = 'In 金额'
        ws_1['C1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

        ws_1['D1'].value = 'Out 金额'
        ws_1['D1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

        ws_1['A2'].value = 'table 1'
        ws_1['A2'].alignment = openpyxl.styles.Alignment(vertical='center')

        ws_1['B2'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

        ws_1['C2'].number_format = '#,##0.00'
        ws_1['C2'].alignment = openpyxl.styles.Alignment(vertical='center')

        ws_1['D2'].number_format = '#,##0.00'
        ws_1['D2'].alignment = openpyxl.styles.Alignment(vertical='center')

        ws_1['A3'].value = 'table 2'
        ws_1['A3'].alignment = openpyxl.styles.Alignment(vertical='center')

        ws_1['B3'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

        ws_1['C3'].number_format = '#,##0.00'
        ws_1['C3'].alignment = openpyxl.styles.Alignment(vertical='center')

        ws_1['D3'].number_format = '#,##0.00'
        ws_1['D3'].alignment = openpyxl.styles.Alignment(vertical='center')

        ws_1['A4'].value = '差异'
        ws_1['A4'].alignment = openpyxl.styles.Alignment(vertical='center')

        ws_1['B4'].value = '=B2-B3'
        ws_1['B4'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

        ws_1['C4'].value = '=C2-C3'
        ws_1['C4'].number_format = '#,##0.00'
        ws_1['C4'].alignment = openpyxl.styles.Alignment(vertical='center')

        ws_1['D4'].value = '=D2-D3'
        ws_1['D4'].number_format = '#,##0.00'
        ws_1['D4'].alignment = openpyxl.styles.Alignment(vertical='center')

        # 获取第二张工作表并赋予它一个名称
        wb.create_sheet(title='数值分层对比_in_共同数值但计数不同_table_1')
        wb.create_sheet(title='数值分层对比_in_共同数值但计数不同_table_2')
        wb.create_sheet(title='数值分层对比_in_只在 table 1 中出现的数值')
        wb.create_sheet(title='数值分层对比_in_只在 table 2 中出现的数值')
        wb.create_sheet(title='数值分层对比_out_共同数值但计数不同_table_1')
        wb.create_sheet(title='数值分层对比_out_共同数值但计数不同_table_2')
        wb.create_sheet(title='数值分层对比_out_只在 table 1 中出现的数值')
        wb.create_sheet(title='数值分层对比_out_只在 table 2 中出现的数值')

        # 保存工作簿会在磁盘上创建文件
        wb.save(file_path)

        return f'The xlsx output file path: {file_path}\nThe xlsx output file was created successfully!\n'

    def in_out_result_fill_total_xlsx(self, file_path, quantity_table_1_df, quantity_table_2_df, amount_in_table_1_df, amount_out_table_1_df, amount_in_table_2_df, amount_out_table_2_df):

        wb = openpyxl.load_workbook(file_path, data_only=False)
        # ws = wb.active
        ws = wb['数值合计对比']

        ws['B2'].value = quantity_table_1_df
        ws['B3'].value = quantity_table_2_df
        ws['C2'].value = amount_in_table_1_df
        ws['C3'].value = amount_in_table_2_df
        ws['D2'].value = amount_out_table_1_df
        ws['D3'].value = amount_out_table_2_df

        wb.save(file_path)

        return 'The xlsx output file was updata successfully!\n'

    def in_out_result_fill_layering_xlsx(self, file_path, fill_sheet, fill_df):

        wb = load_workbook(file_path)
        ws = wb[fill_sheet]

        # 写入数据
        header = fill_df.columns.tolist()

        ws.append(header)
        for row in fill_df.values:
            ws.append(row.tolist())

        wb.save(file_path)

        return 'The xlsx output file was updata successfully!\n'