# pd_DataFrame_tool.py

import os
import pandas as pd
from .openxl_xlsx_tool import openxl_xlsx_tool_class

openxl_xlsx_tool_class_example = openxl_xlsx_tool_class()

class pd_DataFrame_tool_class:

    # 获取工作表名称列表
    def sheetnames_import(self, original_file_path=None):
        try:
            sheet_file = pd.ExcelFile(original_file_path)
            sheetnames = sheet_file.sheet_names
            info = 'Worksheet list successfully read!\n'
            return [True, info, sheetnames]
        except:
            sheetnames = []
            info = 'Failed to read worksheet list!\nThis feature is only supported for xlsx and xls format files.\n'
            return [False, info, sheetnames]

    # 读取数据文件
    def read_xlsx_xls_csv_txt_fun(self, original_file_path = None,
                                        sheet_name = 0,                     # 默认读取第一个工作表 (索引0)
                                        skiprows = None,                    # 默认不跳过任何行
                                        usecols = None,                     # 默认读取所有列
                                        nrows = None,                       # 默认读取所有行
                                        index_col = None,                   # 默认无索引列（生成RangeIndex）
                                        header = 0,                         # 默认第0行（第一行）作为列名，设为 None 时无列名
                                        names = None,                       # 默认使用Excel中的列名（由header决定）。提供列表将覆盖现有列名，通常与 header=None 配合使用
                                        na_values = None,                   # 默认无额外空值标识。提供字符串、列表或字典（按列指定），将指定值识别为NaN
                                        keep_default_na = True,             # 默认启用pandas内置空值识别（如"", "#N/A", "NULL"等）。设为False则仅识别na_values指定的空值
                                        dtype = None,                       # 默认自动推断列数据类型。字典格式：{列名: 数据类型}，强制指定列的数据类型
                                        converters = None,                  # 默认无列转换函数。字典格式：{列索引/列名: 函数}，在读取时对指定列应用函数
                                        sep=',',                            # csv参数，默认逗号
                                        encoding='utf-8',                   # csv参数，默认utf-8
                                        engine_csv='c'                      # csv参数，默认c
                                    ):

        skiprows        = None if skiprows == 'None' else int(skiprows)
        usecols         = None if usecols == 'None' else usecols
        nrows           = None if nrows == 'None' else nrows
        index_col       = None if index_col == 'None' else index_col 
        header          = 0    if header == '0' else int(header)
        names           = None if names == 'None' else names
        na_values       = None if na_values == 'None' else na_values
        keep_default_na = True if keep_default_na == 'True' else False
        dtype           = None if dtype == 'None' else dtype
        converters      = None if converters == 'None' else converters

        try:
            file_extension = os.path.splitext(original_file_path)[1].lower()

            if file_extension in ['.xlsx', '.xls']:
                if file_extension == '.xlsx':
                    engine = 'openpyxl'
                elif file_extension == '.xls':
                    engine = 'xlrd'
                df = pd.read_excel(original_file_path,
                                    sheet_name = sheet_name,             # 默认读取第一个工作表 (索引0)
                                    skiprows = skiprows,                 # 默认不跳过任何行
                                    usecols = usecols,                   # 默认读取所有列
                                    nrows = nrows,                       # 默认读取所有行
                                    index_col = index_col,               # 默认无索引列（生成RangeIndex）
                                    header = header,                     # 默认第0行（第一行）作为列名，设为 None 时无列名
                                    names = names,                       # 默认使用Excel中的列名（由header决定）。提供列表将覆盖现有列名，通常与 header=None 配合使用
                                    na_values = na_values,               # 默认无额外空值标识。提供字符串、列表或字典（按列指定），将指定值识别为NaN
                                    keep_default_na = keep_default_na,   # 默认启用pandas内置空值识别（如"", "#N/A", "NULL"等）。设为False则仅识别na_values指定的空值
                                    dtype = dtype,                       # 默认自动推断列数据类型。字典格式：{列名: 数据类型}，强制指定列的数据类型
                                    converters = converters,             # 默认无列转换函数。字典格式：{列索引/列名: 函数}，在读取时对指定列应用函数
                                    engine = engine)

            elif file_extension in ['.csv', '.txt']:
                if file_extension == '.csv':
                    engine = engine_csv
                elif file_extension == '.txt':
                    engine = 'python'
                df = pd.read_csv(original_file_path,
                                    sep = sep,                          # 默认逗号
                                    encoding = encoding,                # 默认utf-8
                                    skiprows = skiprows,                # 默认不跳过任何行
                                    usecols = usecols,                  # 默认读取所有列
                                    nrows = nrows,                      # 默认读取所有行
                                    index_col = index_col,              # 默认无索引列（生成RangeIndex）
                                    header = header,                    # 默认第0行（第一行）作为列名，设为 None 时无列名
                                    names = names,                      # 默认使用Excel中的列名（由header决定）。提供列表将覆盖现有列名，通常与 header=None 配合使用
                                    na_values = na_values,              # 默认无额外空值标识。提供字符串、列表或字典（按列指定），将指定值识别为NaN
                                    keep_default_na = keep_default_na,  # 默认启用pandas内置空值识别（如"", "#N/A", "NULL"等）。设为False则仅识别na_values指定的空值
                                    dtype = dtype,                      # 默认自动推断列数据类型。字典格式：{列名: 数据类型}，强制指定列的数据类型
                                    converters = converters,            # 默认无列转换函数。字典格式：{列索引/列名: 函数}，在读取时对指定列应用函数
                                    engine = engine)

            else:
                info = 'Data file format is not supported.\nPlease try importing again.\n'
                df = pd.DataFrame()
                return [False, info, df]

        except Exception as e:
            info = f'Data file reading failed.\nPlease try importing again.\n{e}\n'
            df = pd.DataFrame()
            return [False, info, df]

        info = 'Data file read successfully!\n'
        return [True, info, df]

    # 导出xlsx文件
    def df_export_xlsx(self, df, path):

        try:
            df.to_excel(path)
            info = f'Export successfully!\nFile path: {path}\n'
            return [True, info]
        
        except:
            info = 'Export failed!\n'
            return [False, info]

    # 数据清洗
    def df_cleaning(self, df, clean_col, clean_option):

        info = f'DataFrame column: {clean_col}\n数据清洗选项：{clean_option}\n'

        try:
            if clean_option == '删除指定列的重复行':
                df = df.drop_duplicates(subset=[clean_col])                                         # 删除指定列的重复行

            elif clean_option == '填充缺失值为 0':
                df[clean_col] = df[clean_col].fillna(0)                                             # 填充缺失值为 0

            elif clean_option == '填充缺失值为 <空白>':
                df[clean_col] = df[clean_col].fillna('<空白>')                                      # 填充缺失值为 '<空白>'

            elif clean_option == '填充缺失值为 重复上一行':
                df[clean_col] = df[clean_col].ffill()                                               # 填充缺失值为 重复上一行

            elif clean_option == '标准化文本（去除首尾空格及转换为小写英文字母）':
                df[clean_col] = df[clean_col].str.strip().str.lower()                               # 标准化文本（去除首尾空格及转换为小写英文字母）

            elif clean_option == '将数据类型转换为字符型':
                df[clean_col] = df[clean_col].astype(str)                                           # 将数据类型转换为字符型

            elif clean_option == '转换为整数，支持缺失值':
                df[clean_col] = df[clean_col].astype(str).str.replace(',', '')
                df[clean_col] = pd.to_numeric(df[clean_col], errors='coerce').fillna(0).astype(int) # 转换为整数，支持缺失值

            elif clean_option == '转换为浮点数，支持缺失值':
                df[clean_col] = df[clean_col].astype(str).str.replace(',', '')
                df[clean_col] = pd.to_numeric(df[clean_col], errors='coerce').astype(float)         # 转换为浮点数，支持缺失值

            elif clean_option == '将数据类型转换为时间日期类型':
                df[clean_col] = pd.to_datetime(df[clean_col], errors='coerce')                      # 将数据类型转换为时间日期类型


            info += 'Data cleaning successful!\n'
            return [True, info]
        
        except:

            info += 'Data cleaning failed!\n'
            return [False, info]


    # input output value check 对比模式一：总量、总额对比
    def input_output_value_check_sum(self, table_1_name, table_1_in_out_mode, sql_table_1_df, table_1_in_col, table_1_out_col, table_1_in_out_value,
                                     table_1_in_out_col, table_1_in_label, table_1_out_label, table_2_name, table_2_in_out_mode, sql_table_2_df,
                                     table_2_in_col, table_2_out_col, table_2_in_out_value, table_2_in_out_col, table_2_in_label, table_2_out_label,
                                     file_path):

        fill_text = ''

        if table_1_in_out_mode == '双列模式':
            quantity_in_table_1_df = sql_table_1_df[table_1_in_col].dropna().shape[0]
            quantity_out_table_1_df = sql_table_1_df[table_1_out_col].dropna().shape[0]
            amount_in_table_1_df = round(sql_table_1_df[table_1_in_col].sum(), 2)
            amount_out_table_1_df = round(sql_table_1_df[table_1_out_col].sum(), 2)
        elif table_1_in_out_mode == '+/-单列模式':
            quantity_in_table_1_df = sql_table_1_df[sql_table_1_df[table_1_in_out_value] > 0][table_1_in_out_value].dropna().shape[0]
            quantity_out_table_1_df = sql_table_1_df[sql_table_1_df[table_1_in_out_value] < 0][table_1_in_out_value].dropna().shape[0]
            amount_in_table_1_df = round(sql_table_1_df[sql_table_1_df[table_1_in_out_value] > 0][table_1_in_out_value].sum(), 2)
            amount_out_table_1_df = round(-sql_table_1_df[sql_table_1_df[table_1_in_out_value] < 0][table_1_in_out_value].sum(), 2)
        elif table_1_in_out_mode == '标识列单列模式':
            quantity_in_table_1_df = sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_in_label][table_1_in_out_value].dropna().shape[0]
            quantity_out_table_1_df = sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_out_label][table_1_in_out_value].dropna().shape[0]
            amount_in_table_1_df = round(sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_in_label][table_1_in_out_value].sum(), 2)
            amount_out_table_1_df = round(sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_out_label][table_1_in_out_value].sum(), 2)

        if table_2_in_out_mode == '双列模式':
            quantity_in_table_2_df = sql_table_2_df[table_2_in_col].dropna().shape[0]
            quantity_out_table_2_df = sql_table_2_df[table_2_out_col].dropna().shape[0]
            amount_in_table_2_df = round(sql_table_2_df[table_2_in_col].sum(), 2)
            amount_out_table_2_df = round(sql_table_2_df[table_2_out_col].sum(), 2)
        elif table_2_in_out_mode == '+/-单列模式':
            quantity_in_table_2_df = sql_table_2_df[sql_table_2_df[table_2_in_out_value] > 0][table_2_in_out_value].dropna().shape[0]
            quantity_out_table_2_df = sql_table_2_df[sql_table_2_df[table_2_in_out_value] < 0][table_2_in_out_value].dropna().shape[0]
            amount_in_table_2_df = round(sql_table_2_df[sql_table_2_df[table_2_in_out_value] > 0][table_2_in_out_value].sum(), 2)
            amount_out_table_2_df = round(-sql_table_2_df[sql_table_2_df[table_2_in_out_value] < 0][table_2_in_out_value].sum(), 2)

        elif table_2_in_out_mode == '标识列单列模式':
            quantity_in_table_2_df = sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_in_label][table_2_in_out_value].dropna().shape[0]
            quantity_out_table_2_df = sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_out_label][table_2_in_out_value].dropna().shape[0]
            amount_in_table_2_df = round(sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_in_label][table_2_in_out_value].sum(), 2)
            amount_out_table_2_df = round(sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_out_label][table_2_in_out_value].sum(), 2)

        fill_in_xlsx = [['数值合计对比', 0, 1],
                        [25, 20, 20, 20],
                        [20,                20,                          20,                        20,                           20],
                        [['table',      1], ['quantity_in',          1], ['amount_in',          1], ['quantity_out',          1], ['amount_out',          1]],
                        [[table_1_name, 2], [quantity_in_table_1_df, 3], [amount_in_table_1_df, 4], [quantity_out_table_1_df, 3], [amount_out_table_1_df, 4]],
                        [[table_2_name, 2], [quantity_in_table_2_df, 3], [amount_in_table_2_df, 4], [quantity_out_table_2_df, 3], [amount_out_table_2_df, 4]],
                        [['difference', 1], ['=B2-B3',               3], ['=C2-C3',             4], ['=D2-D3',                3], ['=E2-E3',              4]]
                        ]

        # 填写.xlsx文件
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_cell(file_path, fill_in_xlsx)

        return fill_text


    # input output value check 对比模式二：金额分层对比
    # input output value check 对比模式三：按item分类对比
    # input output value check 对比模式四：按时序排列
    def input_output_value_check_layering(self, table_1_name, table_2_name,
                                          table_1_in_out_mode, sql_table_1_df, table_1_in_col, table_1_out_col, table_1_in_out_value,
                                          table_1_in_out_col, table_1_in_label, table_1_out_label,
                                          table_2_in_out_mode, sql_table_2_df, table_2_in_col, table_2_out_col, table_2_in_out_value,
                                          table_2_in_out_col, table_2_in_label, table_2_out_label,
                                          table_1_item, table_2_item,
                                          file_path):

        fill_text = ''

        # 获取分层值
        if table_1_in_out_mode == '双列模式':
            amounts_table_1_in_counts = sql_table_1_df[table_1_in_col].dropna().value_counts().sort_index().reset_index()
            amounts_table_1_in_counts.columns = ['value', 'count_table_1']
            amounts_table_1_in_counts_set = set(amounts_table_1_in_counts['value'])
            table_1_in_df = sql_table_1_df[sql_table_1_df[table_1_in_col].isin(amounts_table_1_in_counts_set)].dropna(subset=[table_1_in_col])
            table_1_in_df_col = table_1_in_col

            amounts_table_1_out_counts = sql_table_1_df[table_1_out_col].dropna().value_counts().sort_index().reset_index()
            amounts_table_1_out_counts.columns = ['value', 'count_table_1']
            amounts_table_1_out_counts_set = set(amounts_table_1_out_counts['value'])
            table_1_out_df = sql_table_1_df[sql_table_1_df[table_1_out_col].isin(amounts_table_1_out_counts_set)].dropna(subset=[table_1_out_col])
            table_1_out_df_col = table_1_out_col

        elif table_1_in_out_mode == '+/-单列模式':
            amounts_table_1_in_counts = sql_table_1_df[sql_table_1_df[table_1_in_out_value] > 0][table_1_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_1_in_counts.columns = ['value', 'count_table_1']
            amounts_table_1_in_counts_set = set(amounts_table_1_in_counts['value'])
            table_1_in_df = sql_table_1_df[sql_table_1_df[table_1_in_out_value].isin(amounts_table_1_in_counts_set)].dropna(subset=[table_1_in_out_value])
            table_1_in_df_col = table_1_in_out_value

            amounts_table_1_out_counts = sql_table_1_df[sql_table_1_df[table_1_in_out_value] < 0][table_1_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_1_out_counts.columns = ['value', 'count_table_1']
            amounts_table_1_out_counts['value'] = -amounts_table_1_out_counts['value']              # 负负得正
            amounts_table_1_out_counts_set = set(amounts_table_1_out_counts['value'])
            amounts_table_1_out_counts_non = pd.DataFrame()
            amounts_table_1_out_counts_non['value'] = -amounts_table_1_out_counts['value']           # 不改变负号
            amounts_table_1_out_counts_set_non = set(amounts_table_1_out_counts_non['value'])
            table_1_out_df = sql_table_1_df[sql_table_1_df[table_1_in_out_value].isin(amounts_table_1_out_counts_set_non)].dropna(subset=[table_1_in_out_value])
            table_1_out_df_col = table_1_in_out_value

        elif table_1_in_out_mode == '标识列单列模式':
            amounts_table_1_in_counts = sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_in_label][table_1_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_1_in_counts.columns = ['value', 'count_table_1']
            amounts_table_1_in_counts_set = set(amounts_table_1_in_counts['value'])
            table_1_in_df = sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_in_label].dropna(subset=[table_1_in_out_value])
            table_1_in_df_col = table_1_in_out_value

            amounts_table_1_out_counts = sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_out_label][table_1_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_1_out_counts.columns = ['value', 'count_table_1']
            amounts_table_1_out_counts_set = set(amounts_table_1_out_counts['value'])
            table_1_out_df = sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_out_label].dropna(subset=[table_1_in_out_value])
            table_1_out_df_col = table_1_in_out_value
        
        if table_2_in_out_mode == '双列模式':
            amounts_table_2_in_counts = sql_table_2_df[table_2_in_col].dropna().value_counts().sort_index().reset_index()
            amounts_table_2_in_counts.columns = ['value', 'count_table_2']
            amounts_table_2_in_counts_set = set(amounts_table_2_in_counts['value'])
            table_2_in_df = sql_table_2_df[sql_table_2_df[table_2_in_col].isin(amounts_table_2_in_counts_set)].dropna(subset=[table_2_in_col])
            table_2_in_df_col = table_2_in_col

            amounts_table_2_out_counts = sql_table_2_df[table_2_out_col].dropna().value_counts().sort_index().reset_index()
            amounts_table_2_out_counts.columns = ['value', 'count_table_2']
            amounts_table_2_out_counts_set = set(amounts_table_2_out_counts['value'])
            table_2_out_df = sql_table_2_df[table_2_out_col].dropna()
            table_2_out_df = sql_table_2_df[sql_table_2_df[table_2_out_col].isin(amounts_table_2_out_counts_set)].dropna(subset=[table_2_out_col])
            table_2_out_df_col = table_2_out_col

        elif table_2_in_out_mode == '+/-单列模式':
            amounts_table_2_in_counts = sql_table_2_df[sql_table_2_df[table_2_in_out_value] > 0][table_2_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_2_in_counts.columns = ['value', 'count_table_2']
            amounts_table_2_in_counts_set = set(amounts_table_2_in_counts['value'])
            table_2_in_df = sql_table_2_df[sql_table_2_df[table_2_in_out_value].isin(amounts_table_2_in_counts_set)].dropna(subset=[table_2_in_out_value])
            table_2_in_df_col = table_2_in_out_value

            amounts_table_2_out_counts = sql_table_2_df[sql_table_2_df[table_2_in_out_value] < 0][table_2_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_2_out_counts.columns = ['value', 'count_table_2']
            amounts_table_2_out_counts['value'] = -amounts_table_2_out_counts['value']                  # 负负得正
            amounts_table_2_out_counts_set = set(amounts_table_2_out_counts['value'])
            amounts_table_2_out_counts_non = pd.DataFrame()
            amounts_table_2_out_counts_non['value'] = -amounts_table_2_out_counts['value']              # 不改变负号
            amounts_table_2_out_counts_set_non = set(amounts_table_2_out_counts_non['value'])
            table_2_out_df = sql_table_2_df[sql_table_2_df[table_2_in_out_value].isin(amounts_table_2_out_counts_set_non)].dropna(subset=[table_2_in_out_value])
            table_2_out_df_col = table_2_in_out_value

        elif table_2_in_out_mode == '标识列单列模式':
            amounts_table_2_in_counts = sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_in_label][table_2_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_2_in_counts.columns = ['value', 'count_table_2']
            amounts_table_2_in_counts_set = set(amounts_table_2_in_counts['value'])
            table_2_in_df = sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_in_label].dropna(subset=[table_2_in_out_value])
            table_2_in_df_col = table_2_in_out_value

            amounts_table_2_out_counts = sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_out_label][table_2_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_2_out_counts.columns = ['value', 'count_table_2']
            amounts_table_2_out_counts_set = set(amounts_table_2_out_counts['value'])
            table_2_out_df = sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_out_label].dropna(subset=[table_2_in_out_value])
            table_2_out_df_col = table_2_in_out_value

        # 对比模式二：金额分层对比
        # 共同值集合
        common_amounts_in = amounts_table_1_in_counts_set & amounts_table_2_in_counts_set
        common_amounts_out = amounts_table_1_out_counts_set & amounts_table_2_out_counts_set

        # 共同金额的计数差异
        common_table_1_in = amounts_table_1_in_counts[amounts_table_1_in_counts['value'].isin(common_amounts_in)].set_index('value')
        common_table_2_in = amounts_table_2_in_counts[amounts_table_2_in_counts['value'].isin(common_amounts_in)].set_index('value')
        common_table_1_out = amounts_table_1_out_counts[amounts_table_1_out_counts['value'].isin(common_amounts_out)].set_index('value')
        common_table_2_out = amounts_table_2_out_counts[amounts_table_2_out_counts['value'].isin(common_amounts_out)].set_index('value')

        # 合并对比
        comparison_in = common_table_1_in.join(common_table_2_in, how='inner', lsuffix='_table_1', rsuffix='_table_2')
        comparison_in['difference'] = comparison_in['count_table_1'] - comparison_in['count_table_2']
        comparison_out = common_table_1_out.join(common_table_2_out, how='inner', lsuffix='_table_1', rsuffix='_table_2')
        comparison_out['difference'] = comparison_out['count_table_1'] - comparison_out['count_table_2']

        # 筛选所有计数差异不为0的项（共同金额但计数不同）
        non_zero_diff_in = comparison_in[comparison_in['difference'] != 0][['count_table_1', 'count_table_2', 'difference']]
        non_zero_diff_out = comparison_out[comparison_out['difference'] != 0][['count_table_1', 'count_table_2', 'difference']]

        # 找出只在流水或明细中出现的金额
        only_table_1_in = amounts_table_1_in_counts_set - amounts_table_2_in_counts_set
        only_table_2_in = amounts_table_2_in_counts_set - amounts_table_1_in_counts_set
        only_table_1_out = amounts_table_1_out_counts_set - amounts_table_2_out_counts_set
        only_table_2_out = amounts_table_2_out_counts_set - amounts_table_1_out_counts_set

        # input
        height_list = [25,]
        fill_in_xlsx = [['数值合计对比', 6, 1],
                        height_list,
                        [20,                 20,                20,                20               ],
                        [['input_value', 1], [table_1_name, 1], [table_2_name, 1], ['difference', 1]],
                        ]

        # 第一部分：共同金额但计数不同的情况
        if len(non_zero_diff_in) > 0:            
            for amount in non_zero_diff_in.index:
                table_1_count = comparison_in.loc[amount, 'count_table_1']
                table_2_count = comparison_in.loc[amount, 'count_table_2']
                diff_desc = table_1_count - table_2_count
                fill_in_xlsx.append([[amount, 4], [table_1_count, 3], [table_2_count, 3], [diff_desc, 3]])
                height_list.append(20)

        # 第二部分：只在 table 1 中出现的金额
        if len(only_table_1_in) > 0:            
            only_table_1_df = amounts_table_1_in_counts[amounts_table_1_in_counts['value'].isin(only_table_1_in)].set_index('value')
            for amount in only_table_1_df.index:
                table_1_count = only_table_1_df.loc[amount, 'count_table_1']
                fill_in_xlsx.append([[amount, 4], [table_1_count, 3], [0, 3], [table_1_count, 3]])
                height_list.append(20)

        # 第三部分：只在 table 2 中出现的金额
        if len(only_table_2_in) > 0:            
            only_table_2_df = amounts_table_2_in_counts[amounts_table_2_in_counts['value'].isin(only_table_2_in)].set_index('value')
            for amount in only_table_2_df.index:
                table_2_count = only_table_2_df.loc[amount, 'count_table_2']
                fill_in_xlsx.append([[amount, 4], [0, 3], [table_2_count, 3], [-table_2_count, 3]])
                height_list.append(20)

        # 填写.xlsx文件
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_cell(file_path, fill_in_xlsx)

        # 明细表
        non_zero_diff_reset_in = non_zero_diff_in.reset_index()
        # 数值分层对比_in_共同数值但计数不同_table_1
        non_zero_diff_reset_in = non_zero_diff_reset_in.rename(columns={'value': table_1_in_df_col})
        if table_1_in_out_mode == '双列模式':
            fill_df = pd.merge(table_1_in_df, non_zero_diff_reset_in[[table_1_in_df_col]], on=table_1_in_df_col, how='inner')
        elif table_1_in_out_mode == '+/-单列模式':
            fill_df = pd.merge(table_1_in_df, non_zero_diff_reset_in[[table_1_in_df_col]], on=table_1_in_df_col, how='inner')
        elif table_1_in_out_mode == '标识列单列模式':
            fill_df = pd.merge(table_1_in_df, non_zero_diff_reset_in[[table_1_in_df_col]], on=table_1_in_df_col, how='inner')
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet(file_path, '共同数值但计数不同table_in_1', fill_df)

        # 数值分层对比_in_共同数值但计数不同_table_2
        non_zero_diff_reset_in = non_zero_diff_reset_in.rename(columns={table_1_in_df_col: table_2_in_df_col})
        if table_2_in_out_mode == '双列模式':
            fill_df = pd.merge(table_2_in_df, non_zero_diff_reset_in[[table_2_in_df_col]], on=table_2_in_df_col, how='inner')
        elif table_2_in_out_mode == '+/-单列模式':
            fill_df = pd.merge(table_2_in_df, non_zero_diff_reset_in[[table_2_in_df_col]], on=table_2_in_df_col, how='inner')
        elif table_2_in_out_mode == '标识列单列模式':
            fill_df = pd.merge(table_2_in_df, non_zero_diff_reset_in[[table_2_in_df_col]], on=table_2_in_df_col, how='inner')
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet(file_path, '共同数值但计数不同table_in_2', fill_df)

        # 数值分层对比_in_只在 table 1 中出现的数值
        if table_1_in_out_mode == '双列模式':
            fill_df = table_1_in_df[table_1_in_df[table_1_in_df_col].isin(only_table_1_in)]
        elif table_1_in_out_mode == '+/-单列模式':
            fill_df = table_1_in_df[table_1_in_df[table_1_in_df_col].isin(only_table_1_in)]
        elif table_1_in_out_mode == '标识列单列模式':
            fill_df = table_1_in_df[table_1_in_df[table_1_in_df_col].isin(only_table_1_in)]
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet(file_path, '只在table_in_1中出现的数值', fill_df)

        # 数值分层对比_in_只在 table 2 中出现的数值
        if table_2_in_out_mode == '双列模式':
            fill_df = table_2_in_df[table_2_in_df[table_2_in_df_col].isin(only_table_2_in)]
        elif table_2_in_out_mode == '+/-单列模式':
            fill_df = table_2_in_df[table_2_in_df[table_2_in_df_col].isin(only_table_2_in)]
        elif table_2_in_out_mode == '标识列单列模式':
            fill_df = table_2_in_df[table_2_in_df[table_2_in_df_col].isin(only_table_2_in)]
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet(file_path, '只在table_in_2中出现的数值', fill_df)

        # output
        h = 6 + len(non_zero_diff_in) + len(only_table_1_in) + len(only_table_2_in) + 3
        height_list = [25,]
        fill_in_xlsx = [['数值合计对比', h, 1],
                        height_list,
                        [20,                  20,                20,                20               ],
                        [['output_value', 1], [table_1_name, 1], [table_2_name, 1], ['difference', 1]],
                        ]

        # 第一部分：共同金额但计数不同的情况
        if len(non_zero_diff_out) > 0:            
            for amount in non_zero_diff_out.index:
                table_1_count = comparison_out.loc[amount, 'count_table_1']
                table_2_count = comparison_out.loc[amount, 'count_table_2']
                diff_desc = table_1_count - table_2_count
                fill_in_xlsx.append([[amount, 4], [table_1_count, 3], [table_2_count, 3], [diff_desc, 3]])
                height_list.append(20)

        # 第二部分：只在 table 1 中出现的金额
        if len(only_table_1_out) > 0:            
            only_table_1_df = amounts_table_1_out_counts[amounts_table_1_out_counts['value'].isin(only_table_1_out)].set_index('value')
            for amount in only_table_1_df.index:
                table_1_count = only_table_1_df.loc[amount, 'count_table_1']
                fill_in_xlsx.append([[amount, 4], [table_1_count, 3], [0, 3], [table_1_count, 3]])
                height_list.append(20)

        # 第三部分：只在 table 2 中出现的金额
        if len(only_table_2_out) > 0:
            only_table_2_df = amounts_table_2_out_counts[amounts_table_2_out_counts['value'].isin(only_table_2_out)].set_index('value')
            for amount in only_table_2_df.index:
                table_2_count = only_table_2_df.loc[amount, 'count_table_2']
                fill_in_xlsx.append([[amount, 4], [0, 3], [table_2_count, 3], [-table_2_count, 3]])
                height_list.append(20)

        # 填写.xlsx文件
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_cell(file_path, fill_in_xlsx)

        non_zero_diff_reset_out = non_zero_diff_out.reset_index()

        # 数值分层对比_out_共同数值但计数不同_table_1
        non_zero_diff_reset_out = non_zero_diff_reset_out.rename(columns={'value': table_1_out_df_col})
        if table_1_in_out_mode == '双列模式':
            fill_df = pd.merge(table_1_out_df, non_zero_diff_reset_out[[table_1_out_df_col]], on=table_1_out_df_col, how='inner')
        elif table_1_in_out_mode == '+/-单列模式':
            fill_df = pd.merge(table_1_out_df, non_zero_diff_reset_out[[table_1_out_df_col]], on=table_1_out_df_col, how='inner')
        elif table_1_in_out_mode == '标识列单列模式':
            fill_df = pd.merge(table_1_out_df, non_zero_diff_reset_out[[table_1_out_df_col]], on=table_1_out_df_col, how='inner')
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet(file_path, '共同数值但计数不同table_out_1', fill_df)

        # 数值分层对比_out_共同数值但计数不同_table_2
        non_zero_diff_reset_out = non_zero_diff_reset_out.rename(columns={table_1_out_df_col: table_2_out_df_col})
        if table_2_in_out_mode == '双列模式':
            fill_df = pd.merge(table_2_out_df, non_zero_diff_reset_out[[table_2_out_df_col]], on=table_2_out_df_col, how='inner')
        elif table_2_in_out_mode == '+/-单列模式':
            fill_df = pd.merge(table_2_out_df, non_zero_diff_reset_out[[table_2_out_df_col]], on=table_2_out_df_col, how='inner')
        elif table_2_in_out_mode == '标识列单列模式':
            fill_df = pd.merge(table_2_out_df, non_zero_diff_reset_out[[table_2_out_df_col]], on=table_2_out_df_col, how='inner')
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet(file_path, '共同数值但计数不同table_out_2', fill_df)

        # 数值分层对比_out_只在 table 1 中出现的数值
        if table_1_in_out_mode == '双列模式':
            fill_df = table_1_out_df[table_1_out_df[table_1_out_df_col].isin(only_table_1_out)]
        elif table_1_in_out_mode == '+/-单列模式':
            fill_df = table_1_out_df[table_1_out_df[table_1_out_df_col].isin(only_table_1_out)]
        elif table_1_in_out_mode == '标识列单列模式':
            fill_df = table_1_out_df[table_1_out_df[table_1_out_df_col].isin(only_table_1_out)]
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet(file_path, '只在table_out_1中出现的数值', fill_df)

        # 数值分层对比_out_只在 table 2 中出现的数值
        if table_2_in_out_mode == '双列模式':
            fill_df = table_2_out_df[table_2_out_df[table_2_out_df_col].isin(only_table_2_out)]
        elif table_2_in_out_mode == '+/-单列模式':
            fill_df = table_2_out_df[table_2_out_df[table_2_out_df_col].isin(only_table_2_out)]
        elif table_2_in_out_mode == '标识列单列模式':
            fill_df = table_2_out_df[table_2_out_df[table_2_out_df_col].isin(only_table_2_out)]
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet(file_path, '只在table_out_2中出现的数值', fill_df)


        # 对比模式三：按item分类对比
        fill_df = (table_1_in_df.assign(**{table_1_item: table_1_in_df[table_1_item].fillna('空白')})
                                .groupby(table_1_item, as_index=False)[[table_1_in_df_col]].sum())
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet_cell(file_path, 'table_in按item分类汇总', fill_df, 1, 1)
        fill_df = (table_2_in_df.assign(**{table_2_item: table_2_in_df[table_2_item].fillna('空白')})
                                .groupby(table_2_item, as_index=False)[[table_2_in_df_col]].sum())
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet_cell(file_path, 'table_in按item分类汇总', fill_df, 1, 4)

        fill_df = (table_1_out_df.assign(**{table_1_item: table_1_out_df[table_1_item].fillna('空白')})
                                .groupby(table_1_item, as_index=False)[[table_1_out_df_col]].sum())
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet_cell(file_path, 'table_out按item分类汇总', fill_df, 1, 1)
        fill_df = (table_2_out_df.assign(**{table_2_item: table_2_out_df[table_2_item].fillna('空白')})
                                .groupby(table_2_item, as_index=False)[[table_2_out_df_col]].sum())
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet_cell(file_path, 'table_out按item分类汇总', fill_df, 1, 4)


        # 对比模式四：按时序排列
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet_cell(file_path, 'table_in按时序排列', table_1_in_df, 1, 1)
        c_w = len(table_1_in_df.columns) + 2
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet_cell(file_path, 'table_in按时序排列', table_2_in_df, 1, c_w)

        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet_cell(file_path, 'table_out按时序排列', table_1_out_df, 1, 1)
        c_w = len(table_1_out_df.columns) + 2
        fill_text += openxl_xlsx_tool_class_example.fill_in_xlsx_sheet_cell(file_path, 'table_out按时序排列', table_2_out_df, 1, c_w)


        return fill_text