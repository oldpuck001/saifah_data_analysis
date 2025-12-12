# pd_DataFrame_tool.py

import os
import pandas as pd

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