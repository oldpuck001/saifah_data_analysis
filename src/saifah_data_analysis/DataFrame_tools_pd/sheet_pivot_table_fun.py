# sheet_pivot_table_fun.py

import pandas as pd
from ..DataFrame_tools_pd import read_xlsx_xls_csv_txt_fun
from ..DataFrame_tools_pd import df_cleaning_fun

def sheet_pivot_table(file_path, sheet_name, row_value, column_value, subtotal_value):

    result = read_xlsx_xls_csv_txt_fun.read_xlsx_xls_csv_txt(original_file_path=file_path, sheet_name=sheet_name)

    if result[0]:
        df = result[2]
        
        df_cleaned = df_cleaning_fun.df_cleaning(df, subtotal_value, 'to_float')

        df_export = pd.pivot_table(df.iloc[0:], values=subtotal_value, index=row_value, columns=column_value, aggfunc='sum')

        return [True, df_export]
    
    else:
        return [False]






