# pd_concat_fun.py

# DataFrame 拼接 concat

import pandas as pd

def pd_concat(df_list):

    try:

        df = pd.concat(df_list, axis=0)

    except:

        pass

    return df