# sheet_subtotals.py

import sys
import os

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.saifah_data_analysis import ui_tk_root

title = '分类汇总表格'
geometry = '960x540+50+50'
control_frame_n = 1
control_frame_config = [
                        {'name':          'sheet_subtotals',
                        'button_name':    ['导入', '生成'],
                        'function_name':  '',
                        'function_para':  ''},
                       ]

app = ui_tk_root.App(title=title, geometry=geometry, control_frame_n=control_frame_n, control_frame_config=control_frame_config)
app.root.mainloop()