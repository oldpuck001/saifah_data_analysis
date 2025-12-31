# sheets_script.py

import sys
import os

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.saifah_data_analysis import ui_tk_root

title = '电子表格拼接'
geometry = '1280x720+50+50'
control_frame_n = 1
control_frame_config = [
                        {'name':          'sheets_script',
                        'button_name':    [10, '添加', '删除', '导出'],
                        'function_name':  '',
                        'function_para':  ''},
                       ]

app = ui_tk_root.App(title=title, geometry=geometry, control_frame_n=control_frame_n, control_frame_config=control_frame_config)
app.root.mainloop()