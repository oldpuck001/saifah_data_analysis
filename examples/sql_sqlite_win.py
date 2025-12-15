# sql_sqlite_win.py

import sys
import os

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.saifah_data_analysis import ui_tk_root

title = 'SQLite'
geometry = '1280x720+50+50'
control_frame_n = 2
control_frame_config = [
                        {'name':          'create_select_sqlite_database',
                        'button_name':    '创建/选择SQLite数据库',
                        'function_name':  '',
                        'function_para':  ''},

                        {'name':          'sql_sqlite_win',
                        'button_name':    '',
                        'function_name':  '',
                        'function_para':  0},
                       ]

app = ui_tk_root.App(title=title, geometry=geometry, control_frame_n=control_frame_n, control_frame_config=control_frame_config)
app.root.mainloop()