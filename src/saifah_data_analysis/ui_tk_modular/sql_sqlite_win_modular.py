# sql_sqlite_win_modular.py

import os
import shutil
import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk
from tkinter import messagebox
import subprocess
from . import fill_area_text
from ..DataFrame_tool_pd import pd_DataFrame_tool_class
from ..sql_tool_sqlite import sql_sqlite_tool_class

pd_DataFrame_tool_example = pd_DataFrame_tool_class()
sql_sqlite_tool_exampoe = sql_sqlite_tool_class()

class sql_sqlite_win_modular_class:

    def sql_sqlite_win_modular_frame(self, root, control_frame_config, text_area):

        frame_result = tk.Frame(root)
        frame_result.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # SQL指令输入区
        frame_1 = tk.Frame(frame_result)
        frame_1.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        sql_command_text_area = ScrolledText(frame_1, height=16)
        sql_command_text_area.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
        frame_1.sql_command_text_area = sql_command_text_area
        frame_result.frame_1 = frame_1

        # 第一行按钮
        frame_2 = tk.Frame(frame_result)
        frame_2.pack(side=tk.TOP, fill=tk.X)
        tk.Button(frame_2, text='执行指令', command=lambda: self.sql_sqlite_button(root, control_frame_config, sql_command_text_area, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_2, text='备份数据库文件', command=lambda: self.sql_sqlite_backup(root, control_frame_config, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_2, text='查询并预览', command=lambda: self.sql_sqlite_review_button(root, control_frame_config, sql_command_text_area, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_2, text='查询并导出xlsx', command=lambda: self.sql_sqlite_export_button(root, control_frame_config, sql_command_text_area, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_2, text='获取所有表名', command=lambda: self.sql_sqlite_table_button(root, control_frame_config, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_2, text='获取表信息', command=lambda: self.sql_sqlite_table_info_button(root, control_frame_config, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_2 = frame_2

        # 第二行按钮
        frame_3 = tk.Frame(frame_result)
        frame_3.pack(side=tk.TOP, fill=tk.X)
        tk.Button(frame_3, text='检查表是否存在', command=lambda: self.find_table_name(root, control_frame_config, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_3, text='清除SQL指令区', command=lambda:self.scrolledtext_clear(sql_command_text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_3, text='清除操作记录区', command=lambda:self.text_area_clear(text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_3, text='删除单个Table', command=lambda:self.del_single_table(root, control_frame_config, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_3, text='删除所有Table', command=lambda:self.del_all_table(root, control_frame_config, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_3, text='SQL指令帮助', command=lambda: self.button_manual_fun(), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        frame_result.frame_3 = frame_3


    # 执行SQL指令按钮函数
    def sql_sqlite_button(self, root, control_frame_config, sql_command_text_area, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()
        sql_command = sql_command_text_area.get('1.0', 'end-1c')

        fill_text = sql_sqlite_tool_exampoe.sql_sqlite_command(sql_path, sql_command)

        fill_area_text.text_area_fill(text_area, fill_text)


    # 备份数据库文件
    def sql_sqlite_backup(self, root, control_frame_config, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()
        folder_path = os.path.dirname(sql_path)
        filename = os.path.basename(sql_path)
        name, ext = os.path.splitext(filename)
        target_path = os.path.join(folder_path, f'{name}_1{ext}')

        counter = 2
        while os.path.exists(target_path):
            new_filename = f'{name}_{counter}{ext}'
            target_path = os.path.join(folder_path, new_filename)
            counter += 1

        shutil.copy(sql_path, target_path)

        fill_text = f'还原点建立成功，文件路径：{target_path}\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # 查询并预览
    def sql_sqlite_review_button(self, root, control_frame_config, sql_command_text_area, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()
        folder_path = os.path.dirname(sql_path)
        filename = os.path.basename(sql_path)
        name, ext = os.path.splitext(filename)
        target_path = os.path.join(folder_path, f'{name}_review.xlsx')
        sql_command = sql_command_text_area.get('1.0', 'end-1c')
        sql_clean = sql_command.strip()
        sql_type = sql_clean.lower().split()[0]
        if sql_type == 'select':
            result_info = sql_sqlite_tool_exampoe.execute_query(sql_path, sql_command)
            if result_info[0]:
                result_df = result_info[1]
                result_df.to_excel(target_path)
                fill_text = f'\n查询指令：\n{sql_command}\n执行完毕！\n'
            else:            
                fill_text = f'\n查询指令：\n{sql_command}\n执行失败！\n错误信息：{str(result_info[1])}\n'
        else:
                fill_text = f'\n查询指令：\n{sql_command}\n，不是查询指令，请输入查询指令。\n'

        fill_area_text.text_area_fill(text_area, fill_text)
        subprocess.run(['open', target_path])


    # 查询并导出xlsx
    def sql_sqlite_export_button(self, root, control_frame_config, sql_command_text_area, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()

        target_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])

        sql_command = sql_command_text_area.get('1.0', 'end-1c')
        sql_clean = sql_command.strip()
        sql_type = sql_clean.lower().split()[0]
        if sql_type == 'select':
            result_info = sql_sqlite_tool_exampoe.execute_query(sql_path, sql_command)
            if result_info[0]:
                result_df = result_info[1]
                result_df.to_excel(target_path)
                fill_text = f'\n查询指令：\n{sql_command}\n执行完毕！\n导出文件路径：{target_path}\n'
            else:            
                fill_text = f'\n查询指令：\n{sql_command}\n执行失败！\n错误信息：{str(result_info[1])}\n'
        else:
                fill_text = f'\n查询指令：\n{sql_command}\n，不是查询指令，请输入查询指令。\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # 查询所有工作表按钮函数
    def sql_sqlite_table_button(self, root, control_frame_config, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()

        result_info = sql_sqlite_tool_exampoe.get_all_tables(sql_path)
        if result_info[0]:
            fill_text += ', '.join(result_info[1])
            if fill_text:
                fill_text += '\n'
            else:
                fill_text += 'There is no table in the database.\n'
        else:
            fill_text += result_info[1]

        fill_area_text.text_area_fill(text_area, fill_text)


    # 获取表信息按钮函数
    def sql_sqlite_table_info_button(self, root, control_frame_config, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()

        result_info = sql_sqlite_tool_exampoe.get_all_tables(sql_path)

        if result_info[0]:
            tables_list = result_info[1]
            table_name = self.select_table(root, tables_list)
            result_text = sql_sqlite_tool_exampoe.get_table_info(sql_path, table_name)
            if result_text[0]:
                fill_text += 'Column Info:\n'
                fill_text += (f"{'CID':<4} {'NAME':<15} {'TYPE':<8} {'NOT NULL':<9} {'DEFAULT':<8} {'PK':<3}\n")
                fill_text += (f"{'-'*4} {'-'*15} {'-'*8} {'-'*9} {'-'*8} {'-'*3}\n")
                for cid, name, col_type, notnull, default, pk in result_text[1][0]:
                    fill_text += (f'{cid:<4} {name:<15} {col_type:<8} {notnull:<9} {str(default):<8} {pk:<3}\n')
                fill_text += f'\nNumber of Rows: {result_text[1][1]}\n'
            else:
                fill_text += result_text[1]
        else:
            fill_text += result_info[1]

        fill_area_text.text_area_fill(text_area, fill_text)


    # 检查表是否存在按钮函数
    def find_table_name(self, root, control_frame_config, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()

        table_name = self.enter_table(root)
        if table_name:
            result_info = sql_sqlite_tool_exampoe.table_exists(sql_path, table_name)
            fill_text = result_info[1]
        else:
            fill_text += 'No table name entered.\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # 清除ScrolledText内容
    def scrolledtext_clear(self, text_area):

        text_area.delete('1.0', 'end')


    # 清除text_area内容
    def text_area_clear(self, text_area):

        fill_area_text.text_area_clear(text_area)


    # 删除单个table
    def del_single_table(self, root, control_frame_config, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()

        result_info = sql_sqlite_tool_exampoe.get_all_tables(sql_path)

        if result_info[0]:
            tables_list = result_info[1]
            table_name = self.select_table(root, tables_list)

            result_yesno_1 = messagebox.askyesno(message='删除table后，无法恢复，确定要继续执行删除操作吗？')
            if result_yesno_1:
                result_yesno_2 = messagebox.askyesno(message='删除table后，无法恢复，确定要继续执行删除操作吗？')
                if result_yesno_2:
                    result_text = sql_sqlite_tool_exampoe.del_single_table(sql_path, table_name)
                    fill_text += result_text[1]
                else:
                    fill_text += 'The deletion has been cancelled.\n'
            else:
                fill_text += 'The deletion has been cancelled.\n'
        else:
            fill_text += result_info[1]

        fill_area_text.text_area_fill(text_area, fill_text)


    # 删除所有table
    def del_all_table(self, root, control_frame_config, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()

        result_yesno_1 = messagebox.askyesno(message='删除table后，无法恢复，确定要继续执行删除操作吗？')
        if result_yesno_1:
            result_yesno_2 = messagebox.askyesno(message='删除table后，无法恢复，确定要继续执行删除操作吗？')
            if result_yesno_2:
                result_yesno_3 = messagebox.askyesno(message='删除table后，无法恢复，确定要继续执行删除操作吗？')
                if result_yesno_3:
                    result_text = sql_sqlite_tool_exampoe.del_all_tables(sql_path)
                    fill_text += result_text[1]
            else:
                fill_text += 'The deletion has been cancelled.\n'
        else:
            fill_text += 'The deletion has been cancelled.\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # SQL指令帮助按钮函数
    def button_manual_fun(self):

        win = tk.Toplevel()
        win.title("Pandas 参数说明")
        win.geometry("780x600")

        text = ScrolledText(win, wrap=tk.WORD, font=("Arial", 12))
        text.pack(expand=True, fill=tk.BOTH)

        params_text = """
    SQLite 指令说明

    """
        text.insert(tk.END, params_text)
        text.config(state="disabled")


    # 选择 table 子窗口
    def select_table(self, root, tables_list):

        win = tk.Toplevel(root)
        win.title('Select Table')
        win.geometry('300x100+600+300')
        win.resizable(False, False)
        selected = tk.StringVar()
        cb = ttk.Combobox(win, textvariable=selected, values=tables_list, width=50, state='readonly')
        cb.pack(side=tk.TOP, padx=15, pady=(15, 5))
        if tables_list:
            cb.current(0)   # 默认选中第一个
        def confirm():
            win.destroy()
        ttk.Button(win, text='Confirm', command=confirm, width=15).pack(side=tk.TOP, padx=15, pady=(5, 15))
        win.grab_set()
        root.wait_window(win)
        return selected.get()


    # 输入 table 子窗口
    def enter_table(self, root):

        table_name = tk.StringVar()
        win = tk.Toplevel(root)
        win.title('Enter Table Name')
        win.geometry('300x100+600+300')
        win.resizable(False, False)
        entry = tk.Entry(win, textvariable=table_name, width=50)
        entry.pack(side=tk.TOP, padx=15, pady=(15, 5))
        entry.focus_set()   # 自动获得焦点
        def confirm():
            win.destroy()
        ttk.Button(win, text='Confirm', command=confirm, width=15).pack(side=tk.TOP, padx=15, pady=(5, 15))
        win.grab_set()
        root.wait_window(win)
        return table_name.get().strip() or None