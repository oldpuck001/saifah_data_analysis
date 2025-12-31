# input_output_value_check_sqlite.py

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from . import fill_area_text
from ..DataFrame_tool_pd import pd_DataFrame_tool_class
from ..openxl_xlsx_tool import openxl_xlsx_tool_class
from ..sql_tool_sqlite import sql_sqlite_tool_class

openxl_xlsx_tool_class_example = openxl_xlsx_tool_class()
pd_DataFrame_tool_class_example = pd_DataFrame_tool_class()
sql_qlite_tool_example = sql_sqlite_tool_class()

class input_output_value_check_sqlite_modular:

    def input_output_value_check_sqlite_frame(self, root, control_frame_config, text_area):

        frame_result = tk.Frame(root)
        frame_result.pack(side=tk.BOTTOM, fill=tk.BOTH, padx=5, pady=5)

        # table_1
        frame_1 = tk.Frame(frame_result)
        frame_1.pack(side=tk.TOP, fill=tk.BOTH)

        tk.Label(frame_1, text='Table 1 Name', width=10, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options = []
        frame_1.combobox_table_1_name = ttk.Combobox(frame_1, values=options, height=10, width=19, state='readonly')
        frame_1.combobox_table_1_name.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_1, text='Table 1 Time Series', width=14, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_time_series = ['正向', '反向']
        frame_1.combobox_table_1_time_series = ttk.Combobox(frame_1, values=options_time_series, height=10, width=15, state='readonly')
        frame_1.combobox_table_1_time_series.set(options_time_series[0])
        frame_1.combobox_table_1_time_series.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_1, text='Table 1 Item', width=9, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_1.combobox_table_1_item = ttk.Combobox(frame_1, values=options, height=10, width=20, state='readonly')
        frame_1.combobox_table_1_item.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_1, text='Table 1 In/Out Mode', width=14, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_in_out_mode = ['双列模式', '+/-单列模式', '标识列单列模式']
        frame_1.combobox_table_1_in_out_mode = ttk.Combobox(frame_1, values=options_in_out_mode, height=10, width=15, state='readonly')
        frame_1.combobox_table_1_in_out_mode.set(options_in_out_mode[0])
        frame_1.combobox_table_1_in_out_mode.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        frame_result.frame_1 = frame_1

        frame_2 = tk.Frame(frame_result)
        frame_2.pack(side=tk.TOP, fill=tk.BOTH)

        tk.Label(frame_2, text='Table 1 In Col', width=10, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_2.combobox_table_1_in_col = ttk.Combobox(frame_2, values=options, height=10, width=19, state='readonly')
        frame_2.combobox_table_1_in_col.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_2, text='Table 1 Out Col', width=12, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_2.combobox_table_1_out_col = ttk.Combobox(frame_2, values=options, height=10, width=17, state='readonly')
        frame_2.combobox_table_1_out_col.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_2, text='Reserved', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_2.combobox_table_1_reserved_1 = ttk.Combobox(frame_2, values=options, height=10, width=21, state='readonly')
        frame_2.combobox_table_1_reserved_1.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_2, text='Reserved', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_2.combobox_table_1_reserved_2 = ttk.Combobox(frame_2, values=options, height=10, width=21, state='readonly')
        frame_2.combobox_table_1_reserved_2.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        frame_result.frame_2 = frame_2

        frame_3 = tk.Frame(frame_result)
        frame_3.pack(side=tk.TOP, fill=tk.BOTH)

        tk.Label(frame_3, text='Table 1 In/Out Col', width=13, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.combobox_table_1_in_out_col = ttk.Combobox(frame_3, values=options, height=10, width=16, state='readonly')
        frame_3.combobox_table_1_in_out_col.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_3, text='Table 1 In Label', width=12, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.combobox_table_1_in_label = ttk.Combobox(frame_3, values=options, height=10, width=17, state='readonly')
        frame_3.combobox_table_1_in_label.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_3, text='Table 1 Out Label', width=13, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.combobox_table_1_out_label = ttk.Combobox(frame_3, values=options, height=10, width=16, state='readonly')
        frame_3.combobox_table_1_out_label.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_3, text='Table 1 In/Out Value', width=14, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.combobox_table_1_in_out_value = ttk.Combobox(frame_3, values=options, height=10, width=15, state='readonly')
        frame_3.combobox_table_1_in_out_value.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        frame_result.frame_3 = frame_3


        # table 2
        frame_4 = tk.Frame(frame_result)
        frame_4.pack(side=tk.TOP, fill=tk.BOTH)

        tk.Label(frame_4, text='Table 2 Name', width=10, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options = []
        frame_4.combobox_table_2_name = ttk.Combobox(frame_4, values=options, height=10, width=19, state='readonly')
        frame_4.combobox_table_2_name.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_4, text='Table 2 Time Series', width=14, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_time_series = ['正向', '反向']
        frame_4.combobox_table_2_time_series = ttk.Combobox(frame_4, values=options_time_series, height=10, width=15, state='readonly')
        frame_4.combobox_table_2_time_series.set(options_time_series[0])
        frame_4.combobox_table_2_time_series.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_4, text='Table 2 Item', width=9, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_4.combobox_table_2_item = ttk.Combobox(frame_4, values=options, height=10, width=20, state='readonly')
        frame_4.combobox_table_2_item.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_4, text='Table 2 In/Out Mode', width=14, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_in_out_mode = ['双列模式', '+/-单列模式', '标识列单列模式']
        frame_4.combobox_table_2_in_out_mode = ttk.Combobox(frame_4, values=options_in_out_mode, height=10, width=15, state='readonly')
        frame_4.combobox_table_2_in_out_mode.set(options_in_out_mode[0])
        frame_4.combobox_table_2_in_out_mode.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        frame_result.frame_4 = frame_4

        frame_5 = tk.Frame(frame_result)
        frame_5.pack(side=tk.TOP, fill=tk.BOTH)

        tk.Label(frame_5, text='Table 2 In Col', width=10, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.combobox_table_2_in_col = ttk.Combobox(frame_5, values=options, height=10, width=19, state='readonly')
        frame_5.combobox_table_2_in_col.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_5, text='Table 2 Out Col', width=12, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.combobox_table_2_out_col = ttk.Combobox(frame_5, values=options, height=10, width=17, state='readonly')
        frame_5.combobox_table_2_out_col.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_5, text='Reserved', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.combobox_table_2_reserved_1 = ttk.Combobox(frame_5, values=options, height=10, width=21, state='readonly')
        frame_5.combobox_table_2_reserved_1.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_5, text='Reserved', width=8, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.combobox_table_2_reserved_2 = ttk.Combobox(frame_5, values=options, height=10, width=21, state='readonly')
        frame_5.combobox_table_2_reserved_2.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        frame_result.frame_5 = frame_5

        frame_6 = tk.Frame(frame_result)
        frame_6.pack(side=tk.TOP, fill=tk.BOTH)

        tk.Label(frame_6, text='Table 2 In/Out Col', width=13, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_6.combobox_table_2_in_out_col = ttk.Combobox(frame_6, values=options, height=10, width=16, state='readonly')
        frame_6.combobox_table_2_in_out_col.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_6, text='Table 2 In Label', width=12, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_6.combobox_table_2_in_label = ttk.Combobox(frame_6, values=options, height=10, width=17, state='readonly')
        frame_6.combobox_table_2_in_label.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_6, text='Table 2 Out Label', width=13, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_6.combobox_table_2_out_label = ttk.Combobox(frame_6, values=options, height=10, width=16, state='readonly')
        frame_6.combobox_table_2_out_label.pack(side=tk.LEFT, padx=(5, 12), pady=5)

        tk.Label(frame_6, text='Table 2 In/Out Value', width=14, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_6.combobox_table_2_in_out_value = ttk.Combobox(frame_6, values=options, height=10, width=15, state='readonly')
        frame_6.combobox_table_2_in_out_value.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        frame_result.frame_6 = frame_6


        # 按钮行
        frame_7 = tk.Frame(frame_result)
        frame_7.pack(side=tk.TOP, fill=tk.BOTH)
        tk.Button(frame_7, text='Read SQL Table List', command=lambda: self.button_read_fun(root, control_frame_config, frame_1, frame_4, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_7, text='Read SQL Table 1 List', command=lambda: self.button_table_col_fun(root, control_frame_config, frame_1.combobox_table_1_name,
                                                                                                   frame_1.combobox_table_1_item, frame_2.combobox_table_1_in_col,
                                                                                                   frame_2.combobox_table_1_out_col, frame_3.combobox_table_1_in_out_col,
                                                                                                   frame_3.combobox_table_1_in_out_value, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_7, text='Read SQL Table 1 Label', command=lambda: self.button_table_label_fun(root, control_frame_config, frame_1.combobox_table_1_name,
                                                                                                      frame_3.combobox_table_1_in_out_col, frame_3.combobox_table_1_in_label,
                                                                                                      frame_3.combobox_table_1_out_label, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_7, text='Read SQL Table 2 List', command=lambda: self.button_table_col_fun(root, control_frame_config, frame_4.combobox_table_2_name,
                                                                                                   frame_4.combobox_table_2_item, frame_5.combobox_table_2_in_col,
                                                                                                   frame_5.combobox_table_2_out_col, frame_6.combobox_table_2_in_out_col,
                                                                                                   frame_6.combobox_table_2_in_out_value, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_7, text='Read SQL Table 2 Label', command=lambda: self.button_table_label_fun(root, control_frame_config, frame_4.combobox_table_2_name,
                                                                                                      frame_6.combobox_table_2_in_out_col, frame_6.combobox_table_2_in_label,
                                                                                                      frame_6.combobox_table_2_out_label, text_area), width=18).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_7, text='Comparison Data', command=lambda: self.button_comparison_fun(root, control_frame_config, frame_1, frame_2, frame_3,
                                                                                              frame_4, frame_5, frame_6, text_area), width=20).pack(side=tk.LEFT, padx=5, pady=5)

        frame_result.frame_7 = frame_7


    # 读取数据库表名单
    def button_read_fun(self, root, control_frame_config, frame_1, frame_4, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()

        if sql_path:
            result_info = sql_qlite_tool_example.get_all_tables(sql_path)
            if result_info[0]:
                option_table_name_list = result_info[1]
                frame_1.combobox_table_1_name['values'] = option_table_name_list
                if option_table_name_list:
                    frame_1.combobox_table_1_name.set(option_table_name_list[0])
                frame_1.combobox_table_1_name.config(state='readonly')

                frame_4.combobox_table_2_name['values'] = option_table_name_list
                if option_table_name_list:
                    frame_4.combobox_table_2_name.set(option_table_name_list[0])
                frame_4.combobox_table_2_name.config(state='readonly')
                fill_text += f'SQL table name list import successful!\n'
            else:
                fill_text += 'SQL table name list import failed!\n'
                fill_text += result_info[1]
        else:
            fill_text += f'Please check that the database path is not empty!\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # 读取列名列表
    def button_table_col_fun(self, root, control_frame_config, combobox_table_name, combobox_table_item, combobox_table_in_col, combobox_table_out_col,
                             combobox_table_in_out_col, combobox_table_in_out_value, text_area):
        
        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()
        table_name = combobox_table_name.get()

        if sql_path and table_name:

            result_info = sql_qlite_tool_example.get_table_columns(sql_path, table_name)
            if result_info[0]:
                option_col_name_list = result_info[1][1:]

                combobox_table_item['values'] = option_col_name_list
                if option_col_name_list[0]:
                    combobox_table_item.set(option_col_name_list[0])
                combobox_table_item.config(state='readonly')

                combobox_table_in_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    combobox_table_in_col.set(option_col_name_list[0])
                combobox_table_in_col.config(state='readonly')

                combobox_table_out_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    combobox_table_out_col.set(option_col_name_list[0])
                combobox_table_out_col.config(state='readonly')

                combobox_table_in_out_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    combobox_table_in_out_col.set(option_col_name_list[0])
                combobox_table_in_out_col.config(state='readonly')

                combobox_table_in_out_value['values'] = option_col_name_list
                if option_col_name_list[0]:
                    combobox_table_in_out_value.set(option_col_name_list[0])
                combobox_table_in_out_value.config(state='readonly')

                fill_text += f'SQL table 1 column name list import successful!\n'
            else:
                fill_text += 'SQL table 1 column name list import failed!\n'
                fill_text += result_info[1]
        else:
            if not sql_path:
                fill_text += f'Please check that the database path is not empty!\n'
            if not table_name:
                fill_text += f'Please check that the table name is not empty!\n'

        fill_area_text.text_area_fill(text_area, fill_text)


    # 读取标签
    def button_table_label_fun(self, root, control_frame_config, combobox_table_name, combobox_table_in_out_col, combobox_table_in_label, combobox_table_out_label, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()
        table_name = combobox_table_name.get()
        table_in_out_col = combobox_table_in_out_col.get()

        unique_index_list_cleaned = []

        sql_template_single = """
                                SELECT *
                                FROM {table_name}
                                WHERE {table_in_out_col} IS NOT NULL
                              """

        sql_table_select = sql_template_single.format(table_name=table_name, table_in_out_col=table_in_out_col)

        result_info_table = sql_qlite_tool_example.execute_query(sql_path, sql_table_select)

        if result_info_table[0]:
            sql_table_1_df = result_info_table[1]
            fill_text += f'{table_name} read in/out label successful!\n'
        else:
            fill_text += result_info_table[1]

        unique_index_np = sql_table_1_df[table_in_out_col].unique()
        unique_index_list = unique_index_np.tolist()

        unique_index_list = [item for item in unique_index_list if item is not None]

        for item in unique_index_list:
            if not isinstance(item, (int, float)):                              # 检查item是否是int或float类型
                unique_index_list_cleaned.append(item.replace('\t', ''))        # 移除字符串中的所有制表符（\t）
            else:
                unique_index_list_cleaned.append(item)

        combobox_table_in_label['values'] = unique_index_list_cleaned
        if unique_index_list_cleaned[0]:
            combobox_table_in_label.set(unique_index_list_cleaned[0])
        combobox_table_in_label.config(state='readonly')
        
        combobox_table_out_label['values'] = unique_index_list_cleaned
        if unique_index_list_cleaned[0]:
            combobox_table_out_label.set(unique_index_list_cleaned[0])
        combobox_table_out_label.config(state='readonly')

        fill_area_text.text_area_fill(text_area, fill_text)


    # 数据对比
    def button_comparison_fun(self, root, control_frame_config, frame_1, frame_2, frame_3, frame_4, frame_5, frame_6, text_area):

        fill_text = ''
        n_config = control_frame_config['function_para']
        sql_path = root.control_frame_list[n_config].entry_widget.get()

        table_1_name = frame_1.combobox_table_1_name.get()
        table_1_time_series = frame_1.combobox_table_1_time_series.get()
        table_1_item = frame_1.combobox_table_1_item.get()
        table_1_in_out_mode = frame_1.combobox_table_1_in_out_mode.get()

        table_1_in_col = frame_2.combobox_table_1_in_col.get()
        table_1_out_col = frame_2.combobox_table_1_out_col.get()

        table_1_in_out_col = frame_3.combobox_table_1_in_out_col.get()
        table_1_in_label = frame_3.combobox_table_1_in_label.get()
        table_1_out_label = frame_3.combobox_table_1_out_label.get()
        table_1_in_out_value = frame_3.combobox_table_1_in_out_value.get()

        table_2_name = frame_4.combobox_table_2_name.get()
        table_2_time_series = frame_4.combobox_table_2_time_series.get()
        table_2_item = frame_4.combobox_table_2_item.get()
        table_2_in_out_mode = frame_4.combobox_table_2_in_out_mode.get()

        table_2_in_col = frame_5.combobox_table_2_in_col.get()
        table_2_out_col = frame_5.combobox_table_2_out_col.get()

        table_2_in_out_col = frame_6.combobox_table_2_in_out_col.get()
        table_2_in_label = frame_6.combobox_table_2_in_label.get()
        table_2_out_label = frame_6.combobox_table_2_out_label.get()
        table_2_in_out_value = frame_6.combobox_table_2_in_out_value.get()

        file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])

        sheet_name_list = ['数值合计对比',
                           '共同数值但计数不同table_in_1',  '共同数值但计数不同table_in_2',
                           '只在table_in_1中出现的数值',    '只在table_in_2中出现的数值',
                           '共同数值但计数不同table_out_1', '共同数值但计数不同table_out_2',
                           '只在table_out_1中出现的数值',   '只在table_out_2中出现的数值',
                           'table_in按item分类汇总',       'table_out按item分类汇总',
                           'table_in按时序排列',           'table_out按时序排列'                          
                           ]
        fill_text += openxl_xlsx_tool_class_example.create_new_xlsx(file_path, sheet_name_list)

        sql_template = """
                        SELECT *
                        FROM {table_name}
                      """

        # table_1
        sql_table_1_select = sql_template.format(table_name=table_1_name, table_in_out_value=table_1_in_out_value)
        result_info_table_1 = sql_qlite_tool_example.execute_query(sql_path, sql_table_1_select)

        if result_info_table_1[0]:
            sql_table_1_df = result_info_table_1[1]
            if table_1_time_series == '反向':
                sql_table_1_df[:] = sql_table_1_df.iloc[::-1].values
            fill_text += f'{table_1_name} read successful!\n'
        else:
            fill_text += result_info_table_1[1]

        # table_2
        sql_table_2_select = sql_template.format(table_name=table_2_name, table_in_out_value=table_2_in_out_value)
        result_info_table_2 = sql_qlite_tool_example.execute_query(sql_path, sql_table_2_select)

        if result_info_table_2[0]:
            sql_table_2_df = result_info_table_2[1]
            if table_2_time_series == '反向':
                sql_table_2_df[:] = sql_table_2_df.iloc[::-1].values
            fill_text += f'{table_2_name} read successful!\n'
        else:
            fill_text += result_info_table_2[1]

        # 对比模式一：对比交易记录的数量与金额
        fill_text += pd_DataFrame_tool_class_example.input_output_value_check_sum(table_1_name, table_1_in_out_mode,
                                                                                  sql_table_1_df, table_1_in_col, table_1_out_col,
                                                                                  table_1_in_out_value, table_1_in_out_col,
                                                                                  table_1_in_label, table_1_out_label,
                                                                                  table_2_name, table_2_in_out_mode,
                                                                                  sql_table_2_df, table_2_in_col, table_2_out_col,
                                                                                  table_2_in_out_value, table_2_in_out_col,
                                                                                  table_2_in_label, table_2_out_label,
                                                                                  file_path)

        # 对比模式二：数值分层对比 + 对比模式三：按item分类汇总 + 对比模式四：按时序排列
        fill_text += pd_DataFrame_tool_class_example.input_output_value_check_layering(table_1_name, table_2_name,
                                                                                       table_1_in_out_mode, sql_table_1_df, table_1_in_col, table_1_out_col,
                                                                                       table_1_in_out_value, table_1_in_out_col, table_1_in_label, table_1_out_label,
                                                                                       table_2_in_out_mode, sql_table_2_df, table_2_in_col, table_2_out_col,
                                                                                       table_2_in_out_value, table_2_in_out_col, table_2_in_label, table_2_out_label,
                                                                                       table_1_item, table_2_item,
                                                                                       file_path)

        fill_area_text.text_area_fill(text_area, fill_text)