

'''
    # 数据库读取核对
    def sql_in_out_check_frame(self, n):
        frame_1 = tk.Frame(self.control_frame_list[n])
        frame_1.pack(side=tk.TOP, padx=5, pady=5)

        tk.Button(frame_1, text='Read SQL Table List', command=lambda n=n: self.button_read_fun(n), width=23).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_1, text='Read SQL Table 1 List', command=lambda n=n: self.button_table_1_fun(n), width=23).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_1, text='Read SQL Table 1 Label', command=lambda n=n: self.button_label_1_fun(n), width=23).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_1, text='Read SQL Table 2 List', command=lambda n=n: self.button_table_2_fun(n), width=23).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_1, text='Read SQL Table 2 Label', command=lambda n=n: self.button_label_2_fun(n), width=23).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_1, text='Comparison Data', command=lambda n=n: self.button_comparison_fun(n), width=23).pack(side=tk.LEFT, padx=5, pady=5)

        self.control_frame_list[n].frame_1 = frame_1

        frame_2 = tk.Frame(self.control_frame_list[n])
        frame_2.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_2, text='Table 1 Name', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options = []
        frame_2.combobox_table_1_name = ttk.Combobox(frame_2, values=options, height=10, width=14, state='readonly')
        frame_2.combobox_table_1_name.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_2, text='Table 1 Time Series', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_time_series = ['正向', '反向']
        frame_2.combobox_table_1_time_series = ttk.Combobox(frame_2, values=options_time_series, height=10, width=14, state='readonly')
        frame_2.combobox_table_1_time_series.set(options_time_series[0])
        frame_2.combobox_table_1_time_series.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_2, text='Table 1 Item', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_2.combobox_table_1_item = ttk.Combobox(frame_2, values=options, height=10, width=14, state='readonly')
        frame_2.combobox_table_1_item.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_2, text='Table 1 In/Out Mode', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        options_in_out_mode = ['双列模式', '+/-单列模式', '标识列单列模式']
        frame_2.combobox_table_1_in_out_mode = ttk.Combobox(frame_2, values=options_in_out_mode, height=10, width=14, state='readonly')
        frame_2.combobox_table_1_in_out_mode.set(options_in_out_mode[0])
        frame_2.combobox_table_1_in_out_mode.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        self.control_frame_list[n].frame_2 = frame_2

        frame_3 = tk.Frame(self.control_frame_list[n])
        frame_3.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_3, text='Table 1 In Col', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.combobox_table_1_in_col = ttk.Combobox(frame_3, values=options, height=10, width=14, state='readonly')
        frame_3.combobox_table_1_in_col.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_3, text='Table 1 Out Col', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_3.combobox_table_1_out_col = ttk.Combobox(frame_3, values=options, height=10, width=14, state='readonly')
        frame_3.combobox_table_1_out_col.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_3, text='Reserved', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        tk.Entry(frame_3, width=17).pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_3, text='Reserved', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        tk.Entry(frame_3, width=17).pack(side=tk.LEFT, padx=(5, 5), pady=5)

        self.control_frame_list[n].frame_3 = frame_3

        frame_4 = tk.Frame(self.control_frame_list[n])
        frame_4.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_4, text='Table 1 In/Out Col', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_4.combobox_table_1_in_out_col = ttk.Combobox(frame_4, values=options, height=10, width=14, state='readonly')
        frame_4.combobox_table_1_in_out_col.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_4, text='Table 1 In Label', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_4.combobox_table_1_in_label = ttk.Combobox(frame_4, values=options, height=10, width=14, state='readonly')
        frame_4.combobox_table_1_in_label.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_4, text='Table 1 Out Label', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_4.combobox_table_1_out_label = ttk.Combobox(frame_4, values=options, height=10, width=14, state='readonly')
        frame_4.combobox_table_1_out_label.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_4, text='Table 1 In/Out Value', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_4.combobox_table_1_in_out_value = ttk.Combobox(frame_4, values=options, height=10, width=14, state='readonly')
        frame_4.combobox_table_1_in_out_value.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        self.control_frame_list[n].frame_4 = frame_4

        frame_5 = tk.Frame(self.control_frame_list[n])
        frame_5.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_5, text='Table 2 Name', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.combobox_table_2_name = ttk.Combobox(frame_5, values=options, height=10, width=14, state='readonly')
        frame_5.combobox_table_2_name.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_5, text='Table 2 Time Series', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.combobox_table_2_time_series = ttk.Combobox(frame_5, values=options_time_series, height=10, width=14, state='readonly')
        frame_5.combobox_table_2_time_series.set(options_time_series[0])
        frame_5.combobox_table_2_time_series.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_5, text='Table 2 Item', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.combobox_table_2_item = ttk.Combobox(frame_5, values=options, height=10, width=14, state='readonly')
        frame_5.combobox_table_2_item.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_5, text='Table 2 In/Out Mode', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_5.combobox_table_2_in_out_mode = ttk.Combobox(frame_5, values=options_in_out_mode, height=10, width=14, state='readonly')
        frame_5.combobox_table_2_in_out_mode.set(options_in_out_mode[0])
        frame_5.combobox_table_2_in_out_mode.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        self.control_frame_list[n].frame_5 = frame_5

        frame_6 = tk.Frame(self.control_frame_list[n])
        frame_6.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_6, text='Table 2 In Col', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_6.combobox_table_2_in_col = ttk.Combobox(frame_6, values=options, height=10, width=14, state='readonly')
        frame_6.combobox_table_2_in_col.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_6, text='Table 2 Out Col', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_6.combobox_table_2_out_col = ttk.Combobox(frame_6, values=options, height=10, width=14, state='readonly')
        frame_6.combobox_table_2_out_col.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_6, text='Reserved', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        tk.Entry(frame_6, width=17).pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_6, text='Reserved', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        tk.Entry(frame_6, width=17).pack(side=tk.LEFT, padx=(5, 5), pady=5)

        self.control_frame_list[n].frame_6 = frame_6

        frame_7 = tk.Frame(self.control_frame_list[n])
        frame_7.pack(side=tk.TOP, padx=5, pady=5)

        tk.Label(frame_7, text='Table 2 In/Out Col', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_7.combobox_table_2_in_out_col = ttk.Combobox(frame_7, values=options, height=10, width=14, state='readonly')
        frame_7.combobox_table_2_in_out_col.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_7, text='Table 2 In Label', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_7.combobox_table_2_in_label = ttk.Combobox(frame_7, values=options, height=10, width=14, state='readonly')
        frame_7.combobox_table_2_in_label.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_7, text='Table 2 Out Label', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_7.combobox_table_2_out_label = ttk.Combobox(frame_7, values=options, height=10, width=14, state='readonly')
        frame_7.combobox_table_2_out_label.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        tk.Label(frame_7, text='Table 2 In/Out Value', width=18, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        frame_7.combobox_table_2_in_out_value = ttk.Combobox(frame_7, values=options, height=10, width=14, state='readonly')
        frame_7.combobox_table_2_in_out_value.pack(side=tk.LEFT, padx=(5, 5), pady=5)

        self.control_frame_list[n].frame_7 = frame_7

    def button_read_fun(self, n):
        fill_text = ''

        n_config = self.control_frame_config[3][n]
        sql_path = self.control_frame_list[n_config].entry_widget.get()

        if sql_path:
            sql_table_import = sql_qlite_tool.sql_sqlite_tool_class()
            result_info = sql_table_import.get_all_tables(sql_path)
            if result_info[0]:
                option_table_name_list = result_info[1]
                frame_2 = self.control_frame_list[n].frame_2
                frame_2.combobox_table_1_name['values'] = option_table_name_list
                if option_table_name_list:
                    frame_2.combobox_table_1_name.set(option_table_name_list[0])
                frame_2.combobox_table_1_name.config(state='readonly')
                frame_5 = self.control_frame_list[n].frame_5
                frame_5.combobox_table_2_name['values'] = option_table_name_list
                if option_table_name_list:
                    frame_5.combobox_table_2_name.set(option_table_name_list[0])
                frame_5.combobox_table_2_name.config(state='readonly')
                fill_text += f'SQL table name list import successful!\n'
            else:
                fill_text += 'SQL table name list import failed!\n'
                fill_text += result_info[1]
        else:
            fill_text += f'Please check that the database path is not empty!\n'

        self.text_area_fill(fill_text)

    def button_table_1_fun(self, n):
        fill_text = ''

        n_config = self.control_frame_config[3][n]
        sql_path = self.control_frame_list[n_config].entry_widget.get()
        table_name = self.control_frame_list[n].frame_2.combobox_table_1_name.get()

        if sql_path and table_name:
            sql_table_import = sql_qlite_tool.sql_sqlite_tool_class()
            result_info = sql_table_import.get_table_columns(sql_path, table_name)
            if result_info[0]:
                option_col_name_list = result_info[1][1:]
                frame_2 = self.control_frame_list[n].frame_2

                frame_2.combobox_table_1_item['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_2.combobox_table_1_item.set(option_col_name_list[0])
                frame_2.combobox_table_1_item.config(state='readonly')

                frame_3 = self.control_frame_list[n].frame_3

                frame_3.combobox_table_1_in_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_3.combobox_table_1_in_col.set(option_col_name_list[0])
                frame_3.combobox_table_1_in_col.config(state='readonly')

                frame_3.combobox_table_1_out_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_3.combobox_table_1_out_col.set(option_col_name_list[0])
                frame_3.combobox_table_1_out_col.config(state='readonly')

                frame_4 = self.control_frame_list[n].frame_4

                frame_4.combobox_table_1_in_out_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_4.combobox_table_1_in_out_col.set(option_col_name_list[0])
                frame_4.combobox_table_1_in_out_col.config(state='readonly')

                frame_4.combobox_table_1_in_out_value['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_4.combobox_table_1_in_out_value.set(option_col_name_list[0])
                frame_4.combobox_table_1_in_out_value.config(state='readonly')

                fill_text += f'SQL table 1 column name list import successful!\n'
            else:
                fill_text += 'SQL table 1 column name list import failed!\n'
                fill_text += result_info[1]
        else:
            if not sql_path:
                fill_text += f'Please check that the database path is not empty!\n'
            if not table_name:
                fill_text += f'Please check that the table name is not empty!\n'

        self.text_area_fill(fill_text)

    def button_label_1_fun(self, n):
        fill_text = ''

        n_config = self.control_frame_config[3][n]
        sql_path = self.control_frame_list[n_config].entry_widget.get()
        table_1_name = self.control_frame_list[n].frame_2.combobox_table_1_name.get()
        table_1_in_out_col = self.control_frame_list[n].frame_4.combobox_table_1_in_out_col.get()
        table_1_in_out_value = self.control_frame_list[n].frame_4.combobox_table_1_in_out_value.get()

        unique_index_list_cleaned = []

        sql_sqlite_tool_class_example = sql_qlite_tool.sql_sqlite_tool_class()

        sql_template_single = """
                                SELECT *
                                FROM {table_name}
                                WHERE {table_in_out_value} IS NOT NULL
                              """

        sql_table_1_select = sql_template_single.format(table_name=table_1_name, table_in_out_value=table_1_in_out_value)

        result_info_table_1 = sql_sqlite_tool_class_example.execute_query(sql_path, sql_table_1_select)

        if result_info_table_1[0]:
            sql_table_1_df = result_info_table_1[1]
            fill_text += f'{table_1_name} read in/out label successful!\n'
        else:
            fill_text += result_info_table_1[1]

        unique_index_np = sql_table_1_df[table_1_in_out_col].unique()
        unique_index_list = unique_index_np.tolist()

        unique_index_list = [item for item in unique_index_list if item is not None]

        for item in unique_index_list:
            if not isinstance(item, (int, float)):                              # 检查item是否是int或float类型
                unique_index_list_cleaned.append(item.replace('\t', ''))        # 移除字符串中的所有制表符（\t）
            else:
                unique_index_list_cleaned.append(item)

        frame_4 = self.control_frame_list[n].frame_4

        frame_4.combobox_table_1_in_label['values'] = unique_index_list_cleaned
        if unique_index_list_cleaned[0]:
            frame_4.combobox_table_1_in_label.set(unique_index_list_cleaned[0])
        frame_4.combobox_table_1_in_label.config(state='readonly')
        
        frame_4.combobox_table_1_out_label['values'] = unique_index_list_cleaned
        if unique_index_list_cleaned[0]:
            frame_4.combobox_table_1_out_label.set(unique_index_list_cleaned[0])
        frame_4.combobox_table_1_out_label.config(state='readonly')

    def button_table_2_fun(self, n):
        fill_text = ''

        n_config = self.control_frame_config[3][n]
        sql_path = self.control_frame_list[n_config].entry_widget.get()
        table_name = self.control_frame_list[n].frame_5.combobox_table_2_name.get()

        if sql_path and table_name:
            sql_table_import = sql_qlite_tool.sql_sqlite_tool_class()
            result_info = sql_table_import.get_table_columns(sql_path, table_name)
            if result_info[0]:
                option_col_name_list = result_info[1][1:]
                frame_5 = self.control_frame_list[n].frame_5

                frame_5.combobox_table_2_item['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_5.combobox_table_2_item.set(option_col_name_list[0])
                frame_5.combobox_table_2_item.config(state='readonly')

                frame_6 = self.control_frame_list[n].frame_6

                frame_6.combobox_table_2_in_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_6.combobox_table_2_in_col.set(option_col_name_list[0])
                frame_6.combobox_table_2_in_col.config(state='readonly')

                frame_6.combobox_table_2_out_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_6.combobox_table_2_out_col.set(option_col_name_list[0])
                frame_6.combobox_table_2_out_col.config(state='readonly')

                frame_7 = self.control_frame_list[n].frame_7

                frame_7.combobox_table_2_in_out_col['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_7.combobox_table_2_in_out_col.set(option_col_name_list[0])
                frame_7.combobox_table_2_in_out_col.config(state='readonly')

                frame_7.combobox_table_2_in_out_value['values'] = option_col_name_list
                if option_col_name_list[0]:
                    frame_7.combobox_table_2_in_out_value.set(option_col_name_list[0])
                frame_7.combobox_table_2_in_out_value.config(state='readonly')

                fill_text += f'SQL table 2 column name list import successful!\n'
            else:
                fill_text += 'SQL table 2 column name list import failed!\n'
                fill_text += result_info[1]
        else:
            if not sql_path:
                fill_text += f'Please check that the database path is not empty!\n'
            if not table_name:
                fill_text += f'Please check that the table name is not empty!\n'

        self.text_area_fill(fill_text)

    def button_label_2_fun(self, n):
        fill_text = ''

        n_config = self.control_frame_config[3][n]
        sql_path = self.control_frame_list[n_config].entry_widget.get()
        table_2_name = self.control_frame_list[n].frame_5.combobox_table_2_name.get()
        table_2_in_out_col = self.control_frame_list[n].frame_7.combobox_table_2_in_out_col.get()
        table_2_in_out_value = self.control_frame_list[n].frame_7.combobox_table_2_in_out_value.get()

        unique_index_list_cleaned = []

        sql_sqlite_tool_class_example = sql_qlite_tool.sql_sqlite_tool_class()

        sql_template_single = """
                                SELECT *
                                FROM {table_name}
                                WHERE {table_in_out_value} IS NOT NULL
                              """

        sql_table_2_select = sql_template_single.format(table_name=table_2_name, table_in_out_value=table_2_in_out_value)

        result_info_table_2 = sql_sqlite_tool_class_example.execute_query(sql_path, sql_table_2_select)

        if result_info_table_2[0]:
            sql_table_2_df = result_info_table_2[1]
            fill_text += f'{table_2_name} read in/out label successful!\n'
        else:
            fill_text += result_info_table_2[1]

        unique_index_np = sql_table_2_df[table_2_in_out_col].unique()
        unique_index_list = unique_index_np.tolist()

        unique_index_list = [item for item in unique_index_list if item is not None]

        for item in unique_index_list:
            if not isinstance(item, (int, float)):                              # 检查item是否是int或float类型
                unique_index_list_cleaned.append(item.replace('\t', ''))        # 移除字符串中的所有制表符（\t）
            else:
                unique_index_list_cleaned.append(item)

        frame_7 = self.control_frame_list[n].frame_7

        frame_7.combobox_table_2_in_label['values'] = unique_index_list_cleaned
        if unique_index_list_cleaned[0]:
            frame_7.combobox_table_2_in_label.set(unique_index_list_cleaned[0])
        frame_7.combobox_table_2_in_label.config(state='readonly')
        
        frame_7.combobox_table_2_out_label['values'] = unique_index_list_cleaned
        if unique_index_list_cleaned[0]:
            frame_7.combobox_table_2_out_label.set(unique_index_list_cleaned[0])
        frame_7.combobox_table_2_out_label.config(state='readonly')

    def button_comparison_fun(self, n):
        fill_text = ''
        amounts_table_1_in_counts_set = set()
        amounts_table_2_in_counts_set = set()
        amounts_table_1_out_counts_set = set()
        amounts_table_2_out_counts_set = set()

        n_config = self.control_frame_config[3][n]
        sql_path = self.control_frame_list[n_config].entry_widget.get()
        table_1_name = self.control_frame_list[n].frame_2.combobox_table_1_name.get()
        table_1_time_series = self.control_frame_list[n].frame_2.combobox_table_1_time_series.get()
        table_1_item = self.control_frame_list[n].frame_2.combobox_table_1_item.get()
        table_1_in_out_mode = self.control_frame_list[n].frame_2.combobox_table_1_in_out_mode.get()
        table_1_in_col = self.control_frame_list[n].frame_3.combobox_table_1_in_col.get()
        table_1_out_col = self.control_frame_list[n].frame_3.combobox_table_1_out_col.get()
        table_1_in_out_col = self.control_frame_list[n].frame_4.combobox_table_1_in_out_col.get()
        table_1_in_label = self.control_frame_list[n].frame_4.combobox_table_1_in_label.get()
        table_1_out_label = self.control_frame_list[n].frame_4.combobox_table_1_out_label.get()
        table_1_in_out_value = self.control_frame_list[n].frame_4.combobox_table_1_in_out_value.get()

        table_2_name = self.control_frame_list[n].frame_5.combobox_table_2_name.get()
        table_2_time_series = self.control_frame_list[n].frame_5.combobox_table_2_time_series.get()
        table_2_item = self.control_frame_list[n].frame_5.combobox_table_2_item.get()
        table_2_in_out_mode = self.control_frame_list[n].frame_5.combobox_table_2_in_out_mode.get()
        table_2_in_col = self.control_frame_list[n].frame_6.combobox_table_2_in_col.get()
        table_2_out_col = self.control_frame_list[n].frame_6.combobox_table_2_out_col.get()
        table_2_in_out_col = self.control_frame_list[n].frame_7.combobox_table_2_in_out_col.get()
        table_2_in_label = self.control_frame_list[n].frame_7.combobox_table_2_in_label.get()
        table_2_out_label = self.control_frame_list[n].frame_7.combobox_table_2_out_label.get()
        table_2_in_out_value = self.control_frame_list[n].frame_7.combobox_table_2_in_out_value.get()

        file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])

        openxl_xlsx_tool_class_example = openxl_xlsx_tool.openxl_xlsx_tool_class()
        fill_text += openxl_xlsx_tool_class_example.in_out_result_new_xlsx(file_path)

        sql_sqlite_tool_class_example = sql_qlite_tool.sql_sqlite_tool_class()

        # sql_template_double = """
        #                         SELECT *
        #                         FROM {table_name}
        #                         WHERE ({table_in_col} IS NOT NULL AND {table_out_col} IS NULL) 
        #                         OR ({table_in_col} IS NULL AND {table_out_col} IS NOT NULL)
        #                       """

        # sql_template_single = """
        #                         SELECT *
        #                         FROM {table_name}
        #                         WHERE {table_in_out_value} IS NOT NULL
        #                       """

        sql_template = """
                        SELECT *
                        FROM {table_name}
                      """

        # table_1
        if table_1_in_out_mode == '双列模式':

            sql_table_1_select = sql_template.format(table_name=table_1_name, table_in_col=table_1_in_col, table_out_col=table_1_out_col)

        else:

            sql_table_1_select = sql_template.format(table_name=table_1_name, table_in_out_value=table_1_in_out_value)

        result_info_table_1 = sql_sqlite_tool_class_example.execute_query(sql_path, sql_table_1_select)

        if result_info_table_1[0]:
            sql_table_1_df = result_info_table_1[1]
            if table_1_time_series == '反向':
                sql_table_1_df[:] = sql_table_1_df.iloc[::-1].values
            fill_text += f'{table_1_name} read successful!\n'
        else:
            fill_text += result_info_table_1[1]

        # table_2
        if table_2_in_out_mode == '双列模式':

            sql_table_2_select = sql_template.format(table_name=table_2_name, table_in_col=table_2_in_col, table_out_col=table_2_out_col)

        else:

            sql_table_2_select = sql_template.format(table_name=table_2_name, table_in_out_value=table_2_in_out_value)

        result_info_table_2 = sql_sqlite_tool_class_example.execute_query(sql_path, sql_table_2_select)

        if result_info_table_2[0]:
            sql_table_2_df = result_info_table_2[1]
            if table_2_time_series == '反向':
                sql_table_2_df[:] = sql_table_2_df.iloc[::-1].values
            fill_text += f'{table_2_name} read successful!\n'
        else:
            fill_text += result_info_table_2[1]


        # 对比模式一：对比交易记录的数量与金额
        quantity_table_1_df = sql_table_1_df.shape[0]
        quantity_table_2_df = sql_table_2_df.shape[0]

        if quantity_table_1_df == quantity_table_2_df:
            fill_text += f'{table_1_name} quantity {quantity_table_1_df}, {table_2_name} quantity {quantity_table_2_df}, same!\n'
        else:
            fill_text += f'{table_1_name} quantity {quantity_table_1_df}, {table_2_name} quantity {quantity_table_2_df}, different!\n\n'

        if table_1_in_out_mode == '双列模式':
            amount_in_table_1_df = round(sql_table_1_df[table_1_in_col].sum(), 2)
            amount_out_table_1_df = round(sql_table_1_df[table_1_out_col].sum(), 2)

        elif table_1_in_out_mode == '+/-单列模式':
            amount_in_table_1_df = round(sql_table_1_df[sql_table_1_df[table_1_in_out_value] > 0][table_1_in_out_value].sum(), 2)
            amount_out_table_1_df = round(-sql_table_1_df[sql_table_1_df[table_1_in_out_value] < 0][table_1_in_out_value].sum(), 2)

        elif table_1_in_out_mode == '标识列单列模式':
            amount_in_table_1_df = round(sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_in_label][table_1_in_out_value].sum(), 2)
            amount_out_table_1_df = round(sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_out_label][table_1_in_out_value].sum(), 2)

        if table_2_in_out_mode == '双列模式':
            amount_in_table_2_df = round(sql_table_2_df[table_2_in_col].sum(), 2)
            amount_out_table_2_df = round(sql_table_2_df[table_2_out_col].sum(), 2)
        
        elif table_2_in_out_mode == '+/-单列模式':
            amount_in_table_2_df = round(sql_table_2_df[sql_table_2_df[table_2_in_out_value] > 0][table_2_in_out_value].sum(), 2)
            amount_out_table_2_df = round(-sql_table_2_df[sql_table_2_df[table_2_in_out_value] < 0][table_2_in_out_value].sum(), 2)

        elif table_2_in_out_mode == '标识列单列模式':
            amount_in_table_2_df = round(sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_in_label][table_2_in_out_value].sum(), 2)
            amount_out_table_2_df = round(sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_out_label][table_2_in_out_value].sum(), 2)


        if amount_in_table_1_df == amount_in_table_2_df:
            fill_text += f'{table_1_name} amount in {amount_in_table_1_df:,.2f}, {table_2_name} amount in {amount_in_table_2_df:,.2f}, same!\n'
        else:
            fill_text += f'{table_1_name} amount in {amount_in_table_1_df:,.2f}, {table_2_name} amount in {amount_in_table_2_df:,.2f}, different!\n'

        if amount_out_table_1_df == amount_out_table_2_df:
            fill_text += f'{table_1_name} amount out {amount_out_table_1_df:,.2f}, {table_2_name} amount out {amount_out_table_2_df:,.2f}, same!\n'
        else:
            fill_text += f'{table_1_name} amount out {amount_out_table_1_df:,.2f}, {table_2_name} amount out {amount_out_table_2_df:,.2f}, different!\n'

        # 填写.xlsx文件
        fill_text += openxl_xlsx_tool_class_example.in_out_result_fill_total_xlsx(file_path, quantity_table_1_df, quantity_table_2_df, amount_in_table_1_df, amount_out_table_1_df, amount_in_table_2_df, amount_out_table_2_df)

        # 对比模式二：数值分层对比
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
            amounts_table_1_out_counts_non['value'] = amounts_table_1_out_counts['value']           # 不改变负号
            amounts_table_1_out_counts_set_non = set(amounts_table_1_out_counts_non['value'])
            table_1_out_df = sql_table_1_df[sql_table_1_df[table_1_out_col].isin(amounts_table_1_out_counts_set_non)].dropna(subset=[table_1_in_out_value])
            table_1_out_df_col = table_1_in_out_value

        elif table_1_in_out_mode == '标识列单列模式':
            amounts_table_1_in_counts = sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_in_label][table_1_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_1_in_counts.columns = ['value', 'count_table_1']
            amounts_table_1_in_counts_set = set(amounts_table_1_in_counts['value'])
            table_1_in_df = sql_table_1_df[sql_table_1_df[table_1_in_out_value].isin(amounts_table_1_in_counts_set)].dropna(subset=[table_1_in_out_value])
            table_1_in_df_col = table_1_in_out_value

            amounts_table_1_out_counts = sql_table_1_df[sql_table_1_df[table_1_in_out_col] == table_1_out_label][table_1_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_1_out_counts.columns = ['value', 'count_table_1']
            amounts_table_1_out_counts_set = set(amounts_table_1_out_counts['value'])
            table_1_out_df = sql_table_1_df[sql_table_1_df[table_1_in_out_value].isin(amounts_table_1_out_counts_set)].dropna(subset=[table_1_in_out_value])
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
            amounts_table_2_out_counts['value'] = -amounts_table_2_out_counts['value']              # 负负得正
            amounts_table_2_out_counts_set = set(amounts_table_2_out_counts['value'])
            amounts_table_2_out_counts_non = pd.DataFrame()
            amounts_table_2_out_counts_non['value'] = amounts_table_2_out_counts['value']              # 不改变负号
            amounts_table_2_out_counts_set_non = set(amounts_table_2_out_counts_non['value'])
            table_2_out_df = sql_table_2_df[sql_table_2_df[table_2_out_col].isin(amounts_table_2_out_counts_set_non)].dropna(subset=[table_2_in_out_value])
            table_2_out_df_col = table_2_in_out_value

        elif table_2_in_out_mode == '标识列单列模式':
            amounts_table_2_in_counts = sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_in_label][table_2_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_2_in_counts.columns = ['value', 'count_table_2']
            amounts_table_2_in_counts_set = set(amounts_table_2_in_counts['value'])
            table_2_in_df = sql_table_2_df[sql_table_2_df[table_2_in_out_value].isin(amounts_table_2_in_counts_set)].dropna(subset=[table_2_in_out_value])
            table_2_in_df_col = table_2_in_out_value

            amounts_table_2_out_counts = sql_table_2_df[sql_table_2_df[table_2_in_out_col] == table_2_out_label][table_2_in_out_value].dropna().value_counts().sort_index().reset_index()
            amounts_table_2_out_counts.columns = ['value', 'count_table_2']
            amounts_table_2_out_counts_set = set(amounts_table_2_out_counts['value'])
            table_2_out_df = sql_table_2_df[sql_table_2_df[table_2_in_out_value].isin(amounts_table_2_out_counts_set)].dropna(subset=[table_2_in_out_value])
            table_2_out_df_col = table_2_in_out_value

        common_amounts_in = amounts_table_1_in_counts_set & amounts_table_2_in_counts_set
        common_amounts_out = amounts_table_1_out_counts_set & amounts_table_2_out_counts_set

        # in 数值分层对比
        result_in_text = self.comparison_layering_fun(amounts_table_1_in_counts, amounts_table_1_in_counts_set, amounts_table_2_in_counts, amounts_table_2_in_counts_set, common_amounts_in,
                                                       file_path, table_1_in_out_mode, table_1_in_df, table_1_in_df_col, table_2_in_out_mode, table_2_in_df, table_2_in_df_col, 'in')
        fill_text += '\n in 数值分层对比：\n'
        fill_text += result_in_text
        fill_text += '\n'

        # out 数值分层对比
        result_out_text = self.comparison_layering_fun(amounts_table_1_out_counts, amounts_table_1_out_counts_set, amounts_table_2_out_counts, amounts_table_2_out_counts_set, common_amounts_out,
                                                        file_path, table_1_in_out_mode, table_1_out_df, table_1_out_df_col, table_2_in_out_mode, table_2_out_df, table_2_out_df_col, 'out')
        fill_text += ' out 数值分层对比：\n'
        fill_text += result_out_text
        fill_text += '\n'

        self.text_area_fill(fill_text)

    def comparison_layering_fun(self, amounts_table_1_counts, amounts_table_1_counts_set, amounts_table_2_counts, amounts_table_2_counts_set, common_amounts,
                                file_path, table_1_in_out_mode, table_1_df, table_1_df_col, table_2_in_out_mode, table_2_df, table_2_df_col, in_our_out):
        result_text = ''

        # 共同金额的计数差异
        common_table_1 = amounts_table_1_counts[amounts_table_1_counts['value'].isin(common_amounts)].set_index('value')
        common_table_2 = amounts_table_2_counts[amounts_table_2_counts['value'].isin(common_amounts)].set_index('value')

        # 合并对比
        comparison = common_table_1.join(common_table_2, how='inner', lsuffix='_table_1', rsuffix='_table_2')
        comparison['计数差异'] = comparison['count_table_1'] - comparison['count_table_2']
        comparison['绝对差异'] = abs(comparison['计数差异'])

        # 筛选所有计数差异不为0的项（共同金额但计数不同）
        non_zero_diff = comparison[comparison['计数差异'] != 0][['count_table_1', 'count_table_2', '计数差异']]

        # 找出只在流水或明细中出现的金额
        only_in_table_1 = amounts_table_1_counts_set - amounts_table_2_counts_set
        only_in_table_2 = amounts_table_2_counts_set - amounts_table_1_counts_set

        result_text += '数值分层对比分析：\n'

        # 第一部分：共同金额但计数不同的情况
        if len(non_zero_diff) > 0:
            result_text += f'1. 共同数值但计数不同（{len(non_zero_diff)}个）：\n'
            
            for amount in non_zero_diff.index:
                table_1_count = comparison.loc[amount, 'count_table_1']
                table_2_count = comparison.loc[amount, 'count_table_2']
                diff = table_1_count - table_2_count
                
                if diff > 0:
                    diff_desc = f'table 1 比 table 2 多 {diff} 次'
                else:
                    diff_desc = f'table 2 比 table 1 多 {abs(diff)} 次'
                    
                result_text += f'   数值 {amount:>10}: table 1 {table_1_count:>2} 次, table 2 {table_2_count:>2} 次, {diff_desc}\n'

        else:
            result_text += '1. 共同数值计数完全一致，无差异。\n'

        # 第二部分：只在 table 1 中出现的金额
        if len(only_in_table_1) > 0:
            result_text += f'2. 只在 table 1 中出现的数值（{len(only_in_table_1)}个）：\n'
            
            only_table_1_df = amounts_table_1_counts[amounts_table_1_counts['value'].isin(only_in_table_1)].set_index('value')
            for amount in only_table_1_df.index:
                table_1_count = only_table_1_df.loc[amount, 'count_table_1']
                result_text += f'   数值 {amount:>10}: 出现 {table_1_count:>2} 次（table 2 中不存在）\n'
        else:
            result_text += '2. 没有只在 table 1 中出现的金额。\n'

        # 第三部分：只在 table 2 中出现的金额
        if len(only_in_table_2) > 0:
            result_text += f'3. 只在 table 2 中出现的数值（{len(only_in_table_2)}个）：\n'
            
            only_table_2_df = amounts_table_2_counts[amounts_table_2_counts['value'].isin(only_in_table_2)].set_index('value')
            for amount in only_table_2_df.index:
                table_2_count = only_table_2_df.loc[amount, 'count_table_2']
                result_text += f"   数值 {amount:>10}: 出现 {table_2_count:>2} 次（table 1 中不存在）\n"
        else:
            result_text += '3. 没有只在 table 2 中出现的数值。\n'

        # 统计总结
        total_common = len(common_amounts)
        total_table_1_only = len(only_in_table_1)
        total_table_2_only = len(only_in_table_2)
        total_total = total_common + total_table_1_only + total_table_2_only

        result_text += f'统计总结：\n'
        result_text += f'  共同出现数值个数: {total_total} 个\n其中：\n'
        result_text += f'  共同出现的数值: {total_common} 个\n'
        result_text += f'  只在 table 1 出现的数值: {total_table_1_only} 个\n'
        result_text += f'  只在 table 2 出现的数值: {total_table_2_only} 个\n'

        # 填写.xlsx文件
        openxl_xlsx_tool_class_example = openxl_xlsx_tool.openxl_xlsx_tool_class()

        if in_our_out == 'in':

            non_zero_diff_reset = non_zero_diff.reset_index()

            # 数值分层对比_in_共同数值但计数不同_table_1
            non_zero_diff_reset = non_zero_diff_reset.rename(columns={'value': table_1_df_col})
            if table_1_in_out_mode == '双列模式':
                fill_df = pd.merge(table_1_df, non_zero_diff_reset[[table_1_df_col]], on=table_1_df_col, how='inner')
            elif table_1_in_out_mode == '+/-单列模式':
                fill_df = pd.merge(table_1_df, non_zero_diff_reset[[table_1_df_col]], on=table_1_df_col, how='inner')
            elif table_1_in_out_mode == '标识列单列模式':
                fill_df = pd.merge(table_1_df, non_zero_diff_reset[[table_1_df_col]], on=table_1_df_col, how='inner')

            result_text += openxl_xlsx_tool_class_example.in_out_result_fill_layering_xlsx(file_path, '数值分层对比_in_共同数值但计数不同_table_1', fill_df)

            # 数值分层对比_in_共同数值但计数不同_table_2
            non_zero_diff_reset = non_zero_diff_reset.rename(columns={table_1_df_col: table_2_df_col})
            if table_2_in_out_mode == '双列模式':
                fill_df = pd.merge(table_2_df, non_zero_diff_reset[[table_2_df_col]], on=table_2_df_col, how='inner')
            elif table_2_in_out_mode == '+/-单列模式':
                fill_df = pd.merge(table_2_df, non_zero_diff_reset[[table_2_df_col]], on=table_2_df_col, how='inner')
            elif table_2_in_out_mode == '标识列单列模式':
                fill_df = pd.merge(table_2_df, non_zero_diff_reset[[table_2_df_col]], on=table_2_df_col, how='inner')

            result_text += openxl_xlsx_tool_class_example.in_out_result_fill_layering_xlsx(file_path, '数值分层对比_in_共同数值但计数不同_table_2', fill_df)

            # 数值分层对比_in_只在 table 1 中出现的数值
            if table_1_in_out_mode == '双列模式':
                fill_df = table_1_df[table_1_df[table_1_df_col].isin(only_in_table_1)]
            elif table_1_in_out_mode == '+/-单列模式':
                fill_df = table_1_df[table_1_df[table_1_df_col].isin(only_in_table_1)]
            elif table_1_in_out_mode == '标识列单列模式':
                fill_df = table_1_df[table_1_df[table_1_df_col].isin(only_in_table_1)]

            result_text += openxl_xlsx_tool_class_example.in_out_result_fill_layering_xlsx(file_path, '数值分层对比_in_只在 table 1 中出现的数值', fill_df)

            # 数值分层对比_in_只在 table 2 中出现的数值
            if table_2_in_out_mode == '双列模式':
                fill_df = table_2_df[table_2_df[table_2_df_col].isin(only_in_table_2)]
            elif table_2_in_out_mode == '+/-单列模式':
                fill_df = table_2_df[table_2_df[table_2_df_col].isin(only_in_table_2)]
            elif table_2_in_out_mode == '标识列单列模式':
                fill_df = table_2_df[table_2_df[table_2_df_col].isin(only_in_table_2)]

            result_text += openxl_xlsx_tool_class_example.in_out_result_fill_layering_xlsx(file_path, '数值分层对比_in_只在 table 2 中出现的数值', fill_df)

        elif in_our_out == 'out':

            non_zero_diff_reset = non_zero_diff.reset_index()

            # 数值分层对比_in_共同数值但计数不同_table_1
            non_zero_diff_reset = non_zero_diff_reset.rename(columns={'value': table_1_df_col})
            if table_1_in_out_mode == '双列模式':
                fill_df = pd.merge(table_1_df, non_zero_diff_reset[[table_1_df_col]], on=table_1_df_col, how='inner')
            elif table_1_in_out_mode == '+/-单列模式':
                fill_df = pd.merge(table_1_df, non_zero_diff_reset[[table_1_df_col]], on=table_1_df_col, how='inner')
            elif table_1_in_out_mode == '标识列单列模式':
                fill_df = pd.merge(table_1_df, non_zero_diff_reset[[table_1_df_col]], on=table_1_df_col, how='inner')

            result_text += openxl_xlsx_tool_class_example.in_out_result_fill_layering_xlsx(file_path, '数值分层对比_out_共同数值但计数不同_table_1', fill_df)

            # 数值分层对比_in_共同数值但计数不同_table_2
            non_zero_diff_reset = non_zero_diff_reset.rename(columns={table_1_df_col: table_2_df_col})
            if table_2_in_out_mode == '双列模式':
                fill_df = pd.merge(table_2_df, non_zero_diff_reset[[table_2_df_col]], on=table_2_df_col, how='inner')
            elif table_2_in_out_mode == '+/-单列模式':
                fill_df = pd.merge(table_2_df, non_zero_diff_reset[[table_2_df_col]], on=table_2_df_col, how='inner')
            elif table_2_in_out_mode == '标识列单列模式':
                fill_df = pd.merge(table_2_df, non_zero_diff_reset[[table_2_df_col]], on=table_2_df_col, how='inner')

            result_text += openxl_xlsx_tool_class_example.in_out_result_fill_layering_xlsx(file_path, '数值分层对比_out_共同数值但计数不同_table_2', fill_df)

            # 数值分层对比_in_只在 table 1 中出现的数值
            if table_1_in_out_mode == '双列模式':
                fill_df = table_1_df[table_1_df[table_1_df_col].isin(only_in_table_1)]
            elif table_1_in_out_mode == '+/-单列模式':
                fill_df = table_1_df[table_1_df[table_1_df_col].isin(only_in_table_1)]
            elif table_1_in_out_mode == '标识列单列模式':
                fill_df = table_1_df[table_1_df[table_1_df_col].isin(only_in_table_1)]

            result_text += openxl_xlsx_tool_class_example.in_out_result_fill_layering_xlsx(file_path, '数值分层对比_out_只在 table 1 中出现的数值', fill_df)

            # 数值分层对比_in_只在 table 2 中出现的数值
            if table_2_in_out_mode == '双列模式':
                fill_df = table_2_df[table_2_df[table_2_df_col].isin(only_in_table_2)]
            elif table_2_in_out_mode == '+/-单列模式':
                fill_df = table_2_df[table_2_df[table_2_df_col].isin(only_in_table_2)]
            elif table_2_in_out_mode == '标识列单列模式':
                fill_df = table_2_df[table_2_df[table_2_df_col].isin(only_in_table_2)]

            result_text += openxl_xlsx_tool_class_example.in_out_result_fill_layering_xlsx(file_path, '数值分层对比_out_只在 table 2 中出现的数值', fill_df)

        return result_text



    def convert_none(self, value):
        if str(value).strip().lower() == 'none':
            return None
        else:
            try:
                return int(value)
            except:
                return value 
            
'''