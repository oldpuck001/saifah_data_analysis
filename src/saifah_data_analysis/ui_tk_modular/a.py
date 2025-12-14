
import pandas as pd

columns = ['name', 'No.', 'country', 'score', 'job']

index=[101, 100, 102, 103, 104, 105, 106, 107, 108, 109]

data = [['Mike', 1, 'Thailand', 80, 'teacher'],
        ['Yang', 2, 'China', 77, 'student'],
        ['Tom', 3, 'England', 85, 'student'],
        ['Losa', 4, 'Japan', 90, 'accounting'],
        ['Tim', 6, 'America', 87, 'student'],
        ['Zhang', 5, 'China', 73, 'student'],
        ['Jack', 9, 'India', 85, 'student'],
        ['Wang', 8, 'China', 89, 'student'],
        ['Chang', 10, 'Thailand', 94, 'accounting'],
        ['Lucy', 7, 'Japan', 91, 'employee']
        ]

df = pd.DataFrame(data, columns=columns, index=index)
print(df)

# groupby方法
print('groupby方法分組數據')

# groupby方法分組數據，使用一個列的值分組數據
print(df.groupby('job')['score'].sum())
print(df.groupby(['job']).sum())

'''
    # 选择文件/文件夹
    def select_file_floder_frame(self, n, control_config):

        tk.Label(self.control_frame_list[n], text=self.control_frame_config[1][n], width=20, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        entry_widget = tk.Entry(self.control_frame_list[n], state='readonly', readonlybackground='white')                                    # 创建Entry并保存引用
        entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        tk.Button(self.control_frame_list[n], text=self.control_frame_config[1][n], command=lambda e=entry_widget: self.select_open_file_path(e, control_config), width=20).pack(side=tk.LEFT, padx=5, pady=5)
        self.control_frame_list[n].entry_widget = entry_widget

    def select_open_file_path(self, entry_widget, control_config):
        fill_text = ''

        if control_config == 'askopenfilename':
            path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx'), ('Excel Files', '*.xls'), ('Text Files', '*.txt'), ('All Files', '*.*')])
        elif control_config == 'askdirectory':
            path = filedialog.askdirectory(title='选择文件夹', initialdir='/')
        else:
            path = ''

        if path:
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            entry_widget.config(state='readonly')
            if control_config == 'askopenfilename':
                fill_text += f'Selected: {path}\n'
                fill_text += 'File selection successful!\n'
            elif control_config == 'askdirectory':
                fill_text += f'Selected: {path}\n'
                fill_text += 'Floder selection successful!\n'
        else:
            if control_config == 'askopenfilename':
                fill_text += f'File selection failed!\n'
            elif control_config == 'askdirectory':
                fill_text += f'Floder selection failed!\n'

        self.text_area_fill(fill_text)

    # 文件保存路径
    def save_as_file_frame(self, n, control_config):
        tk.Label(self.control_frame_list[n], text=self.control_frame_config[1][n], width=20, anchor='w').pack(side=tk.LEFT, padx=5, pady=5)
        entry_widget = tk.Entry(self.control_frame_list[n], state='readonly', readonlybackground='white')                                    # 创建Entry并保存引用
        entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        tk.Button(self.control_frame_list[n], text=self.control_frame_config[1][n], command=lambda e=entry_widget: self.save_as_file_path(e, control_config), width=20).pack(side=tk.LEFT, padx=5, pady=5)

    def save_as_file_path(self, entry_widget, control_config):
        fill_text = ''

        if control_config == 'asksaveasfilename':
            path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files (2007+)', '*.xlsx'),('Excel Files (97-2003)', '*.xls'),('CSV Files', '*.csv'),('Text Files', '*.txt'),('All Files', '*.*')])
        else:
            path = ''

        if path:
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)
            entry_widget.config(state='readonly')
            fill_text = f'Selected: {path}\nFile saved successfully!\n'
        else:
            fill_text = f'File saving failed!\n'

        self.text_area_fill(fill_text)

    # 通用按钮
    def button_general_frame(self, n):
        tk.Button(self.control_frame_list[n], text=self.control_frame_config[1][n], command=lambda n=n: self.button_general_fun(n), width=20).pack(side=tk.LEFT, padx=5, pady=5)

    def button_general_fun(self, n):
        button_fun = self.control_frame_config[2][n]
        n_config = self.control_frame_config[3][n]
        n_path = self.control_frame_list[n_config].entry_widget.get()
        button_fun(original_file_path=n_path)
'''