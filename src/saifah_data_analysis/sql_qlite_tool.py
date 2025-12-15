# sql_sqlite_tool.py

import pandas as pd
import sqlite3
from sqlite3 import Error

class sql_sqlite_tool_class:
    
    def __init__(self):
        self.conn = None

    # 创建数据库
    def create_database_sqlite(self, database_path):
        try:
            self.conn = sqlite3.connect(database_path)                                                 # 建立数据库连接
            return [True, '']
        except Error as e:
            return [False, e]
        finally:
            self.close_connection()


    # 关闭数据库连接
    def close_connection(self):
        if self.conn:
            self.conn.close()
            self.conn = None


    # 导入DataFrame到指定表
    def import_dataframe(self, database_file_path, df, table_name, if_exists, index, index_label):

        try:
            self.conn = sqlite3.connect(database_file_path)                                                 # 建立数据库连接
            df.to_sql(table_name, self.conn, if_exists=if_exists, index=index, index_label=index_label)     # 导入至数据库
            return [True, '']
        except Error as e:
            return [False, e]
        finally:
            self.close_connection()


    # 获取所有表名
    def get_all_tables(self, database_file_path):
        try:
            self.conn = sqlite3.connect(database_file_path)
            cursor = self.conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            return [True, [table[0] for table in tables]]
        except Error as e:
            return [False, e]
        finally:
            self.close_connection()


    # 获取所有列名
    def get_table_columns(self, database_file_path, table_name):
        try:
            self.conn = sqlite3.connect(database_file_path)
            cursor = self.conn.cursor()
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns = cursor.fetchall()
            return [True, [column[1] for column in columns]]                # 提取列名（PRAGMA返回的第二个元素是列名）
        except Error as e:
            return [False, e]
        except Exception as e:
            return [False, e]
        finally:
            self.close_connection()


    # 执行SQL查询并返回结果DataFrame
    def execute_query(self, database_file_path, query):
        try:
            self.conn = sqlite3.connect(database_file_path)         # 建立数据库连接
            result_df = pd.read_sql(query, self.conn)               # 从数据库读取数据
            return [True, result_df]
        except Error as e:
            return [False, e]
        finally:
            self.close_connection()


    # 执行SQL指令
    def sql_sqlite_command(self, sql_path, sql_command):

        conn = sqlite3.connect(sql_path)
        curs = conn.cursor()
        try:
            sql_clean = sql_command.strip()
            sql_type = sql_clean.lower().split()[0]
            curs.execute(sql_clean)
            # 查询语句
            if sql_type == 'select':
                rows = curs.fetchall()
                columns = [desc[0] for desc in curs.description]
                if rows:
                    result_lines = []
                    result_lines.append(" | ".join(columns))
                    result_lines.append("-" * 50)
                    for row in rows:
                        result_lines.append(" | ".join(str(v) for v in row))
                    result_text = (f'\n指令：\n{sql_command}\n'
                                   f'查询结果（{len(rows)} 行）：\n' + '\n'.join(result_lines) + '\n')
                else:
                    result_text = (f'\n指令：\n{sql_command}\n''查询成功，但结果为空。\n')

            # 非查询语句
            else:
                conn.commit()
                result_text = (f'\n指令：\n{sql_command}\n执行成功！影响行数：{curs.rowcount}\n')

        except Exception as e:
            result_text = (f'\n指令：\n{sql_command}\n执行失败！\n错误信息：{str(e)}\n')

        finally:
            curs.close()
            conn.close()

        return result_text


    # 获取表信息
    def get_table_info(self, sql_path, table_name):
        try:
            # 建立数据库连接
            self.conn = sqlite3.connect(sql_path)
            # 表结构
            cursor = self.conn.cursor()
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns = cursor.fetchall()
            # 行数
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            row_count = cursor.fetchone()[0]
            return [True, [columns, row_count]]
        except Error as e:
            return [False, e]
        finally:
            self.close_connection()


    # 检查表是否存在
    def table_exists(self, sql_path, table_name):
        try:
            self.conn = sqlite3.connect(sql_path)
            cursor = self.conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
            result = cursor.fetchone()
            if result is None:
                return [False, f'False! The tabel {table_name} was not found in the database.\n']
            else:
                return [True, f'True! The tabel {table_name} was found in the database.\n']
        except Error as e:
            return [False, e]
        finally:
            self.close_connection()


    # 删除单个table
    def del_single_table(self, sql_path, table_name):
        try:
            self.conn = sqlite3.connect(sql_path)
            cursor = self.conn.cursor()
            cursor.execute(f'DROP TABLE IF EXISTS "{table_name}"')
            return [True, 'Table deletion successful!\n']
        except Error as e:
            return [False, e]
        finally:
            self.close_connection()


    # 删除所有table
    def del_all_tables(self, sql_path):
        try:
            self.conn = sqlite3.connect(sql_path)
            cursor = self.conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")        # 查询所有表名
            tables = cursor.fetchall()
            for (table_name,) in tables:
                cursor.execute(f"DROP TABLE IF EXISTS '{table_name}'")                                              # 逐个删除
            return [True, 'All tables deletion successful!\n']
        except Error as e:
            return [False, e]
        finally:
            self.close_connection()