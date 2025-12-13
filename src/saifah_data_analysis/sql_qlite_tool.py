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


    # # 导出.xlsx文件
    # def sqlite_to_xlsx(self, database_file_path, table_name, xlsx_file_path):
    #     try:
    #         result = self.execute_query(database_file_path, f"SELECT * FROM {table_name}")
    #         result.to_excel(xlsx_file_path)
    #         return True
    #     except:
    #         return False


    # # 检查表是否存在
    # def table_exists(self, database_file_path, table_name):
    #     try:
    #         self.conn = sqlite3.connect(database_file_path)
    #         cursor = self.conn.cursor()
    #         cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
    #         result = cursor.fetchone()
    #         if result is None:
    #             return False
    #         else:
    #             return True
    #     except Error as e:
    #         return False
    #     finally:
    #         self.close_connection()


    # # 获取表信息
    # def get_table_info(self, database_file_path, table_name):
    #     try:
    #         # 建立数据库连接
    #         self.conn = sqlite3.connect(database_file_path)
    #         # 表结构
    #         cursor = self.conn.cursor()
    #         cursor.execute(f"PRAGMA table_info({table_name})")
    #         columns = cursor.fetchall()
    #         # 行数
    #         cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
    #         row_count = cursor.fetchone()[0]
    #         return {
    #             'columns': columns,
    #             'row_count': row_count
    #         }
    #     except Error as e:
    #         return None
    #     finally:
    #         self.close_connection()