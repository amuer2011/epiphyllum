import Excel2Sql

# excel原路径
excel_path = 'F:/test/Omega_reloan_add_credit_0807.xlsx'
# sql文件生成路径
target_file_path = 'F:/test/Omega_reloan_add_credit_0807.sql'
# sql插入语句目标表名
table_name = 'cl_user_activity_rel'
# 单次insert语句插入行数
page_size = 500
# 换行打印insert行数
column_size = 5

# 调用模块方法，将excel转换成批量插入sql文件
Excel2Sql.parse_excel_2_sql(excel_path, target_file_path, table_name, page_size, column_size)
