import openpyxl


def parse_excel_2_sql(excel_path: str, target_file_path: str, table_name: str, page_size: int = 15, column_size: int = 5):
    """ 把exel转换成批量insert语句
    :param excel_path: 本地excel路径
    :param target_file_path: 生成sql文件目录，如：F:/test/test.sql
    :param table_name: 生成sql语句目标表名
    :param page_size: 每个insert语句插入行数，默认15
    :param column_size: 每个insert语句几行换一次行，方便打印，默认5

    .. Note::

        1.excel默认打开第一个sheet，第一行作为表头和table_name生成对应insert语句\n
        2.生成sql文件（路径target_file_path）会以覆盖模式读写，每次都会擦除之前的内容\n
        3.本方法一次性加载excel文件内容到内存中，不适用大文件操作
    """

    if table_name is None or excel_path is None or target_file_path is None or page_size is None or column_size is None:
        print('输入参数不能为空！')
        return

    # 以覆盖写模式打开目标sql文件
    local_file = open(target_file_path, 'w+')

    # 打开excel文件,获取工作簿对象
    wb = openpyxl.load_workbook(excel_path)

    # 从表单中获取单元格的内容
    ws = wb.active  # 当前活跃的表单

    # excel转存数组
    rows = []

    # 遍历并转存excel行数组，方便后续处理
    for row in ws.iter_rows():
                rows.append(row)

    # 总行数
    total_length = len(rows)

    # 当前sql位置，换行标识
    current_index = 0

    # 遍历表格
    for row_index in range(total_length):
        # 获取行对象
        row = rows[row_index]

        # 排除第一行表头
        if row_index == 0:
            # 初始化sql语句
            base_sql = 'insert into ' + table_name + '('
            length = len(row)
            for index in range(length):
                cell = row[index]
                base_sql += str(cell.value)
                if index < length - 1:
                    base_sql += ','
            base_sql += ') values \n'
            # 当前sql语句
            tmp_sql = base_sql
            continue

        # 生成行sql
        tmp_sql += '('
        length = len(row)
        for index in range(length):
            cell = row[index]
            tmp_sql = tmp_sql + '\'' + str(cell.value) + '\''
            if index < length - 1:
                tmp_sql += ','
        tmp_sql += ')'

        if row_index % page_size == 0 and row_index < total_length - 1 or row_index == total_length - 1:
            # 当到最后一行或者达到一页的时候，生成一个sql语句，重置当前索引位置
            tmp_sql += ';\n'
            local_file.write(tmp_sql + '\n')
            tmp_sql = base_sql
            current_index = 0
        else:
            # 每行以逗号分开，批量插入
            tmp_sql += ','
            # 每5行一个换行符，看起来美观一些
            current_index += 1
            if current_index != 1 and current_index % column_size == 0:
                tmp_sql += '\n'

    print('生成sql文件成功！文件路径：' + target_file_path)
    local_file.close()
