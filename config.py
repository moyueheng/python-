Config = {
    'doc_dir_path': 'F:\\学校事务\\python-\\doc文件',
    'docx_dir_path': 'F:\\学校事务\\python-\\docx文件',

    'excel_path': 'F:\\学校事务\\python-\\取出.xlsx',
    'excel_sheet_name': '成绩', # 表单名
    'excel_data_positions':  [1, 14, 16],  # 我们要从excel中的第14，16列取出成绩和排名,第一个元素必须是姓名
    
    # 姓名，身高，体重，高中任职，寝室号，手机号，qq，邮箱，家庭地址，家庭邮编，父亲姓名，父亲号码，母亲姓名，母亲号码，二本线
    'word_data_positions': [(0, 1), (4, 1), (4, 3), (9, 5), (4, 5), (2, 5), (3, 1), (3, 3), (6, 1), (6, 7), (7, 1), (7, 7), (8, 1), (8, 7), (5, 7)], # 第一个元素必须是姓名

    # 需要插入的位置的列号
    'result_excel_path': 'F:\\学校事务\\python-\\存入.xlsx', # 需要存放excel的的文件目录
    'result_excel_sheet_name': 'Sheet1', # 表单名
    'result_excel_index_col': 'B', # 姓名所在列
    'result_positons_from_word': [9, 10, 11, 15, 16, 17, 18, 20, 21, 22, 23, 24, 25, 28], # 从word获得的数据需要存放的位置的列号
    'resulet_positon_from_excel':[29, 30], # 从excel中获取数据需要存放的位置的列号
}

