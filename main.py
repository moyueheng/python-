import os # 拼接路径
import win32com.client as wc # 将doc转化位docx
from config import Config  # 配置文件
import my_doc # doc处理模块
import my_docx
import my_excel

'''获得doc类型的word文档路径'''
def get_doc_paths(doc_dir_path: str)->list:
    doc_names = os.listdir(doc_dir_path)
    doc_paths = []
    for doc_name in doc_names:
        path = os.path.join(doc_dir_path, doc_name)
        doc_paths.append(path)
    return doc_paths


'''同名的获得docx类型的word文档路径'''
def get_docx_paths(doc_dir_path: str, docx_dir_path: str)->list:
    doc_names = os.listdir(doc_dir_path)
    docx_paths = []
    for doc_name in doc_names:
        # 改名
        if doc_name.endswith('.doc'):
            path = os.path.join(
                docx_dir_path, doc_name.replace('.doc', '.docx'))
        else:
            path = os.path.join(docx_dir_path, doc_name)
        docx_paths.append(path)
    return docx_paths

'''doc转docx'''
def doc_to_docx(doc_paths,docx_paths):
    i = 0 
    while True:
        try:
            convertor = my_doc.Convertor()
            convertor.doc2docx(doc_paths[i], docx_paths[i])
            i+=1
        except Exception as ret:
            print("转化完成！！！%s"% ret)
            break

'''从word中获得数据并且放入excel中'''
def add_data_from_word(docx_paths):
    # 1. 从word文档中获得数据
    my_word = my_docx.MyDocx()
    student_data_from_word = my_word.get_data(docx_paths,Config.get('word_data_positions'))
    # 2.存放到excel表中
    my_xl_in = my_excel.MyExcel(Config.get('result_excel_path'))
    my_xl_in.add_data(student_data_from_word,Config.get('result_excel_sheet_name'),Config.get('result_excel_index_col'),Config.get('result_positons_from_word'))
    my_xl_in.save()
    
'''从excel中获得数据并且放入excel中'''
def add_data_from_excel():
    # 3.从excel中取出数据
    my_xl_out = my_excel.MyExcel(Config.get('excel_path'))
    student_data_from_excel = my_xl_out.get_data(Config.get('excel_sheet_name'),Config.get('excel_data_positions'))
    # 4.把从excel取出的数据放入需要存放数据的excel
    my_xl_in = my_excel.MyExcel(Config.get('result_excel_path'))
    my_xl_in.add_data(student_data_from_excel,Config.get('result_excel_sheet_name'),Config.get('result_excel_index_col'),Config.get('resulet_positon_from_excel'))
    my_xl_in.save()


def main():
    # 1.先把所有的doc和docx的绝对路径拼出来
    doc_paths = get_doc_paths(Config.get('doc_dir_path'))
    docx_paths = get_docx_paths(Config.get('doc_dir_path'), Config.get('docx_dir_path'))
    # 2. 把doc转为docx
    doc_to_docx(doc_paths,docx_paths)
    # 3. 读取有效数据放入excel中
    add_data_from_word(docx_paths)
    add_data_from_excel()


if __name__ == "__main__":
    main()
