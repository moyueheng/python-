import docx
import config


class MyDocx():
    def __init__(self):
        pass
    

    def get_data(self,docx_paths:str,data_positions,table_index = 0)->list:
        """从word表格中获得字符串
        Args:
            docx_paths       word文档全部的docx格式的路径列表
            docx_positions  word文档中需要数据的坐标
            table_index     第几个表格
        return:
            按顺序返回一个信息列表
        """
        # 打开word
        for path in docx_paths:
            try:
                docSrt = docx.Document(path)
                # 选中表格
                tb = docSrt.tables[table_index]        
                student_data = []
                for x,y in data_positions:
                    student_data.append(tb.rows[x].cells[y].text)
                yield student_data
                docSrt.save(path)
            except:
                pass


def main():
    my_docx = MyDocx()
    print(my_docx.get_data('取出.docx',config.Config.get('word_data_positions')))


if __name__ == "__main__":
    main()