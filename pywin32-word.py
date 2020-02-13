import os
import win32com.client
import xlrd


class RemoteWord:
    def __init__(self, filename=None):
        self.xlApp = win32com.client.DispatchEx('Word.Application')
        self.xlApp.Visible = 0
        self.xlApp.DisplayAlerts = 0  # 后台运行，不显示，不警告
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.doc = self.xlApp.Documents.Open(filename)
            else:
                self.doc = self.xlApp.Documents.Add()  # 创建新的文档
                self.doc.SaveAs(filename)
        else:
            self.doc = self.xlApp.Documents.Add()
            self.filename = ''

    def add_doc_end(self, string):
        '''在文档末尾添加内容'''
        rangee = self.doc.Range()
        rangee.InsertAfter('\n' + string)

    def add_doc_start(self, string):
        '''在文档开头添加内容'''
        rangee = self.doc.Range(0, 0)
        rangee.InsertBefore(string + '\n')

    def insert_doc(self, insertPos, string):
        '''在文档insertPos位置添加内容'''
        rangee = self.doc.Range(0, insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n' + string)

    def replace_doc(self, string, new_string):
        '''替换文字'''
        self.xlApp.Selection.Find.ClearFormatting()
        self.xlApp.Selection.Find.Replacement.ClearFormatting()
        self.xlApp.Selection.Find.Execute(string, False, False, False, False, False, True, 1, True, new_string, 2)

    def save(self):
        '''保存文档'''
        self.doc.Save()

    def save_as(self, filename):
        '''文档另存为'''
        self.doc.SaveAs(filename)

    def close(self):
        '''关闭文件'''
        self.save()
        self.xlApp.Documents.Close()
        self.xlApp.Quit()


# 遍历找到word文件路径

def find_docx(pdf_path):
    file_list = []
    if os.path.isfile(pdf_path):
        file_list.append(pdf_path)
    else:
        for top, dirs, files in os.walk(pdf_path):
            for filename in files:
                if filename.endswith('.docx') or filename.endswith('.doc'):
                    abspath = os.path.join(top, filename)
                    file_list.append(abspath)
    return file_list


def folder_exists(path):
    tar_path = os.path.join(path, '替换结果')
    is_exists = os.path.exists(tar_path)
    if not is_exists:
        print('替换结果目录不存在！创建新的替换结果目录')
        if os.makedirs(tar_path):
            print(path + ' 创建成功')
    else:
        print('替换结果目录已存在！替换后的文档将保存在目录下')


if __name__ == '__main__':
    path = os.getcwd()  # 文件夹路径
    folder_exists(path)  # 判断替换内容文件夹是否存在当前目录下
    excel_file = os.path.join(path, '替换内容.xlsx')
    if not os.path.isfile(excel_file):
        print('替换内容.xlsx 不存在，请在目录下创建')
        quit()
    else:
        print('正在读取替换内容.xlsx...')
    wb = xlrd.open_workbook(filename=excel_file)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    old_info = []
    new_info = []
    for i in range(sheet1.nrows):
        old_info.append(sheet1.cell(i, 0).value)
        new_info.append(sheet1.cell(i, 1).value)
        # print(old_info)

    files = []  # 目录下的文件
    for file in os.listdir(path):
        if file.endswith(".docx") or file.endswith(".doc"):  # 排除文件夹内的其它干扰文件，只获取word文件
            files.append(os.path.join(path, file))
    for file in files:
        doc = RemoteWord(file)  # 初始化一个doc对象
        for i in range(len(old_info)):
            doc.replace_doc(old_info[i], new_info[i])  # 替换文本内容
        doc.save_as(os.path.join(path, "替换结果\{}".format(file.split("\\")[-1])))
        print("{} 替换完成".format(file))
    doc.close()