import jieba.analyse
import os
import xlrd


def read_file(root, file):
    # 打开一个workbook
    wb = xlrd.open_workbook(root + "/" + str(file))
    # 按照索引打开文件
    booksheet = wb.sheet_by_index(0)

    row = 2  # 开始行数
    cells = []  # 存储工作要求的数据
    row_num = booksheet.nrows  # 获取行数
    first_categoty = booksheet.cell_value(2, 0)
    second_category = booksheet.cell_value(2, 1)
    # print(row_num)
    while row < row_num:
        cell = booksheet.cell_value(row, 3)
        cells.append(cell)
        row = row + 1
    return cells
    # 打印所有cell中的数据
    # for cell in cells:
    #     print(cell)


def key_analyse(cells, root, file):
    data = " ".join(cells)
    # 加载停用词
    jieba.analyse.set_stop_words("stopwords")
    # 加载自定义用户词典
    jieba.load_userdict("dict")
    # 关键字分析
    key = jieba.analyse.extract_tags(data, topK=5, withWeight=False, allowPOS=('n', 'nr'))
    print(file.rstrip(".xlst"))
    print(key)
    with open(root+".txt", "a+") as f:
        f.write(file.rstrip(".xlst")+"   "+"   ".join(key)+"\n")


# 读取目录下所有文件
for root, dirs, files in os.walk('infos/高端职位'):
    # print(root)  # 当前目录路径
    # print(dirs)  # 当前路径下所有子目录
    print(files)  # 当前路径下所有非目录子文件
    for file in files:
        cells = read_file(root, file)
        # print(cells)
        key_analyse(cells, root, file)
