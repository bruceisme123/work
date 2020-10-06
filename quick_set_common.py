# -*- coding: utf-8 -*-
'''
脚本用途：比对两个Excel文件内容的差异
传参规则：“python quick_set_common.py -1 源excel文件名 -2 目标excel文件名 -s 比较的sheet表格名称 -c 比较的列名”；
或者
“python quick_set_common.py --src=源excel文件名 --tar=目标excel文件名 --sheet=比较的sheet表格名称 --col=比较的列名”
---------------前提条件----------------
1、源表和目标表格式一致,表头列名一致
2、不存在合并单元格
3、excel第1行为表头，从第2行开始比对,其他统计数据也仅针对从第2行开始的数据
4、适用于数据为单列，依据此列对数据进行取交集、并集、差集等（例如：数据表中只有学号一列，求两表学号的交集、并集、差集）
5、指定比较的列名，但是输出为每一行数据的list，指定列的数据一致，即认为一致，且此时取源表中对应行的数据。
---------------------------------------------
结果说明：因为字典dict的本身无序，且做了集合操作，所以结果文件顺序与源文件不一致。
---------------------------------------------
'''
import xlrd
import xlsxwriter
import getopt
import sys

# def read_excel(ori_path, tar_path, sheet_name):
#     wb_ori = xlrd.open_workbook(ori_path)  # 打开原始文件
#     wb_tar = xlrd.open_workbook(tar_path)  # 打开目标文件
#     sheet_ori = wb_ori.sheet_by_name(sheet_name)
#     sheet_tar = wb_tar.sheet_by_name(sheet_name)
#     if sheet_ori.name == sheet_tar.name:
#         # sheet表名
#         if sheet_ori.name == sheet_name:
#             print('表名一致')
#         if sheet_ori.row_values(0) == sheet_ori.row_values(0):
#             print('表头列名一致')
#     for col_name in sheet_ori.row_values(0):
#         print("正在比对的列名为" % col_name)
#         diff_single(sheet_ori, sheet_tar, col_name)

# def read_excel(ori_path, tar_path):
#     wb_ori = xlrd.open_workbook(ori_path)  # 打开原始文件
#     wb_tar = xlrd.open_workbook(tar_path)  # 打开目标文件
#     for sheet_name in wb_ori.sheet_names():
#         print("开始比对sheet：%s" % sheet_name)
#         sheet_ori = wb_ori.sheet_by_name(sheet_name)
#         sheet_tar = wb_tar.sheet_by_name(sheet_name)
#         if sheet_ori.name == sheet_tar.name:
#             # sheet表名
#             if sheet_ori.name == sheet_name:
#                 print('表名一致')
#             if sheet_ori.row_values(0) == sheet_ori.row_values(0):
#                 print('表头列名一致')
#         for col_name in sheet_ori.row_values(0):
#             diff_single(sheet_ori, sheet_tar, col_name)

def read_excel(ori_path, tar_path, sheet_name, col_name):
    global biaotou_list
    wb_ori = xlrd.open_workbook(ori_path)  # 打开原始文件
    wb_tar = xlrd.open_workbook(tar_path)  # 打开目标文件
    sheet_ori = wb_ori.sheet_by_name(sheet_name)
    sheet_tar = wb_tar.sheet_by_name(sheet_name)
    if sheet_ori.name == sheet_tar.name:
        # sheet表名
        if sheet_ori.name == sheet_name:
            print('表名一致')
        biaotou_list = sheet_ori.row_values(0)  #表头以源表为准
        if sheet_ori.row_values(0) == sheet_tar.row_values(0):
            print('表头列名一致')
        else:
            print('表头列名不一致，以源表为准')
    diff_single(sheet_ori, sheet_tar, col_name)


def diff_single(sheet_ori, sheet_tar, col_name):
    global inter_list
    global diff_list
    global union_list
    global diff1_list
    global diff2_list
    ori_dict = {}
    tar_dict = {}
    # 获取对比列的index
    col_index = sheet_ori.row_values(0).index(col_name)
    # 第一行为表头，所以从第二行开始进行数据获取
    for rows in range(1, sheet_ori.nrows):
        # 第rows行中，对比列的值作为字典的key（从而可以实现本表中列名值的自动去重），整行数据组成的list作为字典的value
        ori_list = sheet_ori.row_values(rows)  # 源表i行数据
        ori_dict[sheet_ori.row_values(rows)[col_index]] = ori_list
    for rows in range(1, sheet_tar.nrows):
        tar_list = sheet_tar.row_values(rows)
        tar_dict[sheet_tar.row_values(rows)[col_index]] = tar_list
    inter_keys = ori_dict.keys() & tar_dict.keys()
    diff_keys = ori_dict.keys() ^ tar_dict.keys()
    union_keys = ori_dict.keys() | tar_dict.keys()
    diff1_keys = ori_dict.keys() - tar_dict.keys()  # 只在源表中存在的key
    diff2_keys = tar_dict.keys() - ori_dict.keys()  # 只在目的表中存在的key
    print('两表%s列交集共有%d条，分别为：' % (col_name, len(inter_keys)))
    print(inter_keys)
    for key in inter_keys:
        inter_list.append(ori_dict[key])
    print('两表%s列差集共有%d条，分别为：' % (col_name, len(diff_keys)))
    print(diff_keys)
    for key in diff_keys:
        if ori_dict.get(key):
            diff_list.append(ori_dict[key])
        else:
            diff_list.append(tar_dict[key])
    print('两表%s列并集共有%d条，分别为：' % (col_name, len(union_keys)))
    print(union_keys)
    for key in union_keys:
        if ori_dict.get(key):  # 不可使用ori_dict[key]，因为如果key不存在，会报错KeyError
            union_list.append(ori_dict[key])
        else:
            union_list.append(tar_dict[key])
    print('两表%s列，只在源表中存在的有%d条，分别为：' % (col_name, len(diff1_keys)))
    print(diff1_keys)
    for key in diff1_keys:
        diff1_list.append(ori_dict[key])
    print('两表%s列，只在目的表中存在的有%d条，分别为：' % (col_name, len(diff2_keys)))
    print(diff2_keys)
    for key in diff2_keys:
        diff2_list.append(tar_dict[key])
    return biaotou_list, inter_list, diff_list, union_list, diff1_list, diff2_list


def write_excel(biaotou_list, inter_list, diff_list, union_list, diff1_list, diff2_list):
    workbook = xlsxwriter.Workbook('result.xlsx')
    write_para = [['交集', inter_list], ['差集', diff_list], ['并集', union_list], ['表1-交集', diff1_list], ['表2-交集', diff2_list]]
    for sheet in write_para:
        addsheet = workbook.add_worksheet(sheet[0])
        addsheet.write_row('A1', biaotou_list)
        row_num = 2
        for rows in sheet[1]:
            tar_row = 'A' + str(row_num)
            addsheet.write_row(tar_row, rows)
            row_num += 1
    workbook.close()


if __name__ == '__main__':
    # 参数获取，参考https://www.cnblogs.com/yuandonghua/p/10619941.html
    try:
        opts, args = getopt.getopt(sys.argv[1:], "h1:2:s:c:", ["help", "src=", "tar=", "sheet=", "col="])
        print(opts)
        print(args)
        for opt_name, opt_value in opts:
            if opt_name in ('-h', '--help'):
                print('传参规则：“python quick_set_common.py -1 源excel文件名 -2 目标excel文件名 -s 比较的sheet表格名称 -c 比较的列名”；'
                      '或者“python quick_set_common.py --src=源excel文件名 --tar=目标excel文件名 --sheet=比较的sheet表格名称 --col=比较的列名”')
                sys.exit()
            if opt_name in ('-1', '--src'):
                ori_path = opt_value
                print('srcfile=' + ori_path)
            if opt_name in ('-2', '--tar'):
                tar_path = opt_value
                print('tarfile=' + tar_path)
            if opt_name in ('-s', '--sheet'):
                sheet_name = opt_value
                print('tarfile=' + sheet_name)
            if opt_name in ('-c', '--col'):
                col_name = opt_value
                print('tarfile=' + col_name)
    except getopt.GetoptError:
        print('parameter error,please type like this:python <script_name> -i <ip_address> -p <port>')
        sys.exit()
    #正式处理
    biaotou_list = []
    inter_list = []
    diff_list = []
    union_list = []
    diff1_list = []
    diff2_list = []
    read_excel(ori_path, tar_path, sheet_name, col_name)
    # read_excel(r'C:\Users\bruce\Desktop\新.xls', r'C:\Users\bruce\Desktop\旧.xls', 'Sheet1', '学号')
    write_excel(biaotou_list, inter_list, diff_list, union_list, diff1_list, diff2_list)