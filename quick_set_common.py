# -*- coding: utf-8 -*-
##命令行传参，自动根据参数数目选择对应函数。
# 比对两个Excel文件内容的差异
# ---------------------假设条件----------------
# 1、源表和目标表格式一致
# 2、不存在合并单元格
# 3、excel第1行为表头，第2行开始比对,其他统计数据也仅针对从第2行开始的数据
# 4、适用于数据为单列，依据此列对数据进行取交集、并集、差集等（例如：数据表中只有学号一列，求两表学号的交集、并集、差集）
# ---------------------------------------------

import xlrd
import getopt
import sys

def read_excel(ori_path, tar_path, sheet_name):
    wb_ori = xlrd.open_workbook(ori_path)  # 打开原始文件
    wb_tar = xlrd.open_workbook(tar_path)  # 打开目标文件
    sheet_ori = wb_ori.sheet_by_name(sheet_name)
    sheet_tar = wb_tar.sheet_by_name(sheet_name)
    if sheet_ori.name == sheet_tar.name:
        # sheet表名
        if sheet_ori.name == sheet_name:
            print('表名一致')
        if sheet_ori.row_values(0) == sheet_ori.row_values(0):
            print('表头列名一致')
    for col_name in sheet_ori.row_values(0):
        diff_single(sheet_ori, sheet_tar, sheet_name, col_name)

# def read_excel(ori_path, tar_path, sheet_name,col_name):
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
#     diff_single(sheet_ori, sheet_tar, sheet_name, col_name)

def diff_single(sheet_ori, sheet_tar, sheet_name,col_name):
    ori_set = set()
    tar_set = set()
    col_index = sheet_ori.row_values(0).index(col_name)
    # 第一行为表头，所以从第二行开始进行数据获取
    for rows in range(1, sheet_ori.nrows):
        origin_data = sheet_ori.row_values(rows)[col_index]  # 源表第rows行第col_index列数据
        ori_set.add(origin_data)   # 将源表第rows行第col_index列数据写入集合ori_set
    for rows in range(1, sheet_tar.nrows):
        target_data = sheet_tar.row_values(rows)[col_index] # 目标表第rows行第col_index列数据
        tar_set.add(target_data)   # 将目标表第rows行第col_index列数据写入集合tar_set
    inter_set = ori_set & tar_set
    diff_set = ori_set ^ tar_set
    print('两表%s列交集共有%d条，分别为：' % (col_name,len(inter_set)))
    print(inter_set)
    print('两表%s列差集共有%d条，分别为：' % (col_name,len(diff_set)))
    print(diff_set)

def main():
    pass


if __name__ == '__main__':
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hi:p:", ["help", "ip=", "port="])
        print(opts)
        print(args)
        for opt_name, opt_value in opts:
            if opt_name in ('-h', '--help'):
                print('parameter:"-i <ip_address> -p <port>" or "--ip=<ip_address> --port=<port> "')
                sys.exit()
            if opt_name in ('-i', '--ip'):
                print('ip_address=' + opt_value)
            if opt_name in ('-p', '--port'):
                print('port=' + opt_value)
                sys.exit()
    except getopt.GetoptError:
        print('parameter error,please type like this:python <script_name> -i <ip_address> -p <port>')
        sys.exit()

    # read_excel(r'C:\Users\bruce\Desktop\test1.xls', r'C:\Users\bruce\Desktop\test2.xls', 'Sheet1')