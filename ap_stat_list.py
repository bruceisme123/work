# -*- coding: utf-8 -*-
'''
#脚本作用：读取AC数据写入到oracle数据库中
#AC参数格式：usr_name/passwd/ip/device_type；数据库参数格式：usr_name/passwd@ip:port/SID；
#参数示例：python script.py user_name/123456/192.168.1.1/huawei scott/123456@localhost:1521/XE
'''
from netmiko import ConnectHandler
from netmiko.ssh_exception import NetMikoTimeoutException
from paramiko.ssh_exception import AuthenticationException
import cx_Oracle
import time
import sys

#####################获取AC数据##########################
def get_ACdata(ac_para_str):
    ac_para = ac_para_str.split('/')
    print(ac_para)
    try:
        H3C = {
            'device_type': ac_para[3],
            'host': ac_para[2],
            'username': ac_para[0],
            'password': ac_para[1],
        }
        net_connect = ConnectHandler(**H3C)
        print('AC login success,get data')
        temp_out1 = net_connect.send_command_timing('dis wlan ap all client-number')
        output1=temp_out1.rstrip('---- More ----')
        # print(temp_out)
        while "---- More ----" in temp_out1:
            # 遇到more，就多输入几次个空格，normalize=False表示不取消命令前后空格。
            temp_out1 = net_connect.send_command_timing(' ', strip_prompt=True, strip_command=False, normalize=False)
            temp_out1 = temp_out1.lstrip()
            output1 += temp_out1.rstrip('---- More ----')
        output1=output1.rstrip('<AC>')
        output1=output1.rstrip()
        print('dis wlan ap all client-number命令执行最终结果为:')
        print(output1)
        list1 = output1.split('\n')
        list1=list1[1:]
        tup_list1=[]
        # 将ap all client-number每行数据以空格分隔为元组，同时将每行从第2个到末尾的数据转化为整形int，以便数据库sum计算。
        # 示例：['ZHL-10-1001 1 1 0','ZHL-10-1002 3 2 1']转化为[('ZHL-10-1001',1,1,0),('ZHL-10-1002',3,2,1)]
        for row in list1:
            str_list=row.split()
            for i in range(1, len(str_list)):
                str_list[i] = int(str_list[i])
            str_tup=tuple(str_list)
            tup_list1.append(str_tup)
        net_connect.disconnect()
        return tup_list1
    except (EOFError, NetMikoTimeoutException):
        print('Can not connect to Devices 无法登录到设备')
    except AuthenticationException:
        print('username or password wrong! 用户名或密码错误')
    except (ValueError, AuthenticationException):
        print('enable password wrong ! 或设备IP地址错误')


############################将AC数据写入oracle数据库################################
def wr_oracle(conn_para_str,res):
    db = cx_Oracle.connect(conn_para_str)
    cr = db.cursor()
    print('oracle success login,write data')
    dt2= "'"+dt+"'"
    basic_sql = "insert into client_number values(to_date(%s,'yyyy-mm-dd,hh24:mi:ss'),:1,:2,:3,:4)" %dt2
    cr.executemany(basic_sql, res)
    db.commit()
    cr.close()
    db.close()


############################将AC数据写入mysql数据库################################
# # 针对python2.7使用的MySQLdb(MySQL-python数据库)
# def py2_wr_mysql(conn_para_str,res):
#     conn_para=re.split('[/@:]',conn_para_str)
#     db= MySQLdb.connect(host=conn_para[2],user=conn_para[0],passwd=conn_para[1],db=conn_para[4],port=int(conn_para[3]),charset="utf8")
#     cr = db.cursor()
#     print('mysql success login,write data')
#     dt = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
#     dt2= "'"+dt+"'"
#     basic_sql="insert into client_number values(str_to_date("+ dt2 +",'%%Y-%%m-%%d %%H:%%i:%%s'),%s,%s,%s,%s)"
#     cr.executemany(basic_sql, res)
#     db.commit()
#     cr.close()
#     db.close()

# # 针对python3使用pymysql连接数据库
# def py3_wr_mysql(conn_para_str,res):
#     conn_para=re.split('[/@:]',conn_para_str)
#     db= pymysql.connect(host=conn_para[2],user=conn_para[0],password=conn_para[1],database=conn_para[4],charset='utf8')
#     cr = db.cursor()
#     print('mysql success login,write data')
#     dt = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
#     dt2= "'"+dt+"'"
#     basic_sql="insert into client_number values(str_to_date("+ dt2 +",'%%Y-%%m-%%d %%H:%%i:%%s'),%s,%s,%s,%s)"
#     cr.executemany(basic_sql, res)
#     db.commit()
#     cr.close()
#     db.close()

if __name__=="__main__":
    print ('参数列表:\nAC登录信息：%s;\nORACLE数据库信息：%s' % (sys.argv[1],sys.argv[2]))
    dt = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print('current time:' + dt)
    start = time.perf_counter()
    res=get_ACdata(sys.argv[1])
    end1 = time.perf_counter()
    print('Get data Running time: %s Seconds' % (end1 - start))
    wr_oracle(sys.argv[2],res)
    # py2_wr_mysql(sys.argv[2],res)
    # py3_wr_mysql(sys.argv[2],res)
    end2 = time.perf_counter()
    print('write oracle Running time: %s Seconds' % (end2 - end1))
    print('Total Running time: %s Seconds' % (end2 - start))