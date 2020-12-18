# -*- coding: utf-8 -*-
'''
#脚本作用：读取AC数据写入到mysql数据库中
#mysql写库参数格式：AC参数：usr_name/passwd/ip/device_type；数据库参数：usr_name:passwd@ip:port/SID；
#数据库通过pymysql+sqlalchemy进行连接，所以参数依照SQLAlchemy 1.4文档规范
#mysql写库参数示例：python script.py user_name/123456/192.168.1.1/huawei scott:123456@localhost:1521/XE
'''
from netmiko import ConnectHandler
from netmiko.ssh_exception import NetMikoTimeoutException
from paramiko.ssh_exception import AuthenticationException
import pymysql
import pandas as pd
import sqlalchemy
import time
import sys
from io import StringIO
#####################获取AC数据##########################
def get_ACdata(ac_para_str):
    ac_para=ac_para_str.split('/')
    print(ac_para)
    try:
        H3C = {
            'device_type':ac_para[3],
            'host':ac_para[2],
            'username':ac_para[0],
            'password':ac_para[1],
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
        #使用pandas.read_csv处理时，"AP name"会被拆分成两列，造成比实际数据多出一列，所以使用AP_name将其替换处理。
        output1 = output1.replace('AP name', 'AP_name')
        output_iostr= StringIO(output1)
        df = pd.read_csv(output_iostr, sep='\s+')
        # 也可以使用下述语句指定DataFrame的列名，（针对带表头的数据，header=0表示第一行为表头，然后再用names覆盖表头）
        # df = pd.read_csv(output_iostr, sep='\s+', names=['AP_name','Clients','2.4GHz','5GHz'],header=0)
        net_connect.disconnect()
        return df
    except (EOFError, NetMikoTimeoutException):
        print('Can not connect to Devices 无法登录到设备')
    except AuthenticationException:
        print('username or password wrong! 用户名或密码错误')
    except (ValueError, AuthenticationException):
        print('enable password wrong ! 或设备IP地址错误')

############################将AC数据写入数据库################################
def wr_mysql(conn_para_str,df):
    #增加Time列，并转化为datetime类型
    df.insert(0,'Time',dt)
    df["Time"] = pd.to_datetime(df["Time"])
    mysql_engine=sqlalchemy.create_engine('mysql+pymysql://'+ conn_para_str,echo=True, encoding='utf-8')
    # 利用sqlalchemy执行原生sql语句，通过Engine.connect()方法，也可以直接使用Engine方法本身
    # 示例1,通过Engine.connect()方法：
    #     connection = engine.connect()
    #     result = connection.execute("select username from users")
    #     for row in result:
    #         print("username:", row['username'])
    #     connection.close()
    # 示例2,直接使用Engine方法：
    #     result = engine.execute("select username from users")
    #     for row in result:
    #         print("username:", row['username'])
    #
    # 其中name表示SQL表名称，如果没有会新建；con表示sqlalchemy.engine.Engine；if_exists='append'表示如果表已存在，将新值插入现有表。
    # index=False表示不将DataFrame的索引列写入mysql，默认为True，表示将DataFrame索引写为列，再使用index_label参数指明索引的列名
    # chunksize=10000,表示将数据拆分成chunksize大小的数据块批量插入，在不指定这个参数的时候，pandas会一次性插入dataframe中的所有记录，
    # mysql如果服务器不能响应这么大数据量的插入，就会报错"_mysql_exceptions.OperationalError: (2006, 'MySQL server has gone away')"
    df.to_sql(name='client_number', con=mysql_engine, if_exists='append', index=False, chunksize=10000)
    # 上述语句未使用dtype对字段类型进行指定，sqlalchemy会按照默认自动指定，但是可能会浪费空间，所以可以通过dtype对字段类型进行指定，如下
    # df.to_sql('client_number', mysql_engine, index=False, if_exists='append',\
    # dtype={'AP_name':sqlalchemy.types.String(length=30),'Clients': sqlalchemy.types.Integer(),\
    # '2.4GHz': sqlalchemy.types.Integer(),'5GHz': sqlalchemy.types.Integer()})
    print('mysql success login,write data')
    mysql_engine.dispose()
if __name__=="__main__":
    print ('参数列表:\nAC登录信息：%s;\n数据库信息：%s' % (sys.argv[1],sys.argv[2]))
    dt = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print('current time:' + dt)
    start = time.perf_counter()
    ap_df=get_ACdata(sys.argv[1])
    end1 = time.perf_counter()
    print('Get data Running time: %s Seconds' % (end1 - start))
    wr_mysql(sys.argv[2],ap_df)
    end2 = time.perf_counter()
    print('write mysql Running time: %s Seconds' % (end2 - end1))
    print('Total Running time: %s Seconds' % (end2 - start))