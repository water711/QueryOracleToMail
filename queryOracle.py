import cx_Oracle
import xlwt
import time
import xlsxwriter
from sendmail import SendMail


def write_xls(title, result):
    date = time.strftime('%Y-%m-%d',time.localtime()) #获取当前日期
    file_name = date + '.xlsx'
    workbook = xlsxwriter.Workbook(file_name)  #新建excel，以日期命名

    #标题格式：加粗、字体16号、居中
    title_format = workbook.add_format(
        {'bold': True, 'font_size': 16, 'align': 'center'})
    
    #正文内容格式：字体14号、居中
    content_format = workbook.add_format(
        {'font_size': 14, 'align': 'center'})

    #日期格式：字体14号、居中，单元格格式为日期
    date_format = workbook.add_format(
        {'font_size': 14, 'align': 'center', 'num_format': 'yyyy-mm-dd hh:mm:ss'})  

    #添加工作表
    worksheet = workbook.add_worksheet("sheet1")
    
    # 设置工作表前四列的宽度
    worksheet.set_column('A:D', 30)

    #写入标题
    row = 0
    for col, content in enumerate(title):
        worksheet.write(row, col, content, title_format)

    #写入内容
    for row, record in enumerate(result):
        for col, content in enumerate(record):
            worksheet.write(row+1, col, content, content_format)

    workbook.close()
    return file_name
  

def query(user, pwd, host, dbname, port, sql):
    dsn = cx_Oracle.makedsn(host, port, dbname)
    connection = cx_Oracle.connect(
        user, pwd, dsn, encoding="UTF-8", nencoding="UTF-8")
    cursor = connection.cursor()

    cursor.execute(sql)  #执行sql查询
    result = cursor.fetchall()  #返回查询结果
    count = cursor.rowcount  #返回查询到的记录总数

    cursor.close()
    connection.close()
    return count, result

#数据库连接信息
user = "scott"
pwd = "tiger"
host = "172.10.10.1"
port = 1521
dbname = "test"
sql = 'select vipcode, name, mail, time from vipinfo'

#调用query函数，查询
count, result = query(user, pwd, host, dbname, port, sql)

#打印查询结果
print("=====================")
print("Total:", count)
print("=====================")

for row in result:
    print(row)

#将查询结果导出为excel表
title = ['会员号', '姓名', '邮箱', '注册时间']
file_name = write_xls(title, result)  

#发送邮件
mail = SendMail('xx报表', '请查看附件', file_name)
mail.send()
