import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

class SendMail():
    from_addr = "12345678@qq.com"  #发件人
    pwd = "zjdrhdrynorbfib"   #密码/授权码
    to_addr = "987654321@qq.com"  #收件人
    smtp_server = "smtp.qq.com"

    def __init__(self, subject, content, filename=''):
        self.subject = subject  #主题
        self.content = content  #正文
        self.filename = filename  #附件文件名/路径

    def compose(self):
        '''
            使用MIMEMultipart存储邮件内容
        '''
        msg = MIMEMultipart()
        msg["From"] = SendMail.from_addr
        msg["To"] = SendMail.to_addr

        #邮件主题
        msg["Subject"] = self.subject

        #邮件正文内容
        content = MIMEText(self.content)
        msg.attach(content)

        #邮件附件
        if self.filename:
            enclosure = MIMEApplication(open(self.filename, 'rb').read())
            enclosure.add_header('Content-Disposition', 'attachment', filename=self.filename)
            msg.attach(enclosure)    

        #将MIMEText转成str
        self.msg = msg.as_string()

    def send(self):
            self.compose()
            smtp_server = SendMail.smtp_server
            server = smtplib.SMTP(smtp_server,25)
            server.login(SendMail.from_addr, SendMail.pwd)
            server.sendmail(SendMail.from_addr, SendMail.to_addr, self.msg)
            server.close()


# obj_1 = SendMail('主题', '正文内容')
# obj_1.send()

# obj_2 = SendMail('主题', '正文内容', '附件名')
# obj_2.send()

