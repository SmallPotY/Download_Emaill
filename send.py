# -*- coding:utf-8 -*-
import smtplib
from email import encoders
# from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.header import Header
# from email.utils import parseaddr,formataddr

# 第三方 SMTP 服务 登录验证
mail_host = "smtp.qq.com"  # 设置服务器
mail_port = 465
mail_user = "#@qq.com"  # 用户名
mail_pass = "#"   # 口令,QQ邮箱是输入授权码


sender = '#@qq.com'  # 发件邮箱
receivers = ['#@qq.com','#@hxh-ltd.com']  # 接收地址


# 发件内容(内容,类型,编码)
# message = MIMEText('a test for python', 'plain', 'utf-8')
# 实例化邮件对象
message=MIMEMultipart()
# 邮件发件人落款
message['From'] = Header("smallpot", 'utf-8')
# 邮件收件人落款
message['To'] = Header("收件人落款", 'utf-8')

# 邮件标题
subject = '标题测试'
message['Subject'] = Header(subject, 'utf-8')


# 添加正文内容
message.attach(MIMEText('<html><body>'
                        +'<h1>Hello</h1>'
                        +'<p>礼物<img src="cid:Imgid">'
                        +'</body></html>','html','utf-8'))


# 插入图片
# MIMEImage，只要打开相应图片，再用read()方法读入数据，指明src中的代号是多少，如这里是'Imgid’，在HTML格式里就对应输入。
with open(r'C:\Users\dell\Desktop\小图标\images\icon\jssq.png','rb') as f:
    mime=MIMEImage(f.read())
    mime.add_header('Content-ID','Imgid')
    message.attach(mime)



# 添加附件
with open(r'C:\Users\dell\Desktop\小图标\images\icon\圆通送货时效.xlsx', 'rb') as f:
    # 设置附件的MIME和文件名，这里是png类型:
    mime = MIMEBase('Application', 'xlsx', filename='abc.xlsx')
    # mime = MIMEBase('image', 'png', filename='test.png')
    # 加上必要的头信息:
    mime.add_header('Content-Disposition', 'attachment', filename='abc.xlsx')
    mime.add_header('Content-ID', '<0>')
    mime.add_header('X-Attachment-Id', '0')
    # 把附件的内容读进来:
    mime.set_payload(f.read())
    # 用Base64编码:
    encoders.encode_base64(mime)
    # 添加到MIMEMultipart:
    message.attach(mime)


try:
    smtpObj = smtplib.SMTP_SSL(mail_host, mail_port)
    smtpObj.login(mail_user, mail_pass)
    smtpObj.sendmail(sender, receivers, message.as_string())
    smtpObj.quit()
    print(u"邮件发送成功")
except Exception  as e:
    print(e)
