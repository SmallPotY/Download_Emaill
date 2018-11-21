# coding=utf-8

import poplib
import email
import datetime
import time
import os
import xlrd
import xlwt

from email.parser import Parser

from email.header import decode_header
from email.utils import parseaddr

def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

#邮件的Subject或者Email中包含的名字都是经过编码后的str，要正常显示，就必须decode
def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value
#提取文件名，去掉后缀
def GetFileName(filename):
    (filepath,tempfilename) = os.path.split(filename);
    (shotname,extension) = os.path.splitext(tempfilename);
    return shotname
#下载邮件附件
def GetmailAttachment(emailhost,emailuser,emailpass,datestr,keywords):
    host = emailhost
    username = emailuser
    password = emailpass
    keywords = keywords   #查询的邮件的关键字
    datestr = datestr     #查询的邮件日期

    currentpath=os.getcwd()#获取当前目录
    foldername = 'EmaillDownload'#文件夹名和最后输出文件名
    new_path = os.path.join(currentpath, foldername)#文件存储路径
    if os.path.exists(new_path) == False:#如果文件夹不存在，创建文件夹
        os.makedirs(new_path)

    #for 163mail，user POP3 ########
    # pop_conn = poplib.POP3(host)
    #需要验证的邮件服务
    pop_conn = poplib.POP3_SSL(host)
    # 可选:打印POP3服务器的欢迎文字:
    print(pop_conn.getwelcome())
    # 可以打开或关闭调试信息:
    #pop_conn.set_debuglevel(1)
    pop_conn.user(username)
    pop_conn.pass_(password)
    # stat()返回邮件数量和占用空间:
    print('邮件数量: %s. 已使用空间: %s' % pop_conn.stat())
    num = len(pop_conn.list()[1])  #邮件总数
    getfilesucess = 0
    #倒叙遍历邮件
    for i in range(num,0,-1):
        #poplib.rert('邮件号码')方法返回一个元组:(状态信息,邮件,邮件尺寸)
        #hdr,message,octet=server.retr(1) 读去第一个邮件信息.
        #hdr的内容就是响应信息和邮件大小比如'+OK 12498 octets'
        #message 是包含邮件所有行的列表.
        #octet 是这个邮件的内容.
        resp, lines, octets = pop_conn.retr(i)
        # lines存储了邮件的原始文本的每一行,
        # 可以获得整个邮件的原始文本:
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        # 然后解析出邮件:
        msg = Parser().parsestr(msg_content)

        date1 = time.strptime(msg.get("Date")[0:24],'%a, %d %b %Y %H:%M:%S') #格式化收件时间
        date2 = time.strftime("%Y%m%d%H%M", date1)
        #如果日期不满足，跳出循环遍历
        if date2<datestr:
            break
        #获取邮件标题和发件人
        for header in ['From', 'Subject']:
            value = msg.get(header, '')
            if value:
                if header=='Subject':
                    subject = decode_str(value)# 解析邮件标题
                    print('邮件标题：【{}】'.format(subject) )
                else:
                    hdr, addr = parseaddr(value)
                    name = decode_str(hdr)
                    fromname = u'%s' % (name)#解析发件人名称
                    fromaddr = u'%s' % (addr)#解析发件人邮箱
        print('发件人%d: %s' % (i,fromname))
        print( "=======================================")
        if keywords in subject:
            for part in msg.walk():
                filename = part.get_filename()
                #不知为何excel附件格式为application/octet-stream
                if filename: #and (contentType == 'application/vnd.ms-excel' or contentType == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
                    data = part.get_payload(decode=True)
                    os.chdir(new_path)
                    fEx = open("%s%s.xls"%(date2,fromname), 'wb')
                    fEx.write(data)
                    fEx.close()
                    print('文件下载成功')
                    getfilesucess += 1

                else:
                    print('匹配成功, 但无附件!\n')
                    pass
        else:
            print('无匹配邮件!\n')
    if getfilesucess == 0:
        print('未找到符合条件的邮件')
    else:
        if getfilesucess==1:
            print('仅有1个文件，无须合并!')
        else:
            os.chdir(new_path)
            # 返回一个列表，其中包含在目录条目的名称
            files = os.listdir(new_path)
            outfile = xlwt.Workbook()
            for fname in files:
                sourcefile=GetFileName(fname)
                print(sourcefile)
                print(type(sourcefile))
                data = xlrd.open_workbook(os.path.join(new_path, fname))#打开源表格
                table = data.sheet_by_index(0)#打开源表格sheet
                nrows = table.nrows#源表格数据数目
                newsheet = outfile.add_sheet(sourcefile)#目的表格新增sheet
                #遍历源表格中数据，写入目的表格新增sheet
                for i in range(nrows):
                    for j in range(len(table.row(i))):
                        newsheet.write(i, j,table.row(i)[j].value)
            outname='000'+foldername+'.xls'
            outfile.save(outname)
            print('合并%d个文件到%s!' % (getfilesucess,outname))

    pop_conn.quit()

if __name__ == '__main__':
    #emailhost = raw_input('请输入邮箱服务器地址: ')
    emailhost= "***"
    #emailuser = raw_input('请输入邮箱地址: ')
    emailuser= "***"
    emailpass = "***"
    #emailpass='***'
    # datestr = input('请输入起始日期(如20170401): ')
    datestr = '201811202300'
    # keywords = input('请输入关键词: ')
    keywords = "预警编号5372：肇庆圆通出仓数据"
    print('请稍后。。。')
    GetmailAttachment(emailhost,emailuser,emailpass,datestr,keywords)


