# -*- coding: utf-8 -*-
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
from email import encoders
from email.utils import parseaddr, formataddr
from datetime import datetime
from datetime import timedelta
import smtplib, re, time, os, locale
import mimetypes
import win_unicode_console

win_unicode_console.enable()
locale.setlocale(locale.LC_CTYPE, 'chinese')

exp_html_content = '''
各位好：<br/>
<p style="text-indent:21.0pt">
    %s的信控异常业务名单已梳理完毕，请查收。
</p>'''

daily_html_head = '''
<p class=MsoNormal align=left style='text-align:left'><span style='font-size:
12.0pt;font-family:"微软雅黑",sans-serif'>各位，好</span></p>
<p class=MsoNormal align=left style='text-align:left;text-indent:21.0pt'><span
lang=EN-US style='font-size:12.0pt;font-family:"微软雅黑",sans-serif'>%s</span><span
style='font-size:12.0pt;font-family:"微软雅黑",sans-serif'>的数据通报异常应用数已梳理完毕，请查收。</span></p>
<br/>
'''

sig_path = './reference/signature.html'

daily_path = './reference/daily.html'

this_year = time.localtime().tm_year


def _getMIMEType(fileName):
    mime_type = mimetypes.guess_type(fileName)[0]
    types = mime_type.split('/')
    return types[0], types[1]


def addrStr2list(strAddr):
    if strAddr == None:
        return None
    return strAddr.split(',')


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))


def _getSignature():
    return open(sig_path).read()


def _getDailyTable():
    return open(daily_path).read()


def _get_time_by_file(file_name, r_format="%Y年%m月%d日"):
    search = re.search(r'\d{4}', file_name)
    st = str(this_year) + search.group(0)
    return datetime.strptime(st, '%Y%m%d')


class Amail(object):
    def __init__(self, host, port, sslPort, user, pwd, nick):
        self._abspath = os.path.dirname(os.path.realpath(__file__))
        self._user = user
        self._pwd = pwd
        self._nick = nick
        # self._smtp = smtplib.SMTP(host=host, port=port)
        self._smtp = smtplib.SMTP_SSL(host=host, port=sslPort)
        self._smtp.login(user, pwd)

    #发送异常通报表邮件
    def sendExpEmail(self, m_addr, c_addr, files):
        main_msg = MIMEMultipart()
        # 添加邮件正文text
        timeStr = self._create_exp_time(files)
        con_html = exp_html_content % timeStr
        con_html = _getSignature() % con_html
        content_msg = MIMEText(con_html, _subtype='html', _charset='utf-8')
        main_msg.attach(content_msg)
        main_msg = self._createEmailHeader(main_msg, m_addr, c_addr,
                                           timeStr + "的异常业务通报报表")
        # 添加附件
        if not (files == None):
            for f in files:
                main_msg.attach(self._getAnnex(f))
        # 发送邮件
        self._sendEmail(addrStr2list(m_addr), addrStr2list(c_addr), main_msg)

    #发送异常应用数日报邮件
    def sendDailyEmail(self, m_addr, c_addr, daily_list):
        main_msg = MIMEMultipart()
        #将数据的时间构建成邮件内容的字符串
        subStr = daily_list[0].date_time
        if len(daily_list) > 1:
            subStr = subStr + '-' + daily_list[len(daily_list) - 1].date_time
        #构建问候语
        headStr = daily_html_head % subStr
        #构建邮件正文
        contentStr = headStr + self._create_daily_content(daily_list)
        #添加邮件签名，并完成邮件正文的创建
        con_html = _getSignature() % contentStr
        content_msg = MIMEText(con_html, _subtype='html', _charset='utf-8')
        main_msg.attach(content_msg)
        #完成邮件头的创建
        main_msg = self._createEmailHeader(main_msg, m_addr, c_addr,
                                           subStr + "的数据通报异常应用数")
        #发送邮件
        self._sendEmail(addrStr2list(m_addr), addrStr2list(c_addr), main_msg)

    def _sendEmail(self, toAddr, ccAddr, msg):
        self._smtp.set_debuglevel(1)
        self._smtp.login(self._user, self._pwd)
        addr_list = toAddr
        if not (ccAddr == None):
            addr_list = addr_list + ccAddr
        self._smtp.sendmail(self._user, addr_list, msg.as_string())

    def _createEmailHeader(self, msg, toAddr, ccAddr, sub):
        msg['From'] = _format_addr(self._nick + '<%s>' % self._user)
        msg['To'] = toAddr
        if not (ccAddr == None):
            msg['Cc'] = ccAddr
        msg['Subject'] = Header(sub, 'utf-8')
        msg['date'] = time.strftime('%a, %d %b %Y %H:%M:%S %z')
        return msg

    def _getAnnex(self, file):
        main_type, subtype = _getMIMEType(file)
        mime = MIMEBase(
            main_type, subtype, filename=Header(file, 'gbk').encode())
        mime['Content-Disposition'] = 'attachment;filename= "%s"' % Header(
            file, 'gbk').encode()
        mime.add_header('Content-ID', '<0>')
        mime.add_header('X-Attachment-Id', '0')
        mime.set_payload(open(file, 'rb').read())
        # 用Base64编码:
        encoders.encode_base64(mime)
        return mime

    def _create_exp_time(self, files):
        content = ''
        if files == None:
            return content
        sdt = _get_time_by_file(files[0])
        sdt = sdt + timedelta(days=1)
        content = sdt.strftime('%Y年%m月%d日')
        flen = len(files)
        if flen > 1:
            ldt = _get_time_by_file(files[flen - 1])
            ldt = ldt + timedelta(days=1)
            content = content + "-" + ldt.strftime('%Y年%m月%d日')
        return content

    def _create_daily_content(self, dailyList):
        daily_table = _getDailyTable()
        content = ''
        for d in dailyList:
            content = daily_table % (d.date_time, d.app_doc_num,
                                     d.app_undoc_num, d.web_doc_num,
                                     d.web_undoc_num, d.all_num) + content
        return content

    def smtpQuit(self):
        self._smtp.quit()


class DailyInfo(object):
    def __init__(self, date_time, app_doc_num, app_undoc_num, web_doc_num,
                 web_undoc_num, all_num):
        self.date_time = date_time
        self.app_doc_num = app_doc_num
        self.app_undoc_num = app_undoc_num
        self.web_doc_num = web_doc_num
        self.web_undoc_num = web_undoc_num
        self.all_num = all_num
