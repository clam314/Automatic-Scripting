# -*- coding: utf-8 -*-
import config, amail, os, openpyxl,time
import common_utils as cu
import pandas as pd

doc_exp_key = '异常业务通报报表'
excel_daily_key = '收入异常名单'
excel_exp_key = '收入异常'


def _find_exp_doc(key):
    curr_dir = os.path.dirname(os.path.realpath(__file__))
    files = list()
    for f in os.listdir(curr_dir):
        if key in f:
            print("Find an file:", f)
            files.append(f)
    if len(files) == 0:
        print('No excel file found!')
        return None
    else:
        return files


def _createDailyInfo():
    excelList = cu.find_files(excel_daily_key)
    if excelList == None:
        return None
    dailyList = list()
    for e in excelList:
        df_info = pd.read_excel(e, sheet_name='涉及统计').iloc[0]
        dateTime = cu.dateAddStr(cu.get_time_by_file(e),1)
        adn = df_info['应用内点播数']
        audn = df_info['应用内包时长数']
        wdn = df_info['网页点播数']
        wudn = df_info['网页包时长数']
        alln = adn + audn + wdn + wudn
        dy = amail.DailyInfo(dateTime, adn, audn, wdn, wudn, alln)
        dailyList.append(dy)
    return dailyList

def handleEmail():
    cf = config.ConfigHandle("./config/email.ini")
    #获取邮件账户信息
    sh, sp, sslp = cf.getEmailInfo()
    su, spd, sn = cf.getUser()
    #登录并连接邮件服务器
    smail = amail.Amail(sh, sp, sslp, su, spd, sn)
    #获取各类邮件发送人信息
    m_addr_e, c_addr_e = cf.getExpReceiver()
    m_addr_see, c_addr_see = cf.getSynExpReceiver()
    m_addr_d, c_ddr_d = cf.getDailyReceiver()
    #发送各类邮件
    smail.sendExpEmail(m_addr_e, c_addr_e, _find_exp_doc(doc_exp_key))
    smail.sendSynExpEmail(m_addr_see, c_addr_see, cu.find_files(excel_exp_key))
    smail.sendDailyEmail(m_addr_d, c_ddr_d, _createDailyInfo())
    # #退出邮件服务器
    # smail.smtpQuit()
    

handleEmail()