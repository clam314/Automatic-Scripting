# -*- coding: utf-8 -*-
import config, amail, os,openpyxl
import common_utils as cu
import pandas as pd

doc_exp_key = '异常业务通报报表'
excel_daily_key = '收入异常名单'

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
        alln = adn+audn+wdn+wudn
        dy = amail.DailyInfo(dateTime,adn,audn,wdn,wudn,alln)
        dailyList.append(dy)
    return dailyList

def handleEmail():
    cf = config.ConfigHandle("./config/email.ini")
    sh, sp, sslp = cf.getEmailInfo()
    su, spd, sn = cf.getUser()
    smail = amail.Amail(sh, sp, sslp, su, spd, sn)
    m_addr, c_addr = cf.getExpReceiver()
    print(m_addr, c_addr)
    smail.sendExpEmail(m_addr, c_addr, _find_exp_doc(doc_exp_key))
    smail.sendDailyEmail(m_addr,c_addr,_createDailyInfo())
    smail.smtpQuit()

handleEmail()