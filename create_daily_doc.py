from __future__ import division
from mailmerge import MailMerge
from datetime import datetime
from datetime import timedelta
import pandas as pd
import openpyxl, time,os, re,locale
import win_unicode_console

win_unicode_console.enable()
locale.setlocale(locale.LC_CTYPE,'chinese')

file_key = '收入异常名单'
doc_template = "./reference/异常业务通报报表模板.docx"
this_year = time.localtime().tm_year
out_template_format = './output/异常业务通报报表'


out_columns = ['应用名称','应用ID','AP代码',"AP名称",'异常维度个数']

def _get_time_by_file(file_name,r_format="%Y年%m月%d日"):
    search = re.search(r'\d{4}', file_name)
    st = str(this_year)+search.group(0)
    return datetime.strptime(st,'%Y%m%d')

def handel_exp2doc(file_name):
    doc = MailMerge(doc_template)
    dt = _get_time_by_file(file_name)
    df_info = pd.read_excel(file_name, sheet_name='涉及统计').iloc[0]
    df_detail = pd.read_excel(file_name, sheet_name='异常统计')
    _replace_info(doc,df_info,dt)
    _replace_detail_table(doc,df_detail[out_columns])
    dt_out = dt + timedelta(days=1)
    w_file_name = '{0}{1}.docx'.format(out_template_format,dt_out.strftime('%m%d'))
    doc.write(w_file_name)


def _replace_info(doc,info,dt):
    ap_num = '{:.0f}'.format(info['AP数目'])
    app_num = '{:.0f}'.format(info['应用数目'])
    amount = '{:.2f}'.format(info['涉及总金额']/10000)
    data_time = dt.strftime('%Y年%m月%d日')
    head = {
        'time':data_time,
        'ap_num':ap_num,
        'app_num':app_num,
        'amount':amount
    }
    doc.merge(**head)

def _replace_detail_table(doc,detail):
    table = list()
    detail = detail.sort_values(by='异常维度个数',ascending=False)
    detail = detail.applymap(lambda x : str(x))
    for index, row in detail.iterrows():
        table.append(dict(row))
    doc.merge_rows('应用ID',table)
