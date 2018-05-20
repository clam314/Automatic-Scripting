import common_utils as cu
import os,re,time
import handle_excel,create_daily_doc

createDoc = True
exp_source_key = '每日需求'
excelFiles = cu.find_files(exp_source_key)
if excelFiles == None:
    print('No handle and exit!')
    os._exit(0)

for f in excelFiles:
    search = re.search(r'\d{4}', f)
    st = f
    if search:
        st = search.group(0)
    w_out = '收入异常监控%s.xlsx' % st
    w_out_2 = '收入异常名单%s.xlsx' % st
    #处理源数据生成异常业务名单
    handle_excel.handle_excel(f, w_out, w_out_2)
    #根据异常业务名单生成通报表
    if createDoc : 
        create_daily_doc.handel_exp2doc(w_out_2)