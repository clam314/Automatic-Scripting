from __future__ import division
import pandas as pd
import openpyxl, time, os,re
import each_province, focus_time_sheet, e_morning_sheet, increase_sheet
import pd_util as pdu

file_key = '每日需求'


def find_excel():
    curr_dir = os.path.dirname(os.path.realpath(__file__))
    files = list()
    for f in os.listdir(curr_dir):
        if file_key in f:
            print("Find an excel file:", f)
            files.append(f)
    if len(files) == 0:
        print('No excel file found!')
        os._exit(0)
    else:
        return files


def handle_excel(r_file_name, w_file_name, w_file_name_2):
    startTime = time.time()
    excelWriter = pd.ExcelWriter(w_file_name)

    sheetList = list()
    sheetList.append(increase_sheet.IncreaseIncome(excelWriter))
    sheetList.append(e_morning_sheet.EhoursIncome(excelWriter))
    sheetList.append(focus_time_sheet.FocusTimeIncome(excelWriter))
    sheetList.append(each_province.ProvinceIncome(excelWriter))
    [x.create_sheet(r_file_name) for x in sheetList]

    exp_tbs = list()
    [exp_tbs.append(x.get_exception_sheet()) for x in sheetList]
    statistics_tb = exp_tbs[0][list(sheetList[0].statistics_columns[0:4])]
    for i in range(1, len(exp_tbs)):
        statistics_tb = statistics_tb.append(
            exp_tbs[i][sheetList[i].statistics_columns[0:4]])
    statistics_tb = statistics_tb.drop_duplicates()
    statistics_tb.reset_index(drop=True)
    for i in range(0, len(exp_tbs)):
        ll = list(sheetList[i].statistics_columns)
        statistics_tb = pdu.vlookup(statistics_tb, exp_tbs[i][[ll[1], ll[4]]],
                                    '应用ID')
    count_list = list()
    [count_list.append(s.statistics_columns[4]) for s in sheetList]

    statistics_tb['异常维度个数'] = statistics_tb[count_list].apply(
        lambda x: x.count(), axis=1)
    statistics_tb['评分'] = statistics_tb[count_list].apply(
        lambda x: x.max(), axis=1)

    s_ew = pd.ExcelWriter(w_file_name_2)
    for i in range(0, len(exp_tbs)):
        exp_tbs[i].to_excel(
            s_ew, encoding='utf-8', sheet_name=sheetList[i].name, index=False)
    statistics_tb.to_excel(
        s_ew, encoding='utf-8', sheet_name='统计', index=False)
    s_ew.save()
    excelWriter.save()
    print('totaltime:', time.time() - startTime)


fileExcels = find_excel()
for f in fileExcels:
    search = re.search(r'\d{4}',f)
    st = f
    if search :
        st = search.group(0)
    w_out = '收入异常监控' + st + '.xlsx'
    w_out_2 = '收入异常名单' + st + '.xlsx'
    handle_excel(f, w_out, w_out_2)
