from __future__ import division
import pandas as pd
import openpyxl, time, os, re
import each_province, focus_time_sheet, e_morning_sheet, increase_sheet
import pd_util as pdu

ref_dir = './reference/'
ref_file = ref_dir+'网页计费全量计费点（包时长）.xlsx'

def cal_score(se):
    score = se.sum()
    return score


def handle_excel(r_file_name, w_file_name, w_file_name_2):
    excelWriter = pd.ExcelWriter(w_file_name)

    focus_time_sheet.ref_file_name = ref_file

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
        lambda x: x.sum(), axis=1)

    total_income = pd.read_excel(r_file_name, sheet_name='全国', header=1)
    total_income.columns = increase_sheet.IncreaseIncome.sourceHeader
    statistics_tb = pdu.vlookup(statistics_tb, total_income[['应用ID', '当日金额']],
                                '应用ID')

    has_type_sheet = True
    try:
        cal_type_tb = pd.read_excel(r_file_name, sheet_name='计费类型')
    except Exception:
        print('not find sheet(计费类型)！')
        has_type_sheet = False
    
    if has_type_sheet:
        cal_time_tb = pd.read_excel(ref_file)
        statistics_tb = pdu.vlookup(
            statistics_tb, cal_type_tb[['应用ID', "计费类型（网页计费/IAP)"]], '应用ID')
        statistics_tb = pdu.vlookup(statistics_tb, cal_time_tb[['应用ID', "计费点类型"]],
                                    '应用ID')
        statistics_tb['计费点类型'] = statistics_tb['计费点类型'].map(
            lambda x: x if x == '包时长' else '非包时长')

    #对统计表进行统计应用个数、AP个数和金额总数
    all_income = statistics_tb[['当日金额']].apply(lambda x: x.sum())
    app_num = len(statistics_tb)
    ap_num = len(statistics_tb.drop_duplicates(['AP代码']))
    if has_type_sheet:
        web_is_time_num = len(
            statistics_tb[(statistics_tb["计费类型（网页计费/IAP)"] == '网页计费') & (statistics_tb["计费点类型"] == '包时长')])
        web_not_time_num = len(
            statistics_tb[(statistics_tb["计费类型（网页计费/IAP)"] == '网页计费') & (statistics_tb["计费点类型"] == '非包时长')])
        app_is_time_num = len(
            statistics_tb[(statistics_tb["计费类型（网页计费/IAP)"] == '应用内计费') & (statistics_tb["计费点类型"] == '包时长')])
        app_not_time_num = len(
            statistics_tb[(statistics_tb["计费类型（网页计费/IAP)"] == '应用内计费') & (statistics_tb["计费点类型"] == '非包时长')])
        total_info_tb = pd.DataFrame([{
            '应用数目': app_num,
            'AP数目': ap_num,
            '涉及总金额': all_income['当日金额'],
            '网页包时长数': web_is_time_num,
            '网页点播数': web_not_time_num,
            '应用内包时长数': app_is_time_num,
            '应用内点播数': app_not_time_num
        }])
    else:
        total_info_tb = pd.DataFrame([{
            '应用数目': app_num,
            'AP数目': ap_num,
            '涉及总金额': all_income['当日金额']
        }])

    s_ew = pd.ExcelWriter(w_file_name_2)
    statistics_tb.to_excel(
        s_ew, encoding='utf-8', sheet_name='异常统计', index=False)
    total_info_tb.to_excel(
        s_ew, encoding='utf-8', sheet_name='涉及统计', index=False)
    for i in range(0, len(exp_tbs)):
        exp_tbs[i].to_excel(
            s_ew, encoding='utf-8', sheet_name=sheetList[i].name, index=False)
    s_ew.save()
    excelWriter.save()
    pdu.change_sheet_style(w_file_name_2, '异常统计')

