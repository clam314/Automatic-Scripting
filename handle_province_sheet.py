from __future__ import division
import pandas as pd
import openpyxl, time
import pd_util as pdu


class Province(object):

    sourceHeader = [
        '日期', '省份', '应用名称', '应用ID', 'AP代码', 'AP名称', '前日订购人数', '前日订单数',
        '前日计费金额', '当日订购人数', '当日订单数', '当日计费金额'
    ]

    summary_rule = [ 
        ['省份', 'count', '计费省份个数'], 
        ['当日计费金额', 'sum', '当日总计费金额'], 
        ['当日计费金额', 'max', 'TOP省份金额']
    ]

    index = '应用ID'

    def __init__(self, *args, **kwargs):
        Province.theshold = self.__create_thsehold_list()
        self.excelWriter = pd.ExcelWriter('log.xlsx')

    def create_sheet(self, file_name, sheet_name='分省分应用流水', header=1):
        df = self.__handleHeader(
            pd.read_excel(file_name, sheet_name=sheet_name, header=header))
        dateList = df['日期'].unique()
        print(dateList)
        df_list_by_day = list()
        for day in dateList:
            df_list_by_day.append(df[df['日期'] == day])
        df_list_by_day[0].to_excel(self.excelWriter, encoding='utf-8', sheet_name='df_list_by_day')

        self.inda(df_list_by_day[0])
        self.excelWriter.save()
        
        # r =  self.__creat_sheet_by_date(df_list_by_day[0])
        # r.to_excel(self.excelWriter, encoding='utf-8', sheet_name='r')
        # self.excelWriter.save()

    def __handleHeader(self, table):
        table.columns = Province.sourceHeader
        return table

    def __clean_data(self, table):
        return table[table['省份'] == '未知']

    def __creat_sheet_by_date(self, table):
        concat_result = self.__filter_columns_values(table)
        result_list = self.__summary_columns_values(table, Province.index, Province.summary_rule) 
        for i in range(0, len(result_list)):
            concat_result = pdu.vlookup(concat_result, result_list[i], Province.index)
            print(concat_result.columns.values)
        self.__calculate_columns_values(concat_result)
        return concat_result

    #对每个字段进行汇总，需要进行汇总的数据在这里处理
    def __summary_columns_values(self, table, index, rule):
        resultList = list()
        for r in rule:
            result = pdu.get_tb_after_groupby(table, index, r[0], r[1], r[2])
            resultList.append(result)
        return resultList

    #根据表的现有字段计算新字段的值
    def __calculate_columns_values(self, table):
        table['TOP省份金额占比'] = table['TOP省份金额'] / table['当日总计费金额']
        table['占比阈值'] = table['计费省份个数'].apply(
            lambda x: self.__match_threshold_by_num(x))
        table['较占比阈值增长部分'] = table['TOP省份金额占比'] - table['占比阈值']
        # table['当日新增金额'] = table['当日计费金额'] - table['前日计费金额']
        # print(table)

    def inda(self,table):
        tb = pdu.get_tb_after_groupby(table, Province.index, '当日计费金额', 'max', 'MAX金额')
        tb.to_excel(self.excelWriter, encoding='utf-8', sheet_name='inda_tb',index=False)
        tbb = pdu.vlookup(table,tb,Province.index)
        tbb.to_excel(self.excelWriter, encoding='utf-8', sheet_name='inda_tbb',index=False)
        new_tb = tbb[tbb['当日计费金额'] == tbb['MAX金额']]
        new_tb.to_excel(self.excelWriter, encoding='utf-8', sheet_name='new_tb',index=False)
        # return new_tb

    def __filter_columns_values(self, table):
        #降序去重
        return table.sort_values(by='当日计费金额',ascending=False).drop_duplicates(
            subset=['应用ID', '日期'], keep='first').rename(columns={'省份': 'TOP收入省份'})
    
    def __change_values_format(self, table):
        table['TOP省份金额占比'] = table['TOP省份金额占比'].apply(
            lambda x: "{:.2%}".format(x))
        table['占比阈值'] = table['占比阈值'].apply(lambda x: "{:.2%}".format(x))
        table['较占比阈值增长部分'] = table['较占比阈值增长部分'].apply(
            lambda x: "{:.2%}".format(x))
        return table

    def __match_threshold_by_num(self, num):
        return Province.theshold[num]

    theshold = list()

    def __create_thsehold_list(self):
        t = list()
        for i in range(0, 32):
            if i <= 1:
                t.append(2.0)
            elif i <= 3:
                t.append(0.8)
            elif i <= 6:
                t.append(0.7)
            elif i <= 12:
                t.append(0.6)
            elif i <= 18:
                t.append(0.5)
            elif i <= 24:
                t.append(0.4)
            else:
                t.append(0.3)
        return t
