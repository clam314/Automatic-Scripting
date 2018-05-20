from __future__ import division
import pandas as pd
import pd_util as pdu
import openpyxl,math


class IncreaseIncome(object):

    sourceHeader = [
        '日期', '应用名称', '应用ID', 'AP代码', 'AP名称', '前日订购人数', '前日订单数', '前日金额',
        '当日订购人数', '当日订单数', '当日金额'
    ]

    index = '应用ID'

    basic_score = 30

    def __init__(self, excel_writer):
        self.name = '收入异增'
        self.excelWriter = excel_writer
        self.table = ''
        self.ep_table = ''
        self.statistics_columns = ['应用名称','应用ID','AP代码','AP名称','异增评分']

    def create_sheet(self, file_name, sheet_name='全国', header=1):
        self.table = pd.read_excel(
            file_name, sheet_name=sheet_name, header=header)
        self.data_get().data_clean().data_handle().data_analysis().data_output(
        )

    def data_get(self):
        self.table.columns = IncreaseIncome.sourceHeader
        return self

    def data_clean(self):
        self.table = self.table.drop_duplicates()
        # table = table[TimeIncome.index].map(lambda x:str.strip(x))
        return self

    def data_handle(self):
        table = self.table
        table['收入增长'] = table['当日金额'] - table['前日金额']
        table['环比增长'] = table['收入增长'] / table['前日金额']
        self.table = table
        return self

    def data_analysis(self):
        eb = self.table[(self.table['前日金额'] >= 5000)
                        & (self.table['环比增长'] >= 1)].copy()
        eb['环比增长率评分'] = eb['环比增长'].apply(lambda x : self.__score_proportion(x))
        eb['收入增长量评分'] = eb['收入增长'].apply(lambda x : self.__score_value(x))
        eb['异增评分'] = eb['环比增长率评分'] + eb['收入增长量评分'] + IncreaseIncome.basic_score
        self.ep_table = eb
        return self


    def __score_value(self, x):
        n = x/10000
        score = 0.75*math.pow(n,2)+3*n-3.75
        if score > 30 :
            score = 30
        elif score < 0 :
            score = 0
        return score

    def __score_proportion(self, x):
        score = 1.25*math.pow(x,2)+2.5*x-3.75
        if score > 40 :
            score = 40
        elif score < 0 :
            score = 0
        return score 
    
    def data_output(self):
        self.table.to_excel(
            self.excelWriter, encoding='utf-8', sheet_name=self.name, index=False)
        return self

    def get_exception_sheet(self):
        return self.ep_table.copy()