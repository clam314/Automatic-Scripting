from __future__ import division
import pandas as pd
import pd_util as pdu
import openpyxl, math


class EhoursIncome(object):

    index = '应用ID'

    basic_score = 0

    def __init__(self, excel_writer,scoring=True):
        self.name = '闲时集中度'
        self._scoring = scoring
        self.excelWriter = excel_writer
        self.table = ''
        self.ep_table = ''
        self.statistics_columns = ['应用名称', '应用ID', 'AP代码', 'AP名称', '闲时评分']

    def create_sheet(self, file_name, sheet_name='应用流水金额', header=1):
        self.table = pd.read_excel(
            file_name, sheet_name=sheet_name, header=header)
        self.data_clean().data_handle().data_analysis().data_output()

    def data_get(self, mb_table):
        return self

    def data_clean(self):
        self.table = self.table.drop_duplicates()
        # table = table[TimeIncome.index].map(lambda x:str.strip(x))
        return self

    def data_handle(self):
        table = self.table
        table['闲时收入'] = table[list(range(1, 6))].apply(
            lambda x: x.sum(), axis=1)
        table['总收入'] = table[list(range(0, 24))].apply(
            lambda x: x.sum(), axis=1)
        table['闲时占比'] = table['闲时收入'] / table['总收入']
        self.table = table
        return self

    def data_analysis(self):
        tb = self.table
        self.ep_table = tb[(tb['总收入'] >= 1000) & (tb['闲时占比'] >= 0.3)].copy()
        if self._scoring:
            self.ep_table['占比评分'] = self.ep_table['闲时占比'].map(
                lambda x: self.__score_proportion(x))
            self.ep_table[
                '闲时评分'] = self.ep_table['占比评分'] + EhoursIncome.basic_score
        return self

    #占比评分
    def __score_proportion(self, x):
        n = x * 10 - 3
        score = 100/49*math.pow(n,2)
        return score

    def data_output(self):
        self.table.to_excel(
            self.excelWriter,
            encoding='utf-8',
            sheet_name=self.name,
            index=False)
        return self

    def get_exception_sheet(self):
        return self.ep_table.copy()