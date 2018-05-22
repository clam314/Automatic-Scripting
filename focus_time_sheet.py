from __future__ import division
import pandas as pd
import pd_util as pdu
import openpyxl,math

ref_file_name = None

class FocusTimeIncome(object):

    index = '应用ID'

    basic_score = 0
    
    def __init__(self, excel_writer,scoring=True):
        self.name = '时间段集中度'
        self._scoring = scoring
        self.excelWriter = excel_writer
        self.table = ''
        self.mb_table = ''
        self.ep_table = ''
        self.statistics_columns = ['应用名称', '应用ID', 'AP代码', 'AP名称', '时间段评分']

    def create_sheet(self, file_name, sheet_name='应用流水金额', header=1):
        self.table = pd.read_excel(
            file_name, sheet_name=sheet_name, header=header)
        self.mb_table = pd.read_excel(ref_file_name)
        self.data_get(self.mb_table).data_clean().data_handle().data_analysis(
        ).data_output()

    def data_get(self, mb_table):
        self.table = pdu.vlookup(self.table,
                                 mb_table[[FocusTimeIncome.index,
                                           "计费点类型"]], FocusTimeIncome.index)
        self.table['计费点类型'] = self.table['计费点类型'].map(
            lambda x: x if x == '包时长' else '非包时长')
        return self

    def data_clean(self):
        self.table = self.table.drop_duplicates()
        # table = table[TimeIncome.index].map(lambda x:str.strip(x))
        return self

    def data_handle(self):
        table = self.table
        time_frame = list(range(0, 24))
        table['TOP6小时'] = table[time_frame].apply(
            lambda x: x.sort_values(ascending=False)[0:6].sum(), axis=1)
        table['总收入'] = table[time_frame].apply(lambda x: x.sum(), axis=1)
        table['TOP6小时占比'] = table['TOP6小时'] / table['总收入']
        return self

    def data_analysis(self):
        tb = self.table
        self.ep_table = tb[(tb['计费点类型'] != '包时长') & (tb['总收入'] >= 1000) &
                           (tb['TOP6小时占比'] >= 0.95)].copy()
        if self._scoring:
            self.ep_table['收入总量评分'] = self.ep_table['总收入'].map(
                lambda x: self.__score_value(x))
            self.ep_table['时间段占比评分'] = self.ep_table['TOP6小时占比'].map(
                lambda x: self.__score_proportion(x))
            self.ep_table[
                '时间段评分'] = self.ep_table['收入总量评分'] + self.ep_table['时间段占比评分'] + FocusTimeIncome.basic_score
        return self

    #收入总量评分
    def __score_value(self, x):
        score = (x-1000)/49000*15
        return score

    #时间段占比评分
    def __score_proportion(self, x):
        n = x * 100 - 95
        score = 9/5 * math.pow(n, 2)
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