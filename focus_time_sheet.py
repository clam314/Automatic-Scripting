from __future__ import division
import pandas as pd
import pd_util as pdu
import openpyxl


class FocusTimeIncome(object):

    index = '应用ID'

    def __init__(self, excel_writer):
        self.name = '时间段集中度'
        self.excelWriter = excel_writer
        self.table = ''
        self.mb_table = ''
        self.ep_table = ''
        self.statistics_columns = ['应用名称','应用ID','AP代码','AP名称','时间段评分']

    def create_sheet(self, file_name, sheet_name='应用流水金额', header=1):
        self.table = pd.read_excel(
            file_name, sheet_name=sheet_name, header=header)
        self.mb_table = pd.read_excel('网页计费全量计费点（包时长）.xlsx')
        self.data_get(self.mb_table).data_clean().data_handle().data_analysis(
        ).data_output()

    def data_get(self, mb_table):
        self.table = pdu.vlookup(self.table,mb_table[[FocusTimeIncome.index,"计费点类型"]], FocusTimeIncome.index)
        self.table['计费点类型'] = self.table['计费点类型'].map(lambda x : x if x == '包时长' else '非包时长')
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
        self.ep_table['收入体量评分'] = self.ep_table['总收入'].map(
            lambda x: (x - 1000) / 49000 * 50 if (x - 1000) / 49000 * 50 < 50 else 50
        )
        self.ep_table['时间段评分'] = self.ep_table['收入体量评分'].map(lambda x: x + 50)
        return self

    def data_output(self):
        self.table.to_excel(
            self.excelWriter,
            encoding='utf-8',
            sheet_name=self.name,
            index=False)
        return self

    def get_exception_sheet(self):
        return self.ep_table.copy()