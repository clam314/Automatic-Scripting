from __future__ import division
import pandas as pd
import pd_util as pdu
import openpyxl


class IncreaseIncome(object):

    sourceHeader = [
        '日期', '应用名称', '应用ID', 'AP代码', 'AP名称', '前日订购人数', '前日订单数', '前日金额',
        '当日订购人数', '当日订单数', '当日金额'
    ]

    index = '应用ID'

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
        eb['增长率评分'] = eb['环比增长'].apply(
            lambda x: (x - 1) / 4 * 25 if (x - 1) / 4 * 25 < 25 else 25)
        eb['收入体量评分'] = eb['收入增长'].apply(
            lambda x: (x - 5000) / 45000 * 25 if (x - 5000) / 45000 * 25 < 25 else 25
        )
        eb['异增评分'] = eb['增长率评分'] + eb['收入体量评分'] + 50
        self.ep_table = eb
        return self

    def data_output(self):
        self.table.to_excel(
            self.excelWriter, encoding='utf-8', sheet_name=self.name, index=False)
        return self

    def get_exception_sheet(self):
        return self.ep_table.copy()