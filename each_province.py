from __future__ import division
import pandas as pd
import openpyxl, time
import pd_util as pdu


class ProvinceIncome(object):

    sourceHeader = [
        '日期', '省份', '应用名称', '应用ID', 'AP代码', 'AP名称', '前日订购人数', '前日订单数',
        '前日计费金额', '当日订购人数', '当日订单数', '当日计费金额'
    ]

    incomeHeader = [
        '日期', '应用名称', '应用ID', 'AP代码', 'AP名称', '前日订购人数', '前日订单数', '前日金额',
        '当日订购人数', '当日订单数', '总金额'
    ]

    index = '应用ID'

    theshold = None

    def __init__(self, excel_writer):
        self.name = '分省集中度'
        self.excelWriter = excel_writer
        self.table = ''
        self.income_tb = ''
        self.ep_table = ''
        ProvinceIncome.theshold = self.__create_thsehold_list()
        self.statistics_columns = ['应用名称', '应用ID', 'AP代码', 'AP名称', '分省评分']

    def create_sheet(self, file_name, sheet_name='分省分应用流水', header=1):
        self.table = pd.read_excel(
            file_name, sheet_name=sheet_name, header=header)
        self.income_tb = pd.read_excel(file_name, sheet_name='全国', header=1)
        self.data_get().data_clean().data_handle().data_analysis().data_output(
        )

    def data_get(self):
        self.table.columns = ProvinceIncome.sourceHeader
        self.income_tb.columns = ProvinceIncome.incomeHeader
        return self

    def data_clean(self):
        self.table = self.table.drop_duplicates()
        self.table = self.table[(self.table['省份'] != '未知')
                                & (self.table['当日计费金额'] > 0)]
        return self

    def data_handle(self):
        table = self.table
        #通关匹配当日收入最大值类选出收入最高的省份
        max_tb = pdu.get_tb_after_groupby(table, ProvinceIncome.index,
                                          '当日计费金额', 'max', 'MAX金额')
        max_tb = pdu.vlookup(table, max_tb, ProvinceIncome.index)
        max_tb = max_tb[max_tb['当日计费金额'] == max_tb['MAX金额']].drop(
            ['MAX金额'], axis=1)
        #求省份个数
        count_tb = pdu.get_tb_after_groupby(table, ProvinceIncome.index, '省份',
                                            'count', '计费省份个数')
        max_tb = pdu.vlookup(max_tb, count_tb, ProvinceIncome.index)
        #匹配总收入
        max_tb = pdu.vlookup(max_tb,
                             self.income_tb[[ProvinceIncome.index,
                                             '总金额']], ProvinceIncome.index)
        max_tb = max_tb[max_tb['总金额'] > 0]

        max_tb['TOP省份占比'] = max_tb['当日计费金额'] / max_tb['总金额']
        max_tb['占比阈值'] = max_tb['计费省份个数'].map(
            lambda x: ProvinceIncome.theshold[x])
        max_tb['较占比阈值增长部分'] = max_tb['TOP省份占比'] - max_tb['占比阈值']

        self.table = max_tb
        return self

    def data_analysis(self):
        eb = self.table[(self.table['总金额'] >= 5000)
                        & (self.table['较占比阈值增长部分'] > 0)].copy()
        eb['偏离度评分'] = (eb['计费省份个数'] * eb['TOP省份占比'] - 1.6) / 8.4 * 25
        eb['偏离度评分'] = eb['偏离度评分'].map(lambda x: x if x < 25 else 25)
        eb['收入体量评分'] = eb['总金额'].apply(
            lambda x: (x - 5000) / 45000 * 25 if (x - 5000) / 45000 * 25 < 25 else 25
        )
        eb['分省评分'] = eb['偏离度评分'] + eb['收入体量评分'] + 50
        self.ep_table = eb
        return self

    def data_output(self):
        self.table.to_excel(
            self.excelWriter,
            encoding='utf-8',
            sheet_name=self.name,
            index=False)
        self.ep_table.to_excel(
            self.excelWriter,
            encoding='utf-8',
            sheet_name='分省集中度异常',
            index=False)
        return self

    def get_exception_sheet(self):
        return self.ep_table.copy()

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
