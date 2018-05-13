import pandas as pd


def get_tb_after_groupby(table, index, column, method, new_columnName):
    result = table.groupby(index)[column].agg(method).reset_index()
    result.columns = [index, new_columnName]
    return result


def vlookup(out_table, in_table, index):
    return out_table.merge(in_table, on=index, how='left')