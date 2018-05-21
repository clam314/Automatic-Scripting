from datetime import datetime
from datetime import timedelta
import os, time, locale, re

locale.setlocale(locale.LC_CTYPE, 'chinese')
this_year = time.localtime().tm_year

cur_dir = os.path.dirname(os.path.realpath(__file__))


def find_files(fileKey, showPrint=True):
    files = list()
    for f in os.listdir(cur_dir):
        if fileKey in f:
            if showPrint:
                print("Find an file:", f)
            files.append(f)
    if len(files) == 0:
        if showPrint:
            print('No file found!')
        return None
    else:
        return files


def find_files_in_dir(fileKey, source_dir, showPrint=True):
    files = list()
    for f in os.listdir(source_dir):
        if fileKey in f:
            if showPrint:
                print("Find an file:", f)
            files.append(f)
    if len(files) == 0:
        if showPrint:
            print('No file found!')
        return None
    else:
        return files


def get_time_by_file(file_name):
    search = re.search(r'\d{4}', file_name)
    st = str(this_year) + search.group(0)
    return datetime.strptime(st, '%Y%m%d')


def format_date(dateTime, format='%Y年%m月%d日'):
    return dateTime.strftime(format)


def dateAddStr(dateTime, day, format='%Y年%m月%d日'):
    dt_out = dateTime + timedelta(days=day)
    return dt_out.strftime(format)