import os
import common_utils as cu

clean_dir = '\\output'

#查找当前目录所有excel文件并删除
all_excel = cu.find_files('.xlsx',showPrint=False)
if not(all_excel == None):
    for excel in all_excel:
        os.remove(excel)
        print("delete:",excel)
#查找output目录所有word文件并删除
source_dir = cu.cur_dir+clean_dir
all_word = cu.find_files_in_dir('.docx',cu.cur_dir+clean_dir,showPrint=False)
if not(all_word == None):
    for word in all_word:
        os.remove(source_dir+'\\'+word)
        print("delete:",word)
