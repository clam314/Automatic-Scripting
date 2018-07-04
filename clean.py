import os,time
import common_utils as cu

cu.cur_dir = os.path.dirname(os.path.realpath(__file__))

#查找当前目录所有excel文件并删除
all_excel = cu.find_files('.xlsx',showPrint=False)
if not(all_excel == None):
    for excel in all_excel:
        os.remove(excel)
        print("delete:",excel)
#查找output目录所有word文件并删除
all_word = cu.find_files('.docx',showPrint=False)
if not(all_word == None):
    for word in all_word:
        os.remove(word)
        print("delete:",word)

print("path",cu.cur_dir)
time.sleep(2)