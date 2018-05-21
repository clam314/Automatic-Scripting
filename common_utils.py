import os

cur_dir = os.path.dirname(os.path.realpath(__file__))

def find_files(fileKey,showPrint=True):
    files = list()
    for f in os.listdir(cur_dir):
        if fileKey in f:
            if showPrint :
                print("Find an file:", f)
            files.append(f)
    if len(files) == 0:
        if showPrint :
            print('No file found!')
        return None
    else:
        return files

def find_files_in_dir(fileKey,source_dir,showPrint=True):
    files = list()
    for f in os.listdir(source_dir):
        if fileKey in f:
            if showPrint :
                print("Find an file:", f)
            files.append(f)
    if len(files) == 0:
        if showPrint :
            print('No file found!')
        return None
    else:
        return files