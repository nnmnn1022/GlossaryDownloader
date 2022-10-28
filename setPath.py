import os

def run(path):
    del_list = ['"', "'", '& ']
    for char in del_list :
        path = path.replace(char, '')
    path = str(path.replace('\\', '/'))

    return path

def setUniqFileName(path, target=''):
    fileName = path.split('/')[-1]
    dirName = path.replace(fileName, '')
    if target:
        fileName = fileName.replace('.csv', '') + '_' + target
    else :
        fileName = fileName.replace('.csv', '')
    changedName = fileName
    uniq = 1
    while os.path.exists(f'{dirName}/{changedName}.csv'):  # 동일한 파일명이 존재할 때
        changedName = f'{fileName}({uniq})' # 뒤에 숫자 추가
        uniq += 1

    return dirName, changedName

if __name__ == '__main__' :
    run()