import os
import csv
import gspread
import pathlib
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import load_workbook
from os import walk
from setPath import run as setPath

def run() :
    # 세팅 불러오기
    setting = load_setting()

    # 미니 용어집 불러오기
    path = setPath(input("미니 용어집을 드래그하거나 경로(.csv)를 입력해 주세요.\n"))
    # path = 'c:\\Users\\Umoo\\Downloads\\[BG_Mir4] Mini Glossary - 0920 Update.csv'
    mGlossary = loadMGlossary(path, setting)
    mData = mGlossary[0]
    googleGlossaryPath = setting[4][0]
    gGlossary_col_info = setting[7]
    eFile_col_info = setting[10]
    target = mData[1].upper().strip()

    # 구글시트와 비교
    gData = loadGsheet(target, googleGlossaryPath)
    compared_data = gCompare(gData, mGlossary, gGlossary_col_info)

    # 엑셀과 비교
    eData = loadEsheet(eFile_col_info)
    compared_data = Ecompare(eData, compared_data)

    # 엑셀에서 단어가 들어간 내용 찾기
    compared_data = Efind(eData, compared_data)
    
    # 생성되는 파일 이름 관리
    fileName = path.split('/')[-1]
    dirName = path.replace(fileName, '')
    changedName = fileName.replace('.csv', '') + '_' + target
    uniq = 1
    while os.path.exists(dirName + '/' + changedName + '.csv'):  # 동일한 파일명이 존재할 때
        changedName = changedName + f'({uniq})'
        uniq += 1

    writeFile(compared_data, changedName, dirName)

    print("완료")
    os.system("pause")

# 
def writeData(path, compared_data) :
    with open(path, 'w', encoding="utf-8-sig", newline='') as writeFile:
        try:
            csvWriter = csv.writer(writeFile)
            csvWriter.writerows(compared_data)

        except Exception as e:
            print(e)


def loadEsheet(eFile_col_info) :
    data = []
    path = setPath(input("엑셀 파일 경로 또는 폴더의 경로를 입력해 주세요.\n"))
    # path = '''D:\\!Project\\Umoo\\umoo\\Wemade\\1_Work\\0209_8테마선번역_Regular\\ㅗㅜㅑ'''

    if '.xlsx' in os.path.splitext(path)[1].lower() :
        filename = path.split('/')[-1]
        loadedFile = load_workbook(path, data_only=True)
        sheetNames = loadedFile.sheetnames
        source_col = eFile_col_info[0]
        if 'gender' in filename.lower() and len(eFile_col_info) > 2 : source_col = eFile_col_info[2]
        target_col = eFile_col_info[1]
        if 'gender' in filename.lower() and len(eFile_col_info) > 3 : target_col = eFile_col_info[3]

        for sheetName in sheetNames:
            sheet = loadedFile[sheetName]

            for row in sheet.iter_rows(min_row=2):
                row_value = [row[int(source_col)].value, row[int(target_col)].value]
                # for i, cell in enumerate(row):
                #     if i == int(source_col) or i == int(target_col):
                #         if cell.value is None:
                #             cell.value = ''
                #         row_value.append(cell.value)
                if None not in row_value : data.append(row_value)

    else :
        targetFilesAbsolutePaths = []
        for (dirPaths, dirNames, fileNames) in walk(path):
            targetFilesAbsolutePaths.extend([dirPaths + '\\' + fileName for fileName in fileNames])

        for fileAbsolutePath in targetFilesAbsolutePaths:
            if os.path.splitext(fileAbsolutePath)[1].lower() != ".xlsx": continue
            filename = fileAbsolutePath.split('/')[-1]
            loadedFile = load_workbook(fileAbsolutePath, data_only=True)
            sheetNames = loadedFile.sheetnames
            source_col = eFile_col_info[0]
            if 'gender' in filename.lower() and len(eFile_col_info) > 2 : source_col = eFile_col_info[2]
            target_col = eFile_col_info[1]
            if 'gender' in filename.lower() and len(eFile_col_info) > 3 : target_col = eFile_col_info[3]
            for sheetName in sheetNames:
                sheet = loadedFile[sheetName]
                for row in sheet.iter_rows(min_row=2):
                    row_value = [row[int(source_col)].value, row[int(target_col)].value]
                    if None not in row_value : data.append(row_value)

    return(data)

def loadGsheet(target, googleGlossaryPath) :
    print("구글 용어집 확인 중..")
    doc = connectGsheet(googleGlossaryPath)
    worksheets = doc.worksheets()
    sheetname_list = [worksheet.title for worksheet in worksheets]
    if target in sheetname_list :
        worksheet_index = sheetname_list.index(target)
    else :
        # map + lambda 보다 list comprehesion이 더 성능이 좋음
        sheetname_list = [sheetname.upper().replace('KO2','') for sheetname in sheetname_list]
        if target in sheetname_list :
            worksheet_index = sheetname_list.index(target)
        else : print('Target 언어 용어집을 찾을 수 없습니다.')
    
    worksheet = doc.get_worksheet(worksheet_index)
    gdata = worksheet.get_all_values()

    print("구글 용어집 확인 완료..")
    return gdata

def gCompare(gGlossaryData, miniGlossaryData, gGlossary_col_info) :
    print("텀베이스 용어집과 미니 용어집 비교 중...")
    # 해당 언어 전체 구글 시트 내용
    gData = gGlossaryData
    # 미니 용어집의 한글 + 해당 언어열
    mData = miniGlossaryData
    gSelected_source_col = int(gGlossary_col_info[0])
    gSelected_target_col = int(gGlossary_col_info[1])
    data = [mData[0]]

    for mRow in mData[1:] :
        # 같은 한글이 있는지 확인
        for gRow in gData[1:] :
            row_data = [mRow[0]]
            if len(mRow) > 1 and mRow[-1] != '':
                row_data.extend([mRow[1], '√'])
                break

            elif (mRow[0].replace(' ', '') == gRow[gSelected_source_col].replace(' ', '')) and gRow[gSelected_target_col] != '' :
                row_data.extend([gRow[gSelected_target_col], '√'])
                break

        data.append(row_data)
    print("비교 완료")
    return data

def Ecompare(eGlossaryData, miniGlossaryData) :
    eData = eGlossaryData
    mData = miniGlossaryData
    data = [mData[0]]

    for mRow in mData[1:] :
        # 같은 한글이 있는지 확인
        row_data = [mRow[0]]
        for eRow in eData :
            if len(mRow) > 1 and mRow[-1] != '':
                row_data.extend([mRow[1], '√'])
                break

            elif mRow[0].replace(' ', '') == eRow[0].replace(' ', '') and eRow[1] != '' :
                row_data.extend([eRow[1], '√'])
                break

        data.append(row_data)

    return data

def Efind(eGlossaryData, miniGlossaryData) :
    eData = eGlossaryData
    mData = miniGlossaryData
    data = [mData[0]]

    for mRow in mData[1:] :
        # 같은 한글이 있는지 확인
        row_data = [mRow[0]]
        for eRow in eData :
            if len(mRow) > 1 and mRow[-1] != '':
                row_data.extend([mRow[1], '√'])
                break

            elif mRow[0].replace(' ', '') in eRow[0].replace(' ', '') and eRow[1] != '' :
                row_data.append(eRow[1])
                break

        data.append(row_data)

    return data

def connectGsheet(googleGlossaryPath) :
    dirPath = str(pathlib.Path.cwd())

    # json_file_name = dirPath + '/majestic-layout-275109-d180b5dbabbe.json'
    json_file_name = dirPath + '\\majestic-layout-275109-d180b5dbabbe.json'
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
    try :
        credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
        gc = gspread.authorize(credentials)
        doc = gc.open_by_url(googleGlossaryPath)
    except Exception as e :
        print(e)
        print("권한이 없습니다.")
        os.system("pause")

    return doc

def loadMGlossary(path, setting) :
    data = []
    # 미니 용어집 열 정보 받아오기
    selected_mrow = setting[1]
    if '.csv' not in path :
        print('잘못된 파일입니다.')
        path = ''
    else :
        path = path.replace('"','').replace('& ','').replace("'",'')
        
        with open(path, 'r', encoding='utf-8-sig') as csvFile:
            csvReader = list(csv.reader(csvFile))

        print("미니 용어집 확인 중..")
        data = []
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8-sig') as csvFile:
                csvReader = list(csv.reader(csvFile))

        # 한글, 타겟 언어만 가져오기
        for row in csvReader :
            data.append([row[int(selected_mrow[0])].strip(), row[int(selected_mrow[1])]])

        print("미니 용어집 확인 완료")
        return data

def load_setting() :
    dirPath = str(pathlib.Path.cwd())
    setting = []
    print("setting 파일 확인 중..")
    if os.path.exists(f'{dirPath}/glosarry_downloader_settings.csv'):
        with open(f'{dirPath}/glosarry_downloader_settings.csv', 'r', encoding='utf-8-sig') as csvFile:
            csvReader = list(csv.reader(csvFile))
            # 제대로 된 파일이면 setting 데이터를 반환
            if all(
                [bool('미니 용어집 열 정보' in csvReader[0]),
                bool('텀베이스 용어집 주소' in csvReader[3]),
                bool('텀베이스 용어집 열 정보' in csvReader[6]),
                bool(len(csvReader[1]) == 2),
                bool(len(csvReader[4]) == 1),
                bool(len(csvReader[7]) == 2)]) :
                print("setting 파일을 확인했습니다.")
                for row in csvReader :
                    setting.append(row)

    else :
        print("올바르지 않은 파일 형식입니다.\n설정을 진행합니다..")
        setting = setSetting()
        writeFile(setting, 'glosarry_downloader_settings')

    return setting

def selectRow() :
    source_row = input('소스 열 (a~z) : ')
    source_row = ord(source_row.upper().replace(' ','')) - 65

    taget_row = input('타겟 열 (a~z) : ')
    taget_row = ord(taget_row.upper().replace(' ', '')) - 65

    return [source_row, taget_row]


def setSetting() :
    setting = []
    
    # 미니 용어집 열 정보 설정
    setting.append(['미니 용어집 열 정보'])
    print('미니 용어집 열 설정')
    selcted_mrow = selectRow()
    setting.append(selcted_mrow)
    setting.append('')

    # 텀베이스 용어집 주소 설정
    setting.append(['텀베이스 용어집 주소'])
    setting.append([input('텀베이스 용어집 주소 입력\n')])
    setting.append('')

    # 텀베이스 용어집 열 정보 설정
    setting.append(['텀베이스 용어집 열 정보'])
    print('텀베이스 용어집 열 설정')
    selcted_grow = selectRow()
    setting.append(selcted_grow)

    # 엑셀 파일 열 정보 설정
    setting.append(['엑셀 파일 열 정보'])
    print('엑셀 파일 열 설정')
    selcted_grow = selectRow()
    setting.append(selcted_grow)

    return setting


def writeFile(data, filename) :
    dirPath = str(pathlib.Path.cwd()) + f'/{filename}.csv'
    with open(dirPath, 'w', encoding='utf-8-sig', newline='') as writeFile:
        try:
            csvWriter = csv.writer(writeFile)
            csvWriter.writerows(data)
        except Exception as e:
            print(e)

def writeFile(data, filename, dirpath) :
    dirPath = str(dirpath) + f'/{filename}.csv'
    with open(dirPath, 'w', encoding='utf-8-sig', newline='') as writeFile:
        try:
            csvWriter = csv.writer(writeFile)
            csvWriter.writerows(data)
        except Exception as e:
            print(e)

if __name__ == '__main__' :
    run()