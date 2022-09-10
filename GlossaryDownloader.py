import os
import csv
import gspread
import pathlib
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import load_workbook
from os import walk

# 구글 용어집 지정
# 미니 용어집 지정
# 구글 용어집 & 미니 용어집 비교
# 한글이 있으면 해당 열 번역 넣기

def run() :
    # 세팅 불러오기
    setting = load_setting()

    # 미니 용어집 불러오기
    path = input("미니 용어집을 드래그하거나 경로(.csv)를 입력해 주세요.\n")
    path = path.replace('"', '')
    mGlossary = loadMGlossary(path, setting)
    mData = mGlossary[0]
    googleGlossaryPath = setting[4][0]
    gGlossary_col_info = setting[7]
    target = mData[1].upper().strip()

    # 구글시트와 비교
    gData = loadGsheet(target, googleGlossaryPath)
    compared_data = gCompare(gData, mGlossary, gGlossary_col_info)

    # 엑셀과 비교
    eData = loadEsheet()
    compared_data = Ecompare(eData, mGlossary)
    
    # 생성되는 파일 이름 관리
    fileName = os.path.splitext(path.split('\\')[-1])[0]
    changedPath = path.replace(fileName,fileName + f'_{target}')
    uniq = 1
    while os.path.exists(changedPath):  # 동일한 파일명이 존재할 때
        changedPath = path.replace(fileName,fileName + f'_{target}({uniq})')
        uniq += 1

    writeFile(compared_data, changedPath)

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


def loadEsheet() :
    data = []
    path = input("엑셀 파일 경로 또는 폴더의 경로를 입력해 주세요.\n")
    # path = '''D:\\!Project\\Umoo\\umoo\\Wemade\\1_Work\\0209_8테마선번역_Regular\\ㅗㅜㅑ'''
    path = path.replace('"','').replace('& ','').replace("'",'')
    if '.xlsx' in os.path.splitext(path)[1].lower() :
        loadedFile = load_workbook(path, data_only=True)
        sheetNames = loadedFile.sheetnames
        selected_row = selectRow()
        for sheetName in sheetNames:
            sheet = loadedFile[sheetName]
            for row in sheet.iter_rows(min_row=1):
                row_value = []
                for i, cell in enumerate(row):
                    if i == int(selected_row[0]) or i == int(selected_row[1]):
                        if cell.value is None:
                            cell.value = ''
                        row_value.append(cell.value)
                data.append(row_value)

    else :
        targetFilesAbsolutePaths = []
        for (dirPaths, dirNames, fileNames) in walk(path):
            targetFilesAbsolutePaths.extend([dirPaths + '\\' + fileName for fileName in fileNames])

        for fileAbsolutePath in targetFilesAbsolutePaths:
            if os.path.splitext(fileAbsolutePath)[1].lower() != ".xlsx": continue
            loadedFile = load_workbook(fileAbsolutePath, data_only=True)
            sheetNames = loadedFile.sheetnames
            for sheetName in sheetNames:
                sheet = loadedFile[sheetName]
                for row in sheet.iter_rows(min_row=1):
                    row_value = []
                    for i, cell in enumerate(row):
                        if i == int(selected_row[0]) or i == int(selected_row[1]) :
                            if cell.value is None :
                                cell.value = ''
                            row_value.append(cell.value)
                    data.append(row_value)

    return(data)

def loadGsheet(target, googleGlossaryPath) :
    doc = connectGsheet(googleGlossaryPath)
    print("구글 용어집 확인 중..")
    worksheets = doc.worksheets()
    sheetname_list = [worksheet.title for worksheet in worksheets]
    if target in sheetname_list :
        worksheet_index = sheetname_list.index(target)
    else :
        sheetname_list = [sheetname.upper().replace('KO2','') for sheetname in sheetname_list]
        if target in sheetname_list :
            worksheet_index = sheetname_list.index(target)
        else : print('Target 언어 용어집을 찾을 수 없습니다.')
    
    worksheet = doc.get_worksheet(worksheet_index)
    gdata = worksheet.get_all_values()

    return gdata

def gCompare(gGlossaryData, miniGlossaryData, gGlossary_col_info) :
    gData = gGlossaryData
    mData = miniGlossaryData
    gSelected_source_col = int(gGlossary_col_info[0])
    gSelected_target_col = int(gGlossary_col_info[1])
    data = []

    for i, mRow in enumerate(mData[1:]) :
        # 같은 한글이 있는지 확인
        flag = True
        for gRow in gData[1:] :
            if (mRow[0].replace(' ', '') == gRow[gSelected_source_col].replace(' ', '')) and gRow[gSelected_target_col] != '' :
                if len(mRow) < 2 :
                    mRow.append(gRow[gSelected_target_col])
                elif len(mRow) == 2 and mRow[1] == '' :
                    mRow[1] = gRow[gSelected_target_col]
        data.append(mRow)

    return mData

def Ecompare(eGlossaryData, miniGlossaryData) :
    eData = eGlossaryData
    mData = miniGlossaryData

    for i, mRow in enumerate(mData[1:]) :
        # 같은 한글이 있는지 확인
        flag = True
        for gRow in eData[1:] :
            try :
                if mRow[0].replace(' ', '') == gRow[0].replace(' ', '') and gRow[1] != '' :
                    if len(mRow) < 2 :
                        mData[i + 1].append(gRow[1])
                    elif len(mRow) == 2 :
                        mData[i+1][1] = gRow[1]
            except Exception as e :
                print(e)

    return mData

def connectGsheet(googleGlossaryPath) :
    dirPath = str(pathlib.Path.cwd())

    json_file_name = dirPath + '/majestic-layout-275109-d180b5dbabbe.json'
    json_file_name = dirPath + '\\majestic-layout-275109-d180b5dbabbe.json'
    # json_file_name = dirPath + '\\majestic-layout-275109-54bcf57ed64c.json'
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
            data.append([row[int(selected_mrow[0])], row[int(selected_mrow[1])]])

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

    return setting


def writeFile(data, filename) :
    dirPath = str(pathlib.Path.cwd()) + f'/{filename}.csv'
    with open(dirPath, 'w', encoding='utf-8-sig', newline='') as writeFile:
        try:
            csvWriter = csv.writer(writeFile)
            csvWriter.writerows(data)
        except Exception as e:
            print(e)

if __name__ == '__main__' :
    run()