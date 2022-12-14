# GlossaryDownloader
## Glossary_Downloader 신규 개선사항  
1. 예외가 필요한 파일명 및 열을 받아서 처리하도록 함  
2. 그에따른 예외처리 및 문구 추가  
3. 설정 파일을 외부 파일에서 변경할 경우 인식이 되지 않던 버그 수정

## Guide
1. `GlossaryDownloader.exe` 파일을 실행시킵니다.  

2. 순서에 따라 세팅을 진행합니다.  
(세팅이 이미 되어 있는 경우 `glosarry_downloader_settings.csv`에서 확인 및 수정할 수 있습니다.)  
    2-1. csv를 직접 수정하는 경우 모든 행/열은 `1`이 아니라 `0` 부터 시작됨에 유의하세요.  
    2-2. `엑셀 파일 열 정보`는 처음 만들어질 때는 2개가 입력되도록 설계되었지만,  
         필요에 따라 4개까지 입력할 수 있습니다.  
         기본 입력(1,2번)이 아닌 3,4번은 파일 이름에 gender가 들어 있을 때 사용됩니다.  

3. 미니 용어집 (.csv) 파일의 경로를 입력합니다. (파일 이름 포함)  
    드래그해서 넣은 뒤 엔터를 눌러도 됩니다.  
    다만 cmd 창 특성 상 드래그 한 뒤에도 본인을 가리키지 않으니 창을 한 번 클릭해 주세요.  

4. 엑셀 파일의 경로 또는 엑셀 파일이 들어 있는 폴더 경로를 입력하세요.  

5. 실행 순서는 아래와 같습니다.  
    
    // 비교하여 완전히 동일한 한글이 있는 경우 타겟 언어를 가져옵니다.  
    1) 미니 용어집 : 구글 시트 비교  
    2) 1)의 결과물 중 번역이 비어있는 부분 : 엑셀 파일 비교         
    3) 2)의 결과물 중 번역이 비어있는 부분 : 엑셀 파일에서 해당 내용이 포함 된 텍스트  
