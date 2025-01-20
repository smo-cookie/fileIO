import zipfile
from xml.etree import ElementTree as ET

def read_word_as_xml(docx_file):
    # .docx 파일은 실제로 ZIP 형식이므로, zipfile을 이용하여 압축을 풀 수 있습니다.
    with zipfile.ZipFile(docx_file, 'r') as docx_zip:
        # 'word/document.xml' 경로로 XML 파일을 찾습니다.
        xml_content = docx_zip.read('word/document.xml')
        
        # XML 콘텐츠를 ElementTree로 파싱
        tree = ET.ElementTree(ET.fromstring(xml_content))
        
        # XML 트리 반환
        return tree



def read_excel_as_xml(xlsx_file):
    # .xlsx 파일은 ZIP 형식이므로 zipfile로 압축을 풉니다.
    with zipfile.ZipFile(xlsx_file, 'r') as xlsx_zip:
        # 'xl/worksheets/sheet1.xml' 경로로 XML 파일을 찾습니다.
        xml_content = xlsx_zip.read('xl/worksheets/sheet1.xml')
        
        # XML 콘텐츠를 ElementTree로 파싱
        tree = ET.ElementTree(ET.fromstring(xml_content))
        
        # XML 트리 반환
        return tree

    

# 사용 예시(word)
docx_file = 'document.docx'  # 읽을 워드 파일 경로
xml_tree = read_word_as_xml(docx_file)

# 사용 예시(엑셀)
xlsx_file = 'example.xlsx'  # 읽을 엑셀 파일 경로
xml_tree = read_excel_as_xml(xlsx_file)


# XML 트리의 루트 요소를 출력 (예시)
print(ET.tostring(xml_tree.getroot(), encoding='utf-8').decode())
