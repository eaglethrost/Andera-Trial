from openpyxl import load_workbook
import zipfile
import xml.etree.ElementTree as ET

# try:
#     wb = load_workbook("excel/output2.xlsx")
#     print("Excel file loaded successfully!")
# except Exception as e:
#     print("Error opening file:", e)

excel_file = "excel/output.xlsx"

with zipfile.ZipFile(excel_file, "r") as zip_ref:
    for file_name in zip_ref.namelist():
        if file_name.endswith(".xml"):
            try:
                xml_content = zip_ref.read(file_name)
                ET.fromstring(xml_content)  # Try parsing XML
            except ET.ParseError as e:
                print(f"XML Error in {file_name}: {e}")

def travel(path):
    tree = ET.parse(path)
    root = tree.getroot()
    for child in root:
        print(child.attrib)

travel("excel_xml/output_zip/xl/worksheets/sheet2.xml")