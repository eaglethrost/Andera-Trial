import openpyxl
import xml.etree.ElementTree as ET

def excel_to_xml(file_name):
    wb = openpyxl.load_workbook(file_name)
    sheet_name = "1 User Report"
    sheet = wb[sheet_name]
    root = ET.Element("root")
    for row in sheet.rows:
        for cell in row:
            ET.SubElement(root, "cell", value=str(cell.value))
    tree = ET.ElementTree(root)
    tree.write("xml/input.xml")
        
def xml_to_excel():
    return

if __name__ == "__main__":
    file_name = "excel/sample.xlsx"
    excel_to_xml(file_name)