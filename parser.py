import openpyxl
import xml.etree.ElementTree as ET

def excel_to_xml(file_name):
    wb = openpyxl.load_workbook(file_name)
    root = ET.Element("workbook")

    for wb_sheet in wb.worksheets:
        sheet_name = wb_sheet.title
        sheet_root = ET.Element("worksheet", {"name": sheet_name}) 
        sheet = wb[sheet_name]
        for row in sheet.rows:
            row_root = ET.Element("row")
            for excel_cell in row:
                cell = ET.SubElement(row_root, "cell")
                cell.text = str(excel_cell.value)
            sheet_root.append(row_root)
        root.append(sheet_root)
        
    tree = ET.ElementTree(root)
    tree.write("xml/input.xml")
        
def xml_to_excel():
    return

if __name__ == "__main__":
    file_name = "excel/sample.xlsx"
    excel_to_xml(file_name)