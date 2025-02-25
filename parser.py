import os

import openpyxl
import xlsxwriter
import xml.etree.ElementTree as ET

def excel_to_xml(file_name):
    wb = openpyxl.load_workbook(file_name)
    root = ET.Element("workbook", {"title": file_name})

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
    output_file = "excel/output.xlsx"
    if os.path.exists(output_file):
        os.remove(output_file)
    tree = ET.parse("xml/input.xml")
    root = tree.getroot()

    workbook = xlsxwriter.Workbook(output_file)
    for sheet in root.findall("worksheet"):
        sheet_name = sheet.get("name")
        ws = workbook.add_worksheet(sheet_name)
        # parse row xml to its data
        for i, row in enumerate(sheet.findall("row")):
            for j, cell in enumerate(row.findall("cell")):
                data = cell.text
                if (data != "None"):
                    ws.write(i, j, data)

    workbook.close()
    return

if __name__ == "__main__":
    file_name = "excel/sample.xlsx"
    # excel_to_xml(file_name)
    xml_to_excel()