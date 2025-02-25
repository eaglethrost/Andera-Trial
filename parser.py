import os

import openpyxl
from openpyxl_image_loader import SheetImageLoader
import xlsxwriter
import xml.etree.ElementTree as ET

from helpers import SheetImages

def excel_to_xml(file_name):
    wb = openpyxl.load_workbook(file_name)
    root = ET.Element("workbook", {"title": file_name})

    for ws in wb.worksheets:
        sheet_name = ws.title
        sheet_root = ET.Element("worksheet", {"name": sheet_name}) 
        sheet = wb[sheet_name]
        image_loader = SheetImages(sheet_name, sheet._images)

        for i, row in enumerate(sheet.rows):
            row_root = ET.Element("row")
            for cell in row:
                cell_root = ET.SubElement(row_root, "cell")
                cell_root.text = str(cell.value)
                cell_coor = cell.coordinate
                # mark if cell contains image
                if image_loader.image_in(cell_coor): 
                    cell_root.set("image", image_loader.get_image_file_name(cell_coor))

            # format row height
            if ws.row_dimensions[i+1].height:
                row_root.set("height", str(ws.row_dimensions[i+1].height))
                
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
            # format row height
            if "height" in row.attrib:
                row_h = float(row.get("height"))
                ws.set_row(i, row_h)
                
            for j, cell in enumerate(row.findall("cell")):
                data = cell.text
                if data != "None":
                    ws.write(i, j, data)
                # write image if it exists in the cell
                if "image" in cell.attrib:
                    ws.insert_image(i, j, cell.get("image"))
                    
    workbook.close()
    return

if __name__ == "__main__":
    file_name = "excel/sample.xlsx"
    excel_to_xml(file_name)
    xml_to_excel()