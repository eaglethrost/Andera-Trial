import os

import openpyxl
from openpyxl_image_loader import SheetImageLoader
import xlsxwriter
import xml.etree.ElementTree as ET

from helpers import SheetImages, ExcelHelper

class ExcelParser:
    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file
        self.input_zip = "excel_xml/sample_zip"
        self.output_zip = "excel_xml/output_zip"
        self.excel_helper = ExcelHelper()
    
    def excel_to_xml(self):
        file_name = self.input_file
        wb = openpyxl.load_workbook(file_name)
        root = ET.Element("workbook", {"title": file_name})

        # unzip excel to process its xmls
        self.excel_helper.unzip_excel(file_name, self.input_zip)

        for ws in wb.worksheets:
            sheet_name = ws.title
            sheet_root = ET.Element("worksheet", {"name": sheet_name}) 
            sheet = wb[sheet_name]
            image_loader = SheetImages(sheet_name, sheet._images)

            # extract sheet drawings
            self.excel_helper.extract_sheet_drawings(self.input_zip, sheet_name)

            # store column metadata info
            column_metadata = ET.SubElement(sheet_root, "column_metadata")
            for col, dim in ws.column_dimensions.items():
                col_meta_root = ET.SubElement(column_metadata, col)
                col_meta_root.set("width", str(dim.width))

            for i, row in enumerate(sheet.rows):
                row_root = ET.Element("row")
                for j, cell in enumerate(row):
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
        
    def xml_to_excel(self):
        output_file = self.output_file
        if os.path.exists(output_file):
            os.remove(output_file)
        tree = ET.parse("xml/input.xml")
        root = tree.getroot()

        workbook = xlsxwriter.Workbook(output_file)

        for sheet in root.findall("worksheet"):
            sheet_name = sheet.get("name")
            ws = workbook.add_worksheet(sheet_name)

            # apply column metadata
            for col in sheet.find("column_metadata"):
                ws_col = f"{col.tag}:{col.tag}"
                ws.set_column(ws_col, float(col.get("width")))

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
    input_file = "excel/sample.xlsx"
    output_file = "excel/output.xlsx"

    excel_parser = ExcelParser(input_file, output_file)

    excel_parser.excel_to_xml()
    # excel_parser.xml_to_excel()