# inspired by https://github.com/ultr4nerd/openpyxl-image-loader/blob/master/openpyxl_image_loader/sheet_image_loader.py
import string
import io
import os
import openpyxl
import xml.etree.ElementTree as ET

import zipfile
import shutil

def workbook_search_tag(tag):
    return ".//{*}" + tag

def worksheet_search_tag(tag):
    return "ns0:" + tag

tags = {
    "wb_rel": "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
}

class SheetImages:
    def __init__(self, sheet_name, sheet_images):
        """ 
        saves all images in a sheet locally
        naming scheme: <sheet_name>_<cell>.jpg
        cell: 
        """
        self.image_cells = {}
        for image in sheet_images:
            row = image.anchor._from.row + 1
            col = string.ascii_uppercase[image.anchor._from.col]
            cell = f"{col}{row}"
            img_data = io.BytesIO(image._data())

            file_name = f"images/{sheet_name}_{cell}.jpg"
            self.image_cells[cell] = file_name
            with open(file_name, "wb") as img_file:
                img_file.write(img_data.getvalue())

    def get_image_file_name(self, cell):
        return self.image_cells[cell]

    def image_in(self, cell):
        return cell in self.image_cells
    
class ExcelHelper:
    def __init__(self):
        pass

    def unzip_excel(self, input_file, zip_folder):
        """Unzip an Excel file (.xlsx) to a directory and get its underlying xmls"""
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(zip_folder)

    def rezip_excel(self, zip_folder, output_file):
        """Recompress the folder back into a .xlsx file"""
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, _, files in os.walk(zip_folder):
                for file in files:
                    full_path = os.path.join(root, file)
                    # Preserve the internal structure of the Excel file
                    arcname = os.path.relpath(full_path, zip_folder)
                    zip_ref.write(full_path, arcname)

    def extract_sheet_drawings(self, zip_folder, sheet_name):
        # map sheet name to worksheet xml using the workbook.xml.rels
        ws_file = self.get_sheet_name(zip_folder, sheet_name)

        # check if sheet xml has a drawing tag
        ws_path = f"{zip_folder}/xl/{ws_file}"
        ws_tree = ET.parse(ws_path)
        ws_root = ws_tree.getroot()
        ws_drawing_rId = None
        for child in ws_root:
            if child.tag.endswith("drawing"):
                tag = tags["wb_rel"]+'id'
                ws_drawing_rId = child.get(tag)
        if ws_drawing_rId == None:
            return
        
        # if so, use worksheet/_rels to get the file path of sheet drawings
        ws_name = ws_file.split("worksheets/")[-1]
        ws_rel_path = f"{zip_folder}/xl/worksheets/_rels/{ws_name}.rels"
        ws_rel_tree = ET.parse(ws_rel_path)
        ws_rel_root = ws_rel_tree.getroot()
        ws_drawing_path = ""
        for child in ws_rel_root:
            if child.get("Id") == ws_drawing_rId:
                ws_drawing_path = child.get("Target")

        # in the sheet drawings, find all xdr:oneCellAnchor or xdr:twoCellAnchor tags

        # store each anchor tags, currently in its raw xml string first

        # check if there is a drawing.xml.rels we need to parse too

        # store all media images in memory or just write to disk immediately

        return
    
    def get_sheet_name(self, zip_folder, sheet_name):
        workbook_path = f"{zip_folder}/xl/workbook.xml"
        wb_tree = ET.parse(workbook_path)
        root = wb_tree.getroot()

        rId = ""
        for sheet in root.find(workbook_search_tag("sheets")):
            if sheet.get("name") == sheet_name:
                tag = tags["wb_rel"]+'id'
                rId = sheet.get(tag)
                break
        
        wb_rels_tree = ET.parse(f"{zip_folder}/xl/_rels/workbook.xml.rels")
        root = wb_rels_tree.getroot()
        for rel in root:
            if rel.get("Id") == rId:
                path = rel.get("Target")
                return path
        
    def inject_sheet_drawings(self, sheet_name, sheet_drawings):
        # unzip excel

        # if there is at least 1 drawing, 
        #   create a <drawing> tag in the end of the sheet xml 
        #   create a .rels file in worksheets/_rels file that links to a drawings xml file in /drawings

        # init drawings file with correct xml headers & schema

        # add all sheet drawings xml to the drawing file

        return
    
