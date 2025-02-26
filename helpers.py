# inspired by https://github.com/ultr4nerd/openpyxl-image-loader/blob/master/openpyxl_image_loader/sheet_image_loader.py
import string
import io
import os
import openpyxl

import zipfile
import shutil

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

    def unzip_excel(file_path, extract_to):
        """Unzip an Excel file (.xlsx) to a directory and get its underlying xmls"""
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)

    def rezip_excel(folder_path, output_file):
        """Recompress the folder back into a .xlsx file"""
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    full_path = os.path.join(root, file)
                    # Preserve the internal structure of the Excel file
                    arcname = os.path.relpath(full_path, folder_path)
                    zip_ref.write(full_path, arcname)

    def extract_sheet_drawings(unzip_folder, sheet_name):
        # map sheet name to worksheet xml using the workbook.xml.rels

        # check if sheet xml has a drawing tag. if so, use worksheet/_rels to get the file path of sheet drawings

        # in the sheet drawings, find all xdr:oneCellAnchor or xdr:twoCellAnchor tags

        # store each anchor tags, currently in its raw xml string first

        # check if there is a drawing.xml.rels we need to parse too

        # store all media images in memory or just write to disk immediately

        return

    def inject_sheet_drawings(sheet_name, sheet_drawings):
        # unzip excel

        # if there is at least 1 drawing, 
        #   create a <drawing> tag in the end of the sheet xml 
        #   create a .rels file in worksheets/_rels file that links to a drawings xml file in /drawings

        # init drawings file with correct xml headers & schema

        # add all sheet drawings xml to the drawing file

        return