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