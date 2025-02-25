# inspired by https://github.com/ultr4nerd/openpyxl-image-loader/blob/master/openpyxl_image_loader/sheet_image_loader.py
import string
import io
import openpyxl

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
