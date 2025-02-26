# inspired by https://github.com/ultr4nerd/openpyxl-image-loader/blob/master/openpyxl_image_loader/sheet_image_loader.py
import string
import io
import os
import openpyxl
import xml.etree.ElementTree as ET
from PIL import Image

import zipfile
import shutil

def workbook_search_tag(tag):
    return ".//{*}" + tag

def worksheet_search_tag(tag):
    return "ns0:" + tag

tags = {
    "wb_rel": "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
}

def process_anchor(anchor):
    s = str(anchor)
    s = s.replace("ns0", "xdr")
    s = s.replace("ns1", "a")
    s = s.replace("ns2", "a16")
    s = s[2:]
    s = s[:-1]
    return s

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

        ws_rels = {}
        ws_rels[ws_name] = {}
        for child in ws_rel_root:
            if child.get("Id") == ws_drawing_rId:
                ws_drawing_path = child.get("Target")
            ws_rels[ws_name][child.get("Id")] = {
                "Type": child.get("Type"),
                "Target": child.get("Target")
            }

        # in the sheet drawings, find all xdr:oneCellAnchor or xdr:twoCellAnchor tags
        ws_drawing_path = ws_drawing_path[3:]
        ws_drawing_path = f"{zip_folder}/xl/{ws_drawing_path}"
        ws_drawing_tree = ET.parse(ws_drawing_path)
        ws_drawing_root = ws_drawing_tree.getroot()

        anchor_tags = []
        for child in ws_drawing_root:
            if child.tag.endswith("oneCellAnchor") or child.tag.endswith("twoCellAnchor"):
                anchor_tag_xml = ET.tostring(child)
                anchor_tags.append(process_anchor(anchor_tag_xml))

        # store each anchor tags, currently in its raw xml string first
        sheet_drawings = {
            sheet_name: anchor_tags
        }
        
        # check if there is a drawing.xml.rels we need to parse too
        ws_drawing_file_name = ws_drawing_path.split("/")[-1]
        ws_drawing_rels_path = f"{zip_folder}/xl/drawings/_rels/{ws_drawing_file_name}.rels"

        # store all media images in memory or just write to disk immediately
        media_images = {}
        if os.path.exists(ws_drawing_rels_path):
            ws_draw_rels_tree = ET.parse(ws_drawing_rels_path)
            ws_draw_rels_root = ws_draw_rels_tree.getroot()
            for child in ws_draw_rels_root:
                img_path = child.get("Target")
                img_path = img_path[3:]
                file_name = img_path.split("/")[-1]
                img_path = f"{zip_folder}/xl/{img_path}"
                img = Image.open(img_path)
                img_data = img.tobytes()
                media_images[file_name] = img_data
                
        return {
            "drawings": sheet_drawings,
            "images": media_images,
            "rels": ws_rels
        }
    
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
        
    def inject_sheet_drawings(self, sheet_drawings, output_zip_folder):
        if len(sheet_drawings["drawings"]) > 0:
            self.create_drawing_folders(output_zip_folder)

        # add all sheet drawings xml to the drawing file
        self.add_sheet_drawings(sheet_drawings["drawings"], output_zip_folder)

        # add relations
        self.add_sheet_rels(sheet_drawings, output_zip_folder)
        
        return
    
    def create_drawing_folders(self, output_zip_folder):
        os.makedirs(f"{output_zip_folder}/xl/worksheets/_rels", exist_ok=True)
        os.makedirs(f"{output_zip_folder}/xl/drawings", exist_ok=True)
        os.makedirs(f"{output_zip_folder}/xl/drawings/_rels", exist_ok=True)
        return

    def add_sheet_drawings(self, drawings_data, output_zip_folder):
        for i, (sheet_name, drawings) in enumerate(drawings_data.items()):
            # create drawing.xml file in drawings/
            with open(f"{output_zip_folder}/xl/drawings/drawing{i+1}.xml", "w") as file:
                # create template
                header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                file.write(header)
                # add the string xmls to the template child
                for drawing in drawings:
                    file.write(drawing)
                closing = '</xdr:wsDr>'
                file.write(closing)
        return
    
    def add_sheet_rels(self, sheet_drawings, output_zip_folder):
        for sheet_name, relationships in sheet_drawings["rels"].items():
            # add drawing tag in the end of the sheet.xml
            ws_path = f"{output_zip_folder}/xl/worksheets/{sheet_name}"
            # if sheet_name == "sheet2.xml":
            #     self.add_drawing_tag(ws_path)

            # create sheet.xml.rels file
            ws_rels_path = f"{output_zip_folder}/xl/worksheets/_rels/{sheet_name}.rels"
            with open(ws_rels_path, "w") as file:
                # put rel inside template
                root = ET.Element("Relationships", {
                    "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
                })
                # add relationship
                for rel_id, attributes in relationships.items():
                    ET.SubElement(root, "Relationship", {
                        "Id": rel_id,
                        "Type": attributes["Type"],
                        "Target": attributes["Target"]
                    })
                # Create the XML tree and write to file
                tree = ET.ElementTree(root)
                tree.write(ws_rels_path, encoding="UTF-8", xml_declaration=True)

        return
    
    def add_drawing_tag(self, ws_path):
        ET.register_namespace('', "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        # Parse the XML file
        tree = ET.parse(ws_path)
        root = tree.getroot()

        # Define namespace mapping to strip it from tags
        namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ns = {"main": namespace}

        # Remove namespace prefix from tags for easier searching
        for elem in root.iter():
            if elem.tag.startswith(f"{{{namespace}}}"):
                elem.tag = elem.tag[len(f"{{{namespace}}}"):]  # Strip namespace

        # Find <sheetData> to place <drawing> after it
        sheet_data = root.find("sheetData")

        # Create the <drawing> tag
        drawing_tag = ET.Element("drawing", {"r:id": "rId1"})
        root.append(drawing_tag)

        # Save back to the XML file without adding namespace prefixes
        tree.write(ws_path, encoding="UTF-8", xml_declaration=True)

