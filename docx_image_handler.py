import os
import re
import uuid
import zipfile
import tempfile
import shutil
from typing import Dict, List, Tuple
from lxml import etree
import base64


class DOCXImageHandler:
    """Handles image insertion into DOCX files"""

    def __init__(self):
        self.namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "rels": "http://schemas.openxmlformats.org/package/2006/relationships",
        }

    def add_images_to_docx(self, docx_bytes: bytes, images: List[Dict]) -> bytes:
        """
        Add images to a DOCX file at marker positions

        Args:
            docx_bytes: The DOCX file as bytes
            images: List of image dictionaries with marker, data, format, width, height, description

        Returns:
            Modified DOCX file as bytes
        """
        if not images:
            return docx_bytes

        # Create temp directory
        temp_dir = tempfile.mkdtemp()

        try:
            # Save input to temp file and extract
            temp_input = os.path.join(temp_dir, "input.docx")
            with open(temp_input, "wb") as f:
                f.write(docx_bytes)

            # Extract DOCX
            extract_dir = os.path.join(temp_dir, "extracted")
            with zipfile.ZipFile(temp_input, "r") as zip_ref:
                zip_ref.extractall(extract_dir)

            # Process images
            self._process_images(extract_dir, images)

            # Repackage DOCX
            output_path = os.path.join(temp_dir, "output.docx")
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zip_out:
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_path = os.path.relpath(file_path, extract_dir)
                        arc_path = arc_path.replace(os.sep, "/")
                        zip_out.write(file_path, arc_path)

            # Read result
            with open(output_path, "rb") as f:
                result = f.read()

            return result

        finally:
            # Clean up
            shutil.rmtree(temp_dir)

    def _process_images(self, extract_dir: str, images: List[Dict]):
        """Process all images and insert them into the document"""
        # Load existing relationships to get next ID
        rels_path = os.path.join(extract_dir, "word", "_rels", "document.xml.rels")
        tree = etree.parse(rels_path)
        root = tree.getroot()

        # Find highest relationship ID
        max_id = 0
        ns = self.namespaces["rels"]
        for rel in root.findall(f".//{{{ns}}}Relationship"):
            rel_id = rel.get("Id", "")
            if rel_id.startswith("rId"):
                try:
                    num = int(rel_id[3:])
                    max_id = max(max_id, num)
                except ValueError:
                    pass

        rel_counter = max_id + 1

        # Find highest image number
        max_img = 0
        media_dir = os.path.join(extract_dir, "word", "media")
        if os.path.exists(media_dir):
            for filename in os.listdir(media_dir):
                if filename.startswith("image"):
                    try:
                        num = int(re.search(r"image(\d+)", filename).group(1))
                        max_img = max(max_img, num)
                    except:
                        pass

        img_counter = max_img + 1

        # Process each image
        doc_path = os.path.join(extract_dir, "word", "document.xml")
        doc_tree = etree.parse(doc_path)
        doc_root = doc_tree.getroot()

        for image_data in images:
            marker = image_data.get("marker", "")
            if not marker.startswith("{{IMAGE:"):
                marker = f"{{{{IMAGE:{marker}}}}}"

            # Decode image data
            img_bytes = image_data.get("data", "")
            if isinstance(img_bytes, str):
                img_bytes = base64.b64decode(img_bytes)

            # Save image file
            os.makedirs(media_dir, exist_ok=True)
            img_format = image_data.get("format", "png")
            img_filename = f"image{img_counter}.{img_format}"
            img_path = os.path.join(media_dir, img_filename)
            with open(img_path, "wb") as f:
                f.write(img_bytes)

            # Add relationship
            rel_id = f"rId{rel_counter}"
            rel_elem = etree.SubElement(root, f"{{{ns}}}Relationship")
            rel_elem.set("Id", rel_id)
            rel_elem.set(
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            )
            rel_elem.set("Target", f"media/{img_filename}")

            # Insert image at marker
            self._insert_image_at_marker(
                doc_root,
                marker,
                rel_id,
                image_data.get("width", 400),
                image_data.get("height", 300),
                image_data.get("description", "Image"),
            )

            rel_counter += 1
            img_counter += 1

        # Save modified files
        tree.write(rels_path, encoding="UTF-8", xml_declaration=True, pretty_print=True)
        doc_tree.write(
            doc_path, encoding="UTF-8", xml_declaration=True, pretty_print=True
        )

    def _insert_image_at_marker(
        self,
        doc_root,
        marker: str,
        rel_id: str,
        width: int,
        height: int,
        description: str,
    ):
        """Insert image XML at marker location"""
        ns = self.namespaces

        # Find text nodes containing the marker
        for text_elem in doc_root.iter(f'{{{ns["w"]}}}t'):
            if text_elem.text and marker in text_elem.text:
                # Find the run and paragraph
                run = text_elem.getparent()
                while run is not None and run.tag != f'{{{ns["w"]}}}r':
                    run = run.getparent()

                if run is not None:
                    para = run.getparent()
                    while para is not None and para.tag != f'{{{ns["w"]}}}p':
                        para = para.getparent()

                    if para is not None:
                        # Create image XML
                        image_xml = self._create_image_xml(
                            rel_id, width, height, description
                        )

                        # Create new run for image
                        new_run = etree.Element(f'{{{ns["w"]}}}r')
                        new_run.append(image_xml)

                        # Insert after current run
                        run_index = list(para).index(run)
                        para.insert(run_index + 1, new_run)

                        # Remove marker from text
                        text_elem.text = text_elem.text.replace(marker, "")

    def _create_image_xml(self, rel_id: str, width: int, height: int, description: str):
        """Create the image XML structure"""
        ns = self.namespaces

        # Convert pixels to EMUs
        width_emu = width * 9525
        height_emu = height * 9525

        # Build the drawing XML structure
        drawing = etree.Element(f'{{{ns["w"]}}}drawing')

        inline = etree.SubElement(drawing, f'{{{ns["wp"]}}}inline')
        inline.set("distT", "0")
        inline.set("distB", "0")
        inline.set("distL", "0")
        inline.set("distR", "0")

        extent = etree.SubElement(inline, f'{{{ns["wp"]}}}extent')
        extent.set("cx", str(width_emu))
        extent.set("cy", str(height_emu))

        effect_extent = etree.SubElement(inline, f'{{{ns["wp"]}}}effectExtent')
        effect_extent.set("l", "0")
        effect_extent.set("t", "0")
        effect_extent.set("r", "0")
        effect_extent.set("b", "0")

        doc_pr = etree.SubElement(inline, f'{{{ns["wp"]}}}docPr')
        doc_pr.set("id", str(uuid.uuid4().int % 100000))
        doc_pr.set("name", description)
        doc_pr.set("descr", description)

        graphic = etree.SubElement(inline, f'{{{ns["a"]}}}graphic')
        graphic_data = etree.SubElement(graphic, f'{{{ns["a"]}}}graphicData')
        graphic_data.set(
            "uri", "http://schemas.openxmlformats.org/drawingml/2006/picture"
        )

        pic = etree.SubElement(graphic_data, f'{{{ns["pic"]}}}pic')

        nv_pic_pr = etree.SubElement(pic, f'{{{ns["pic"]}}}nvPicPr')
        c_nv_pr = etree.SubElement(nv_pic_pr, f'{{{ns["pic"]}}}cNvPr')
        c_nv_pr.set("id", "0")
        c_nv_pr.set("name", description)
        c_nv_pic_pr = etree.SubElement(nv_pic_pr, f'{{{ns["pic"]}}}cNvPicPr')

        blip_fill = etree.SubElement(pic, f'{{{ns["pic"]}}}blipFill')
        blip = etree.SubElement(blip_fill, f'{{{ns["a"]}}}blip')
        blip.set(f'{{{ns["r"]}}}embed', rel_id)

        stretch = etree.SubElement(blip_fill, f'{{{ns["a"]}}}stretch')
        fill_rect = etree.SubElement(stretch, f'{{{ns["a"]}}}fillRect')

        sp_pr = etree.SubElement(pic, f'{{{ns["pic"]}}}spPr')
        xfrm = etree.SubElement(sp_pr, f'{{{ns["a"]}}}xfrm')
        off = etree.SubElement(xfrm, f'{{{ns["a"]}}}off')
        off.set("x", "0")
        off.set("y", "0")
        ext = etree.SubElement(xfrm, f'{{{ns["a"]}}}ext')
        ext.set("cx", str(width_emu))
        ext.set("cy", str(height_emu))

        prst_geom = etree.SubElement(sp_pr, f'{{{ns["a"]}}}prstGeom')
        prst_geom.set("prst", "rect")
        av_lst = etree.SubElement(prst_geom, f'{{{ns["a"]}}}avLst')

        return drawing
