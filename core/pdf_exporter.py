from __future__ import annotations

from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


class PdfExporter:
    def export_pages(
        self,
        pages: list[list[dict]],
        output_path: str,
        scene_width: int,
        scene_height: int,
        drawing_pages: list[list[dict]] | None = None,
    ) -> str:
        out = Path(output_path)
        c = canvas.Canvas(str(out), pagesize=A4)
        page_w, page_h = A4
        scale = min(page_w / max(scene_width, 1), page_h / max(scene_height, 1))

        drawing_pages = drawing_pages or [[] for _ in pages]
        for page_index, page_blocks in enumerate(pages):
            for drawing in drawing_pages[page_index] if page_index < len(drawing_pages) else []:
                self._draw_drawing(c, drawing, page_h, scale)
            for block in page_blocks:
                image_path = block.get("temp_path")
                if not image_path:
                    image_path = str(Path(output_path).with_suffix('')) + f"_page{page_index+1}_block{page_blocks.index(block)+1}.png"
                    block["temp_path"] = image_path
                if not Path(image_path).exists():
                    block["image"].save(image_path)
                x = block["x"] * scale
                y = page_h - (block["y"] + block["h"]) * scale
                w = block["w"] * scale
                h = block["h"] * scale
                c.drawImage(image_path, x, y, width=w, height=h, preserveAspectRatio=False, mask="auto")
            c.showPage()
        c.save()
        return str(out)

    def _draw_drawing(self, c: canvas.Canvas, drawing: dict, page_h: float, scale: float) -> None:
        if drawing.get('type') == 'textbox':
            self._draw_textbox(c, drawing, page_h, scale)
            return
        width = max(0.1, float(drawing.get('width', 0.5)))
        c.setLineWidth(width)
        c.setStrokeColorRGB(0, 0, 0)
        x1 = float(drawing.get('x1', 0.0)) * scale
        y1 = page_h - float(drawing.get('y1', 0.0)) * scale
        x2 = float(drawing.get('x2', 0.0)) * scale
        y2 = page_h - float(drawing.get('y2', 0.0)) * scale
        c.line(x1, y1, x2, y2)

    def _draw_textbox(self, c: canvas.Canvas, drawing: dict, page_h: float, scale: float) -> None:
        x = float(drawing.get('x', 0.0)) * scale
        y_top = float(drawing.get('y', 0.0)) * scale
        w = float(drawing.get('w', 0.0)) * scale
        h = float(drawing.get('h', 0.0)) * scale
        y = page_h - y_top - h
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0, 0, 0)
        c.rect(x, y, w, h, stroke=1, fill=0)
        font_size = max(6.0, float(drawing.get('font_size', 14)))
        c.setFont('Helvetica', font_size)
        text = str(drawing.get('text', ''))
        text_obj = c.beginText(x + 4 * scale, y + h - font_size - 4 * scale)
        for line in text.splitlines() or ['']:
            text_obj.textLine(line)
        c.drawText(text_obj)
