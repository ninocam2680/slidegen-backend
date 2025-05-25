from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from io import BytesIO
from flask_cors import CORS
import requests
import os

app = Flask(__name__)
CORS(app, origins=["https://areaprompt.com"])

SHARED_SECRET = "slidegen-2024-key-Zx4r9Lp1"
TEMPLATE_DIR = "templates"

LAYOUTS = {
    "solo testo": {
        "text": (1.0, 1.5, 8.0, 4.5)
    },
    "immagine a sinistra": {
        "image": (0.7, 1.5, 3.2, 3.8),
        "text": (4.1, 1.5, 5.2, 3.8)
    },
    "immagine a destra": {
        "image": (6.3, 1.5, 3.2, 3.8),
        "text": (0.7, 1.5, 5.2, 3.8)
    },
    "testo centrato": {
        "text": (2.0, 2.0, 6.5, 3.0)
    }
}

def load_template(style):
    filename = f"{style.lower()}.pptx"
    path = os.path.join(TEMPLATE_DIR, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Template non trovato: {filename}")
    return Presentation(path)

def _rgb(hex_color):
    if not hex_color or not isinstance(hex_color, str) or len(hex_color) < 6:
        return RGBColor(0, 0, 0)
    hex_color = hex_color.lstrip("#")
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def remove_default_slides(prs):
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

def apply_font_from_template(paragraph, ref_paragraph):
    paragraph.font.size = ref_paragraph.font.size
    paragraph.font.bold = ref_paragraph.font.bold
    paragraph.font.color.rgb = ref_paragraph.font.color.rgb

def convert_bullets(text):
    lines = text.split('\n')
    items = []
    for line in lines:
        line = line.strip()
        if line.startswith('- '):
            items.append(('li', line[2:].strip()))
        else:
            items.append(('p', line))
    return items

def create_presentation(slides_data, title=None, style=None, format="16:9", dimensions=None, fonts=None):
    prs = load_template(style)
    remove_default_slides(prs)

    ref_slide = prs.slides.add_slide(prs.slide_layouts[0])
    ref_title, ref_content = None, None
    for shape in ref_slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.idx == 0:
            ref_title = shape.text_frame.paragraphs[0]
        elif shape.is_placeholder and shape.placeholder_format.idx == 1:
            ref_content = shape.text_frame.paragraphs[0]
    prs.slides.remove(ref_slide)

    for slide_info in slides_data:
        layout = slide_info.get("layout", "solo testo").lower()
        layout_spec = LAYOUTS.get(layout, LAYOUTS["solo testo"])

        slide = prs.slides.add_slide(prs.slide_layouts[6])

        for shape in list(slide.shapes):
            if shape.is_placeholder:
                shape.element.getparent().remove(shape.element)

        title_text = slide_info.get("title", "")
        if title_text:
            title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(8.0), Inches(1.0))
            tf = title_box.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = title_text
            if ref_title:
                apply_font_from_template(p, ref_title)

        content_text = slide_info.get("content", "")
        if content_text and "text" in layout_spec:
            left, top, width, height = layout_spec["text"]
            content_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
            tf = content_box.text_frame
            tf.clear()
            tf.word_wrap = True

            for type_, txt in convert_bullets(content_text):
                para = tf.add_paragraph() if tf.text else tf.paragraphs[0]
                para.text = txt
                if type_ == 'li':
                    para.level = 0
                if ref_content:
                    apply_font_from_template(para, ref_content)

        image_url = slide_info.get("image_url")
        if image_url and "image" in layout_spec:
            try:
                img_data = requests.get(image_url, timeout=8).content
                image_stream = BytesIO(img_data)
                left, top, width, height = layout_spec["image"]
                slide.shapes.add_picture(image_stream, Inches(left), Inches(top), width=Inches(width), height=Inches(height))
            except Exception as e:
                print(f"Errore immagine: {e}")

    return prs

@app.route("/generate", methods=["POST"])
def generate_pptx():
    data = request.get_json()

    if not data or "slides" not in data or data.get("secret") != SHARED_SECRET:
        return jsonify({"error": "Unauthorized or invalid input"}), 403

    try:
        prs = create_presentation(
            slides_data=data["slides"],
            title=data.get("title"),
            style=data.get("style"),
            format=data.get("format"),
            dimensions=data.get("dimensions"),
            fonts=data.get("fonts") or {}
        )

        pptx_io = BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)

        return send_file(
            pptx_io,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"presentazione_{data.get('style','default')}.pptx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
