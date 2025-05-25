from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from io import BytesIO
from flask_cors import CORS
import requests
import os

app = Flask(__name__)
CORS(app, origins=["https://areaprompt.com"])

SHARED_SECRET = "slidegen-2024-key-Zx4r9Lp1"
TEMPLATE_DIR = "templates"

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

def create_presentation(slides_data, title=None, style=None, format="16:9", dimensions=None, fonts=None):
    prs = load_template(style)
    remove_default_slides(prs)

    for slide_info in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        title_text = slide_info.get("title", "")
        if title_text:
            title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(8), Inches(1))
            tf = title_box.text_frame
            p = tf.add_paragraph() if not tf.text else tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(fonts.get("title", {}).get("size", 32))
            p.font.bold = True
            p.font.color.rgb = _rgb(fonts.get("title", {}).get("color"))

        content_text = slide_info.get("content", "")
        if content_text:
            content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(5.5), Inches(4))
            tf = content_box.text_frame
            tf.word_wrap = True
            p = tf.add_paragraph()
            p.text = content_text
            p.font.size = Pt(fonts.get("content", {}).get("size", 20))
            p.font.color.rgb = _rgb(fonts.get("content", {}).get("color"))

        image_url = slide_info.get("image_url")
        layout = slide_info.get("layout", "").lower()

        if image_url and "solo testo" not in layout:
            try:
                img_data = requests.get(image_url, timeout=8).content
                image_stream = BytesIO(img_data)

                if "sinistra" in layout:
                    left, top = Inches(6.5), Inches(1.5)
                elif "destra" in layout:
                    left, top = Inches(0.5), Inches(1.5)
                else:
                    left, top = Inches(2), Inches(3.5)

                slide.shapes.add_picture(image_stream, left, top, width=Inches(3))
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

