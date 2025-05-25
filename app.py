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

# Layout positions in inches for strict alignment
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

def create_presentation(slides_data, title=None, style=None, format="16:9", dimensions=None, fonts=None):
    prs = load_template(style)
    remove_default_slides(prs)

    for slide_info in slides_data:
        layout = slide_info.get("layout", "solo testo").lower()
        layout_spec = LAYOUTS.get(layout, LAYOUTS["solo testo"])

        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Title
        title_text = slide_info.get("title", "")
        if title_text:
            box = slide.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(8.5), Inches(1))
            tf = box.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(fonts.get("title", {}).get("size", 36))
            p.font.bold = True
            p.font.color.rgb = _rgb(fonts.get("title", {}).get("color"))

        # Content
        content_text = slide_info.get("content", "")
        if content_text and "text" in layout_spec:
            x, y, w, h = layout_spec["text"]
            box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
            tf = box.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = content_text
            p.font.size = Pt(fonts.get("content", {}).get("size", 22))
            p.font.color.rgb = _rgb(fonts.get("content", {}).get("color"))

        # Image
        image_url = slide_info.get("image_url")
        if image_url and "image" in layout_spec:
            try:
                img_data = requests.get(image_url, timeout=10).content
                image_stream = BytesIO(img_data)
                x, y, w, h = layout_spec["image"]
                slide.shapes.add_picture(image_stream, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
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

