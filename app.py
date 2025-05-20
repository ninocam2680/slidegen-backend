from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
import requests

app = Flask(__name__)

LAYOUT_MAP = {
    "testo centrato": "Testo centrato",
    "solo testo": "Solo testo",
    "immagine a sinistra": "Immagine a sinistra",
    "immagine a destra": "Immagine a destra"
}

def create_presentation(slides_data):
    prs = Presentation("template.pptx")  # Template base

    for slide_info in slides_data:
        layout_name = LAYOUT_MAP.get(slide_info.get("layout", "").lower(), "Testo centrato")
        layout = next((l for l in prs.slide_layouts if l.name == layout_name), prs.slide_layouts[0])
        slide = prs.slides.add_slide(layout)

        title = slide.shapes.title
        content_box = slide.placeholders[1] if len(slide.placeholders) > 1 else None

        if title:
            title.text = slide_info.get("title", "")

        if content_box:
            content_box.text = slide_info.get("content", "")

        # Aggiunge immagine se presente e se layout lo supporta
        image_url = slide_info.get("image_url")
        if image_url and "solo testo" not in layout_name.lower():
            try:
                img_data = requests.get(image_url).content
                image_stream = BytesIO(img_data)
                slide.shapes.add_picture(image_stream, Inches(5), Inches(2), width=Inches(4))
            except:
                pass  # Silenzia errori immagine

    return prs

@app.route("/generate-pptx", methods=["POST"])
def generate_pptx():
    data = request.get_json()
    if not data or "slides" not in data:
        return jsonify({"error": "Invalid input"}), 400

    prs = create_presentation(data["slides"])
    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)

    return send_file(pptx_io, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                     as_attachment=True, download_name="presentazione.pptx")

if __name__ == "__main__":
    app.run(debug=True)
