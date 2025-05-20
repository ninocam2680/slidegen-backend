from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from flask_cors import CORS
import requests

app = Flask(__name__)
CORS(app, origins=["https://areaprompt.com"])

# 🔐 Chiave segreta condivisa con WordPress
SHARED_SECRET = "slidegen-2024-key-Zx4r9Lp1"

# 📦 Genera presentazione dinamicamente
def create_presentation(slides_data, title=None, style=None):
    prs = Presentation()

    for slide_info in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # layout vuoto

        # ➤ Titolo
        title_text = slide_info.get("title", "")
        if title_text:
            title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(8), Inches(1))
            tf = title_box.text_frame
            p = tf.add_paragraph() if not tf.text else tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(32)
            p.font.bold = True

        # ➤ Contenuto
        content_text = slide_info.get("content", "")
        if content_text:
            content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(5.5), Inches(4))
            tf = content_box.text_frame
            tf.word_wrap = True
            p = tf.add_paragraph()
            p.text = content_text
            p.font.size = Pt(20)

        # ➤ Immagine (se presente)
        image_url = slide_info.get("image_url")
        layout = slide_info.get("layout", "").lower()

        if image_url and "solo testo" not in layout:
            try:
                img_data = requests.get(image_url, timeout=8).content
                image_stream = BytesIO(img_data)

                # posizione immagine a seconda del layout
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


# 🌐 Endpoint API
@app.route("/generate", methods=["POST"])
def generate_pptx():
    data = request.get_json()

    if not data or "slides" not in data or data.get("secret") != SHARED_SECRET:
        return jsonify({"error": "Unauthorized or invalid input"}), 403

    try:
        prs = create_presentation(
            slides_data=data["slides"],
            title=data.get("title"),
            style=data.get("style")
        )

        pptx_io = BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)

        return send_file(
            pptx_io,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="presentazione.pptx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)
