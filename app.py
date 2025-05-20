from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from flask_cors import CORS
import requests

app = Flask(__name__)
CORS(app, origins=["https://areaprompt.com"])


# 🔐 Chiave segreta condivisa con WordPress
SHARED_SECRET = "slidegen-2024-key-Zx4r9Lp1"  # <-- imposta anche in wp_localize_script()

# Mappa dei layout
LAYOUT_MAP = {
    "testo centrato": "Testo centrato",
    "solo testo": "Solo testo",
    "immagine a sinistra": "Immagine a sinistra",
    "immagine a destra": "Immagine a destra"
}

# Funzione per creare la presentazione
def create_presentation(slides_data, title=None, style=None):
    prs = Presentation("template.pptx")  # Template di base

    # Slide di apertura con titolo, se presente
    if title:
        try:
            title_slide_layout = prs.slide_layouts[0]  # Titolo e sottotitolo
            slide = prs.slides.add_slide(title_slide_layout)
            if slide.shapes.title:
                slide.shapes.title.text = title
        except Exception as e:
            print(f"Errore slide iniziale: {e}")

    for slide_info in slides_data:
        layout_name = LAYOUT_MAP.get(slide_info.get("layout", "").lower(), "Testo centrato")
        layout = next((l for l in prs.slide_layouts if l.name == layout_name), prs.slide_layouts[0])
        slide = prs.slides.add_slide(layout)

        # Titolo
        title_shape = slide.shapes.title
        if title_shape:
            title_shape.text = slide_info.get("title", "")

        # Contenuto
        content_box = None
        for placeholder in slide.placeholders:
            try:
                if placeholder.placeholder_format.idx != 0:
                    content_box = placeholder
                    break
            except Exception:
                continue

        if content_box:
            content_box.text = slide_info.get("content", "")

        # Immagine
        image_url = slide_info.get("image_url")
        if image_url and "solo testo" not in layout_name.lower():
            try:
                img_data = requests.get(image_url).content
                image_stream = BytesIO(img_data)
                slide.shapes.add_picture(image_stream, Inches(5), Inches(2), width=Inches(4))
            except Exception as e:
                print(f"Errore nel caricamento immagine: {e}")

    return prs

# Endpoint API
@app.route("/generate", methods=["POST"])
def generate_pptx():
    data = request.get_json()

    # ❌ Verifica input e segreto
    if not data or "slides" not in data or data.get("secret") != SHARED_SECRET:
        return jsonify({"error": "Unauthorized or invalid input"}), 403

    try:
        # ✅ Crea la presentazione con i dati ricevuti
        prs = create_presentation(
            slides_data=data["slides"],
            title=data.get("title"),
            style=data.get("style")  # facoltativo
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

# Avvio locale (sviluppo)
if __name__ == "__main__":
    app.run(debug=True)

