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
        raise FileNotFoundError(f"Template not found: {filename}")
    return Presentation(path)

def _rgb(hex_color):
    if not hex_color or not isinstance(hex_color, str) or len(hex_color) < 6:
        return RGBColor(0, 0, 0)
    hex_color = hex_color.lstrip("#")
    return RGBColor(
        int(hex_color[0:2], 16), 
        int(hex_color[2:4], 16), 
        int(hex_color[4:6], 16))

def remove_default_slides(prs):
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

def apply_font_style(target_font, source_font):
    """Apply all font attributes from source to target"""
    if source_font.size:
        target_font.size = source_font.size
    if source_font.bold is not None:
        target_font.bold = source_font.bold
    if source_font.italic is not None:
        target_font.italic = source_font.italic
    if source_font.underline is not None:
        target_font.underline = source_font.underline
    if source_font.color and source_font.color.rgb:
        target_font.color.rgb = source_font.color.rgb

def convert_bullets(text):
    if not text:
        return []
    lines = text.split('\n')
    items = []
    for line in lines:
        line = line.strip()
        if line.startswith('- '):
            items.append(('li', line[2:].strip()))
        elif line:
            items.append(('p', line))
    return items

def extract_template_styles(prs):
    """Extract styles from the template's master slides"""
    styles = {
        'title': None,
        'content': None,
        'subtitle': None,
        'image': None
    }
    
    # Check all layouts in the template
    for layout in prs.slide_layouts:
        try:
            # Create temporary slide to inspect placeholders
            slide = prs.slides.add_slide(layout)
            
            for shape in slide.shapes:
                if not shape.is_placeholder:
                    continue
                    
                placeholder = shape.placeholder_format
                ph_type = placeholder.type if hasattr(placeholder, 'type') else None
                
                # Determine placeholder type
                if ph_type:
                    if 'TITLE' in str(ph_type):
                        if shape.has_text_frame and shape.text_frame.text:
                            styles['title'] = shape.text_frame.paragraphs[0].font
                    elif 'BODY' in str(ph_type):
                        if shape.has_text_frame and shape.text_frame.text:
                            styles['content'] = shape.text_frame.paragraphs[0].font
                    elif 'SUBTITLE' in str(ph_type):
                        if shape.has_text_frame and shape.text_frame.text:
                            styles['subtitle'] = shape.text_frame.paragraphs[0].font
                    elif 'PICTURE' in str(ph_type):
                        styles['image'] = True
            
            prs.slides.remove(slide)
            
            # Stop if we've found all styles
            if all(styles.values()):
                break
                
        except Exception as e:
            print(f"Warning processing layout: {e}")
            if 'slide' in locals():
                prs.slides.remove(slide)
    
    return styles

def create_presentation(slides_data, title=None, style="default", format="16:9", dimensions=None, fonts=None):
    try:
        prs = load_template(style)
    except FileNotFoundError:
        prs = Presentation()
        # Set default slide size for 16:9 if new presentation
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(5.625)
    
    remove_default_slides(prs)

    # Extract styles from template
    template_styles = extract_template_styles(prs)

    for slide_info in slides_data:
        layout = slide_info.get("layout", "solo testo").lower()
        layout_spec = LAYOUTS.get(layout, LAYOUTS["solo testo"])

        # Use blank layout (usually index 6)
        slide_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)

        # 1. Add Title
        title_text = slide_info.get("title", "")
        if title_text:
            title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(8.0), Inches(1.0))
            title_frame = title_box.text_frame
            title_frame.clear()
            title_para = title_frame.paragraphs[0]
            title_para.text = title_text
            
            # Apply style from template or use defaults
            if template_styles['title']:
                apply_font_style(title_para.font, template_styles['title'])
            else:
                title_para.font.size = Pt(44)
                title_para.font.bold = True
                title_para.font.color.rgb = RGBColor(0x1A, 0x35, 0x6B)  # Dark blue

        # 2. Add Subtitle if exists in data
        subtitle_text = slide_info.get("subtitle", "")
        if subtitle_text:
            subtitle_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(7.5), Inches(0.8))
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.clear()
            subtitle_para = subtitle_frame.paragraphs[0]
            subtitle_para.text = subtitle_text
            
            if template_styles['subtitle']:
                apply_font_style(subtitle_para.font, template_styles['subtitle'])
            elif template_styles['content']:
                apply_font_style(subtitle_para.font, template_styles['content'])
            else:
                subtitle_para.font.size = Pt(28)
                subtitle_para.font.italic = True
                subtitle_para.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)  # Gray

        # 3. Add Main Content
        content_text = slide_info.get("content", "")
        if content_text and "text" in layout_spec:
            x, y, w, h = layout_spec["text"]
            content_box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
            content_frame = content_box.text_frame
            content_frame.clear()
            content_frame.word_wrap = True

            for type_, txt in convert_bullets(content_text):
                para = content_frame.add_paragraph()
                para.text = txt
                
                if template_styles['content']:
                    apply_font_style(para.font, template_styles['content'])
                else:
                    para.font.size = Pt(28)
                    para.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # Black
                
                if type_ == 'li':
                    para.level = 0
                    para.font.bold = False

        # 4. Add Image if specified in layout
        image_url = slide_info.get("image_url", "")
        if image_url and "image" in layout_spec:
            try:
                response = requests.get(image_url, timeout=8)
                response.raise_for_status()
                image_stream = BytesIO(response.content)
                x, y, w, h = layout_spec["image"]
                slide.shapes.add_picture(image_stream, Inches(x), Inches(y), Inches(w), Inches(h))
            except Exception as e:
                print(f"Image error: {e}")
                # Add placeholder
                x, y, w, h = layout_spec["image"]
                placeholder = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
                placeholder.text_frame.text = "Image not available"
                placeholder.text_frame.paragraphs[0].font.size = Pt(12)

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
            style=data.get("style", "default"),
            format=data.get("format", "16:9"),
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
