import io
import os
import tempfile
import requests
from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image, ImageDraw, ImageFilter, ImageFont

# ---------- CONFIG ----------
# Path to your uploaded PPTX (developer-provided asset). We'll use it as optional template / source for logo.
TEMPLATE_PPTX_PATH = "/mnt/data/ACCINZIA (1).pptx"

# Font to use for title / body. Install on host for best results.
FONT_NAME = os.environ.get("FONT_NAME", "Calibri")  # change to "Inter", "Montserrat", etc. if installed
TITLE_FONT_SIZE = 36
BODY_FONT_SIZE = 14

# Default theme mapping (simple)
THEME_MAP = {
    "tech": ("#1F4E79", "#6FA8DC"),
    "food": ("#2E7D32", "#A5D6A7"),
    "finance": ("#0B1A34", "#6B8EA3"),
    "education": ("#6A1B9A", "#C39BD3"),
    "health": ("#EF6C00", "#FFCC80"),
    "default": ("#3E3E3E", "#BDBDBD"),
}

app = Flask(__name__)

# ---------- HELPERS ----------
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def pick_theme(company_name, business_type):
    key = (company_name + " " + business_type).lower()
    if any(k in key for k in ["tech", "ai", "saas", "software"]):
        return THEME_MAP["tech"]
    if any(k in key for k in ["food", "organic", "agri", "restaurant"]):
        return THEME_MAP["food"]
    if any(k in key for k in ["finance", "bank", "invest"]):
        return THEME_MAP["finance"]
    if any(k in key for k in ["education", "edtech", "school"]):
        return THEME_MAP["education"]
    if any(k in key for k in ["health", "wellness", "clinic"]):
        return THEME_MAP["health"]
    return THEME_MAP["default"]

def make_gradient_image(color_from, color_to, size=(1280, 720), vertical=True, blur=True):
    """Create a gradient image (Pillow) and return bytes."""
    img = Image.new("RGB", size, color_from)
    draw = ImageDraw.Draw(img)
    for i in range(size[1] if vertical else size[0]):
        ratio = i / float(size[1] - 1) if vertical else i / float(size[0] - 1)
        r = int(color_from[0] + (color_to[0] - color_from[0]) * ratio)
        g = int(color_from[1] + (color_to[1] - color_from[1]) * ratio)
        b = int(color_from[2] + (color_to[2] - color_from[2]) * ratio)
        if vertical:
            draw.line([(0, i), (size[0], i)], fill=(r, g, b))
        else:
            draw.line([(i, 0), (i, size[1])], fill=(r, g, b))
    if blur:
        img = img.filter(ImageFilter.GaussianBlur(radius=8))
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return bio

def set_slide_background_image(slide, img_bytes):
    """Set an image as slide background - python-pptx technique: add full-screen picture as z-order background."""
    pic = slide.shapes.add_picture(img_bytes, Inches(0), Inches(0), width=slide.part.slide_width, height=slide.part.slide_height)
    # Optionally send to back - python-pptx doesn't have z-order send_to_back API, but adding first usually works.

def fetch_image_bytes(url):
    try:
        r = requests.get(url, timeout=8)
        r.raise_for_status()
        return io.BytesIO(r.content)
    except Exception as e:
        return None

def safe_text(value):
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n".join([str(x) for x in value])
    return str(value)

# ---------- PPTX ASSEMBLY ----------
@app.route("/generate", methods=["POST"])
def generate():
    payload = request.get_json(force=True)
    if not payload:
        return jsonify({"error": "no json payload provided"}), 400

    company = payload.get("company_name", "Company")
    business_type = payload.get("nature_of_business", "")
    theme_from, theme_to = pick_theme(company, business_type)
    color_from = hex_to_rgb(theme_from)
    color_to = hex_to_rgb(theme_to)
    font_title = FONT_NAME
    font_body = FONT_NAME

    prs = Presentation()

    # create gradient image once for slides; we will vary slightly per slide to create unique themes
    base_gradient = make_gradient_image(color_from, color_to, size=(1280, 720), vertical=True, blur=True)

    # OPTIONAL: try to extract logo from provided logo_url or from template PPTX asset
    logo_img_bytes = None
    if payload.get("logo_url"):
        logo_img_bytes = fetch_image_bytes(payload["logo_url"])

    # attempt to extract first image from your TEMPLATE_PPTX_PATH if no logo was found
    if not logo_img_bytes and os.path.exists(TEMPLATE_PPTX_PATH):
        try:
            tpl = Presentation(TEMPLATE_PPTX_PATH)
            for s in tpl.slides:
                for shp in s.shapes:
                    if shp.shape_type == 13 and hasattr(shp, "image"):  # picture
                        img = shp.image
                        logo_img_bytes = io.BytesIO(img.blob)
                        break
                if logo_img_bytes:
                    break
        except Exception:
            logo_img_bytes = None

    # ---------------------------------
    # Slide builder helper
    # ---------------------------------
    def add_styled_slide(title_text, body_lines=None, slide_index=0, use_gradient=True, chart=None):
        # choose gradient variant per slide index to vary themes
        grad = make_gradient_image(
            tuple(int((color_from[i] + (color_to[i] - color_from[i]) * (slide_index/12))) for i in range(3)),
            tuple(int((color_from[i] + (color_to[i] - color_from[i]) * ((slide_index+1)/12))) for i in range(3)),
            size=(1280, 720), vertical=True, blur=True
        )
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # use blank layout for full control
        set_slide_background_image(slide, grad)

        # Title box
        left, top, width, height = Inches(0.5), Inches(0.3), Inches(9), Inches(1)
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_tf = title_box.text_frame
        title_tf.text = title_text
        title_p = title_tf.paragraphs[0]
        title_p.font.size = Pt(TITLE_FONT_SIZE)
        title_p.font.bold = True
        title_p.font.name = font_title
        title_p.font.color.rgb = RGBColor(255, 255, 255)

        # Insert logo at top-right
        if logo_img_bytes:
            try:
                slide.shapes.add_picture(logo_img_bytes, prs.slide_width - Inches(2.2), Inches(0.2), width=Inches(2))
                # reset stream position for reuse
                logo_img_bytes.seek(0)
            except Exception:
                pass

        # Body
        if body_lines:
            left, top, width, height = Inches(0.5), Inches(1.6), Inches(9), Inches(4.5)
            tb = slide.shapes.add_textbox(left, top, width, height)
            tf = tb.text_frame
            tf.word_wrap = True
            for i, line in enumerate(body_lines):
                if i == 0:
                    tf.text = line
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                    p.text = line
                p.level = 0
                p.font.size = Pt(BODY_FONT_SIZE)
                p.font.name = font_body
                p.font.color.rgb = RGBColor(255, 255, 255)
        # Chart (optional)
        if chart:
            # chart must be a dict: {"type":"pie" or "bar", "categories": [...], "values":[...]}
            try:
                chart_data = ChartData()
                chart_data.categories = chart.get("categories", [])
                chart_data.add_series("Series 1", chart.get("values", []))
                left_c, top_c, w_c, h_c = Inches(6.5), Inches(2), Inches(3.5), Inches(3)
                if chart.get("type") == "pie":
                    slide.shapes.add_chart(XL_CHART_TYPE.PIE, left_c, top_c, w_c, h_c, chart_data)
                else:
                    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, left_c, top_c, w_c, h_c, chart_data)
            except Exception as e:
                print("chart error", e)

        return slide

    # ---------------------------------
    # Build slides (12)
    # ---------------------------------
    # Slide 1: Company hero
    hero_lines = [safe_text(payload.get("tagline") or payload.get("company_name")), safe_text(payload.get("nature_of_business"))]
    add_styled_slide(payload.get("company_name", "Company"), body_lines=hero_lines, slide_index=0)

    # Slide 2: Nature of business / Overview
    overview = [safe_text(payload.get("nature_of_business", "")), safe_text(payload.get("short_description", ""))]
    add_styled_slide("Nature of Business", body_lines=overview, slide_index=1)

    # Slide 3: Vision + Mission
    vm_lines = [f"Vision: {safe_text(payload.get('vision',''))}", f"Mission: {safe_text(payload.get('mission',''))}"]
    add_styled_slide("Vision & Mission", body_lines=vm_lines, slide_index=2)

    # Slide 4: Problems
    problems = [f"• {p}" for p in payload.get("consumer_problems", [])] or ["No specific problems provided."]
    add_styled_slide("Problems Faced by Consumers", body_lines=problems, slide_index=3)

    # Slide 5: Solutions
    solutions = [f"• {s}" for s in payload.get("solutions_provided", [])] or ["No solutions provided."]
    add_styled_slide("Solutions Provided", body_lines=solutions, slide_index=4)

    # Slide 6: Products / Services
    prods = [f"• {p}" for p in payload.get("products_services", [])] or ["No products/services listed."]
    add_styled_slide("Products & Services", body_lines=prods, slide_index=5)

    # Slide 7: Market share (pie)
    market = payload.get("market_share", [])
    if isinstance(market, dict):
        # accept either dict or list
        market_items = [{"segment": k, "percent": v} for k, v in market.items()]
    else:
        market_items = market
    categories = [m.get("segment","") for m in market_items] if market_items else ["N/A"]
    values = [float(m.get("percent",0)) for m in market_items] if market_items else [100]
    add_styled_slide("Market Share", body_lines=[f"Total market coverage: {sum(values)}%"], slide_index=6,
                     chart={"type":"pie", "categories":categories, "values":values})

    # Slide 8: Target Market
    target_lines = [f"• {t}" for t in payload.get("target_market", [])] or [payload.get("target_market","Not specified")]
    add_styled_slide("Target Market", body_lines=target_lines, slide_index=7)

    # Slide 9: USP
    usp_lines = [f"• {u}" for u in (payload.get("usp") or [])] or [payload.get("usp","Not specified")]
    add_styled_slide("Unique Selling Proposition (USP)", body_lines=usp_lines, slide_index=8)

    # Slide 10: Contact Details
    contact = payload.get("company_contact", payload.get("contact_details", "N/A"))
    add_styled_slide("Contact Details", body_lines=[safe_text(contact)], slide_index=9)

    # Slide 11: Directors
    directors = payload.get("directors", [])
    if not directors and isinstance(payload.get("director_name"), str):
        directors = [{
            "name": payload.get("director_name"),
            "phone": payload.get("director_phone"),
            "email": payload.get("director_email"),
            "education": payload.get("director_qualification")
        }]
    dir_lines = []
    for d in directors:
        line = f"Name: {d.get('name','')}\nPhone: {d.get('phone','')}\nEmail: {d.get('email','')}\nEducation: {d.get('education','')}"
        dir_lines.append(line)
    add_styled_slide("Directors", body_lines=dir_lines or ["No director data provided."], slide_index=10)

    # Slide 12: Fund Deployment (pie)
    fund = payload.get("fund_deployment", {})
    if not fund:
        # fallback to individual keys
        fund = {
            "testing_manufacturing": float(payload.get("fund_deployment_testing", 0)),
            "man_power": float(payload.get("fund_deployment_manpower", 0)),
            "outsourced_services": float(payload.get("fund_deployment_outsourced", 0)),
            "ip_costs": float(payload.get("fund_deployment_ip_costs", 0)),
            "travel": float(payload.get("fund_deployment_travel", 0)),
            "consumables": float(payload.get("fund_deployment_consumables", 0)),
            "contingency": float(payload.get("fund_deployment_contingency", 0)),
            "others": float(payload.get("fund_deployment_others", 0)),
        }
    categories_fd = list(fund.keys()) if fund else ["N/A"]
    values_fd = [float(v) for v in fund.values()] if fund else [100]
    add_styled_slide("Fund Deployment", body_lines=[f"Total: {sum(values_fd)}%"], slide_index=11,
                     chart={"type":"pie","categories":categories_fd,"values":values_fd})

    # Save to temporary file and return
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)
    tmp.flush()
    tmp.seek(0)

    return send_file(tmp.name, as_attachment=True, download_name=f"{company}_pitchdeck.pptx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001)
