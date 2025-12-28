import io
import re
import zipfile
import img2pdf
from PIL import ImageDraw, ImageFont, Image

def has_descenders(text):
    return bool(re.search('[fgjpqy]', text.lower()))

def get_font(font_bytes, font_size):
    try:
        return ImageFont.truetype(io.BytesIO(font_bytes), font_size)
    except Exception:
        return ImageFont.load_default()

def generate_certificate_image(name, offset, template_img, font, text_color):
    certificate = template_img.copy()
    draw = ImageDraw.Draw(certificate)

    W, H = certificate.size
    _, _, w1, h1 = draw.textbbox((0, 0), name, font=font, stroke_width=0)
    
    left1 = (W - w1) / 2
    top1 = (H - h1) / 2 + offset

    draw.text((left1, top1), name, font=font, fill=text_color)
    return certificate

def create_final_bundle(normal_names, desc_names, template_img, loaded_font, color, off_norm=0, off_desc=0,):
    zip_buffer = io.BytesIO()
    
    if template_img.mode in ("RGBA", "LA"):
        background = Image.new("RGB", template_img.size, (255, 255, 255))
        background.paste(template_img, mask=template_img.split()[-1])
        base_template = background
    else:
        base_template = template_img.convert("RGB")
    
    tasks = [
        (normal_names, off_norm),
        (desc_names, off_desc)
    ]
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for name_list, offset in tasks:
            if len(name_list) != 0:
                for name in name_list:
                    
                    img = generate_certificate_image(name, offset, base_template, loaded_font, color)
                    
                    safe_name = "".join(c for c in name if c.isalnum() or c in " _-").rstrip()
                    
                    img_bytes_io = io.BytesIO()
                    img.save(img_bytes_io, format="PNG")
                    img_data = img_bytes_io.getvalue()
                    
                    # Internal folder structure per your requirement
                    zip_file.writestr(f"{safe_name}/{safe_name}.png", img_data)
                    zip_file.writestr(f"{safe_name}/{safe_name}.pdf", img2pdf.convert(img_data))
                
    zip_buffer.seek(0)
    return zip_buffer

