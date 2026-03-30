"""
Grameen Bank Employee Card Generator — Final Corrected Version
Features:
- Fixed NameError (excel_path -> EXCEL_PATH)
- Green & Red color scheme from logo
- Updated LOGO_PATH for local Windows environment
- "Card Holder Sign" label
- Inward-facing semi-circle on the back side
- QR Code integration
"""
import os, io, sys, argparse, math
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from PIL import Image, ImageDraw, ImageFont
import qrcode

# ══════════════════════════════════════════════════════════════
#  DIMENSIONS & SETTINGS
# ══════════════════════════════════════════════════════════════
EXCEL_PATH    = "EmployeeList.xlsx"
OUTPUT_DIR    = "output_cards"
OUTPUT_FORMAT = "PNG"
DPI           = 300
CARD_W=215; CARD_H=310; SCALE=4
RW=CARD_W*SCALE; RH=CARD_H*SCALE
GAP_PX=20; GAP_R=GAP_PX*SCALE; SHEET_RH=RH*2+GAP_R

# COLORS (Extracted from logo)
C_WHITE = (255, 255, 255)
C_BLACK = (12, 12, 12)
C_RED   = (237, 28, 36)
C_GREEN = (0, 166, 81)
C_DARK_GREY = (55, 55, 55)
C_LIGHT_GREY = (180, 180, 180)
C_BLUE = (0,0,255)

# LOGO PATH (Updated as requested)
LOGO_PATH = r"D:\Grameen Bank\Auto Employee Card Generator\logo\grameelogo.png"

def _find_font(*names):
    for d in ["/usr/share/fonts","/usr/local/share/fonts",os.path.expanduser("~/.fonts"),
              "C:/Windows/Fonts","/Library/Fonts",os.path.expanduser("~/Library/Fonts")]:
        if not os.path.exists(d): continue
        for root,_,files in os.walk(d):
            for fn in files:
                if fn.lower() in [n.lower() for n in names]: return os.path.join(root,fn)
    return None

def load_fonts():
    r=_find_font("DejaVuSerif.ttf","FreeSerif.ttf","times.ttf")
    b=_find_font("DejaVuSerif-Bold.ttf","FreeSerifBold.ttf","timesbd.ttf")
    def f(p,sz):
        try: return ImageFont.truetype(p or r,sz*SCALE)
        except: return ImageFont.load_default()
    return {
        "name": f(b,14), "pos": f(r,9), "label": f(r,8), "value": f(b,8),
        "back_text": f(r,8.5), "logo_text": f(b,12), "logo_sub": f(r,7),
        "sig_label": f(r, 7.5)
    }

def draw_dashed_circle(draw, center, radius, fill, width, dash_length):
    cx, cy = center
    circumference = 2 * math.pi * radius
    dashes = int(circumference / dash_length)
    for i in range(dashes):
        if i % 2 == 0:
            start_angle = i * (360 / dashes)
            end_angle = (i + 1) * (360 / dashes)
            draw.arc([cx - radius, cy - radius, cx + radius, cy + radius], start_angle, end_angle, fill=fill, width=width)

def paste_photo(img, raw, cx, cy, radius):
    try:
        photo = Image.open(io.BytesIO(raw)).convert("RGB")
        w, h = photo.size; side = min(w, h)
        photo = photo.crop(((w-side)//2, (h-side)//2, (w+side)//2, (h+side)//2))
        photo = photo.resize((radius*2, radius*2), Image.LANCZOS)
        mask = Image.new("L", (radius*2, radius*2), 0)
        ImageDraw.Draw(mask).ellipse([0, 0, radius*2, radius*2], fill=255)
        img.paste(photo, (cx - radius, cy - radius), mask)
    except: pass

def paste_logo(img, path, x, y, mw, mh):
    # Try the user path, then fallback to current dir
    paths = [path, "grameelogo.png"]
    for p in paths:
        if os.path.isfile(p):
            try:
                lg = Image.open(p).convert("RGBA"); lg.thumbnail((mw, mh), Image.LANCZOS)
                img.paste(lg, (x, y), lg); return lg.size
            except: pass
    return 0, 0

def generate_qr(data, size):
    if not data or not data.strip():  # ✅ guard
        return Image.new("RGB", (size, size), (255, 255, 255))  # blank white
    qr = qrcode.QRCode(version=1, box_size=6, border=1)  # ✅ small but scannable
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    return img.resize((size, size), Image.LANCZOS)

def draw_front(card, fonts, raw_photo=None, raw_ch_sign=None, raw_auth_sign=None):
    img = Image.new("RGB", (RW, RH), C_WHITE)
    d = ImageDraw.Draw(img)
    
    # Side bar using Green
    side_bar_w = int(RW * 0.35)
    d.rectangle([RW - side_bar_w, 0, RW, RH], fill=C_GREEN)
    
    # Logo
    lw, lh = paste_logo(img, LOGO_PATH, 80, 80, 200, 80)
    if lw == 0:
        d.text((80, 80), "Grameen Bank", font=fonts["logo_text"], fill=C_GREEN)
        d.text((80, 135), "Nobel Peace Prize Winner", font=fonts["logo_sub"], fill=C_RED)
        
    # Photo Zone
    cx, cy = int(RW * 0.45), int(RH * 0.4)
    outer_r, dash_r, photo_r = 280, 260, 240
    d.ellipse([cx - outer_r, cy - outer_r, cx + outer_r, cy + outer_r], fill=C_WHITE)
    draw_dashed_circle(d, (cx, cy), dash_r, fill=C_LIGHT_GREY, width=4, dash_length=20)
    if raw_photo: paste_photo(img, raw_photo, cx, cy, photo_r)
    else: d.ellipse([cx - photo_r, cy - photo_r, cx + photo_r, cy + photo_r], fill=(230, 235, 240))
        
    # Employee Info
    name_y = cy + outer_r + 40
    d.text((80, name_y), card["name"].upper(), font=fonts["name"], fill=C_BLUE)
    d.text((80, name_y + 60), card["designation"].upper(), font=fonts["pos"], fill=C_RED)
    det_y = name_y + 140
    # Mobile No fix
    mobile_raw = str(card["mobile"]).strip().split(".")[0]
    mobile_fixed = mobile_raw if mobile_raw.startswith("0") else "0" + mobile_raw
    details = [("ID No", f": {card['id']}"), ("Blood Group", f": {card['blood_group']}"),
               ("Mobile", f": {mobile_fixed}"), ("Issue Date", f": {card['issuing_date']}")]
    for lbl, val in details:
        d.text((80, det_y), lbl, font=fonts["label"], fill=C_DARK_GREY)
        d.text((240, det_y), val, font=fonts["value"], fill=C_BLACK)
        det_y += 45
    return img

def draw_back(card, fonts, raw_ch_sign=None):  # ✅ add raw_ch_sign
    img = Image.new("RGB", (RW, RH), C_WHITE)
    d = ImageDraw.Draw(img)

    side_bar_w = int(RW * 0.35)
    d.rectangle([RW - side_bar_w, 0, RW, RH], fill=C_GREEN)

    cx, cy = RW - side_bar_w, int(RH * 0.4)
    r = 240
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=C_WHITE)

    paste_logo(img, LOGO_PATH, 80, 80, 200, 80)

    y_text = 350
    lines = [
        "If found, please return to:", "",
        "Grameen Bank", card["bank_address"], "",
        "Phone: 028694485", "",
        "Email: g_iprog@grameenbank.org.bd"
    ]
    #    d.text((80, name_y), card["name"].upper(), font=fonts["name"], fill=C_GREEN)
    for line in lines:
        d.text((80, y_text), line, font=fonts["back_text"], fill=C_DARK_GREY)
        y_text += 40

    # Signature
    y_sig = 850
    d.line([(80, y_sig + 50), (400, y_sig + 50)], fill=C_LIGHT_GREY, width=2)

    # ✅ Paste signature image if available, else fallback to text
    if raw_ch_sign:
        try:
            sign_img = Image.open(io.BytesIO(raw_ch_sign)).convert("RGBA")
            sign_img.thumbnail((300, 80), Image.LANCZOS)
            img.paste(sign_img, (80, y_sig - 30), sign_img)
        except:
            pass

    d.text((160, y_sig + 60), "Card Holder Sign", font=fonts["sig_label"], fill=C_DARK_GREY)

    # QR Code
    qr_data = (
    f"Name: {card['name']}\n"
    f"ID: {card['id']}\n"
    f"Mobile: {card['mobile']}\n"
    )
    
    # In draw_back — before generating QR
    mobile_raw = str (card["mobile"]).strip().split(".")[0]
    mobile_fixed = mobile_raw if mobile_raw.startswith("0") else "0" + mobile_raw
    qr_data = f"{card['id']}|{card['name']}|{mobile_fixed}"

    if qr_data.strip("|").strip():  # ✅ only generate if data exists
        qr_img = generate_qr(qr_data, 220)
        img.paste(qr_img, (280, 1000))
    else:
        print(f"⚠️ Skipping QR — no data for card: {card['id']}")
    return img

def get_images(excel_path):
    wb = load_workbook(excel_path); ws = wb.active
    res = {"N": {}, "L": {}, "M": {}}
    for io_ in getattr(ws, '_images', []):
        try:
            a = io_.anchor
            row = (a._from.row+1 if hasattr(a, '_from') else a.row+1)
            col = (a._from.col+1 if hasattr(a, '_from') else a.col+1)
            raw = io_._data(); 
            if callable(raw): raw = raw()
            if col == 14: res["N"][row] = raw # N
            elif col == 12: res["L"][row] = raw # L
            elif col == 13: res["M"][row] = raw # M
        except: pass
    return res

def generate():
    print("Starting...")  # ✅ check if function runs

    if not os.path.exists(EXCEL_PATH):
        print(f"Excel not found at: {EXCEL_PATH}, using dummy data")
        df = pd.DataFrame([{"EmployeeID":"123456", "Name":"MD. RAHIM UDDIN", "Employee Designation":"Senior Officer", "Blood Group":"O+", "Mobile Number":"01711223344", "Issuing Date":"12 Jan 2024", "Bank's Address":"Head Office, Mirpur-2, Dhaka"}])
        images = {"N":{}, "L":{}, "M":{}}
    else:
        print(f"Excel found: {EXCEL_PATH}")
        images = get_images(EXCEL_PATH)
        df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")
        print(f"Rows loaded: {len(df)}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    fonts = load_fonts()
    print("Fonts loaded")

    for i, (_, row) in enumerate(df.iterrows()):
        idx = i + 2
        print(f"Processing row {idx}...")  # ✅ see which row fails

        try:
            card = {
                "id":           str(row.get("EmployeeID", "")),
                "name":         str(row.get("Name", "")),
                "designation":  str(row.get("Employee Designation", "")),
                "blood_group":  str(row.get("Blood Group", "")),
                "mobile":       str(row.get("Mobile Number", "")),
                "issuing_date": str(row.get("Issuing Date", "")),
                "bank_address": str(row.get("Bank's Address", "")),
            }

            raw_photo     = images["N"].get(idx)
            raw_ch_sign   = images["L"].get(idx)
            raw_auth_sign = images["M"].get(idx)

            fr = draw_front(card, fonts, raw_photo, raw_ch_sign, raw_auth_sign)
            bk = draw_back(card, fonts, raw_ch_sign)

            sheet = Image.new("RGB", (RW, SHEET_RH), (235, 237, 240))
            sheet.paste(fr, (0, 0)); sheet.paste(bk, (0, RH + GAP_R))

            dd = ImageDraw.Draw(sheet); dy = RH + GAP_R//2
            for x in range(0, RW, 80): dd.line([(x, dy), (x+48, dy)], fill=(140, 148, 158), width=4)

            final = sheet.resize((CARD_W, CARD_H*2 + GAP_PX), Image.LANCZOS)
            save_path = os.path.join(OUTPUT_DIR, f"{card['id'] or i}.png")
            final.save(save_path, "PNG", dpi=(DPI, DPI))
            print(f"✅ Generated: {save_path}")

        except Exception as e:
            print(f"❌ Error on row {idx}: {e}")  # ✅ shows exact error
            import traceback; traceback.print_exc()

if __name__ == "__main__":
    generate()