# Composes ad creatives (1:1 1080x1080 and 4:5 1080x1350) from ChatGPT base images + headline bands.
from PIL import Image, ImageDraw, ImageFont, ImageOps
import os, sys
SRC = "build/creatives"; OUT = "product/ads"; os.makedirs(OUT, exist_ok=True)
CREAM=(250,247,242); INK=(31,61,43); SAGE=(127,169,139); MUST=(227,178,60); WHITE=(255,255,255)
def font(size, bold=True):
    for p in ["/System/Library/Fonts/Avenir Next.ttc", "/System/Library/Fonts/HelveticaNeue.ttc", "/System/Library/Fonts/Helvetica.ttc"]:
        if os.path.exists(p):
            try:
                idx = 0
                # Avenir Next.ttc: find Bold/Demi face index by trying
                for i in range(0, 12):
                    f = ImageFont.truetype(p, size, index=i)
                    name = f.getname()[1].lower()
                    if (bold and ("bold" in name or "demi" in name) and "italic" not in name) or (not bold and name in ("regular","medium")):
                        return f
                return ImageFont.truetype(p, size, index=0)
            except Exception:
                continue
    return ImageFont.load_default()

def wrap(draw, text, f, maxw):
    words = text.split(); lines=[]; cur=""
    for w in words:
        t = (cur+" "+w).strip()
        if draw.textlength(t, font=f) <= maxw: cur = t
        else: lines.append(cur); cur = w
    if cur: lines.append(cur)
    return lines

def cover(im, W, H):
    return ImageOps.fit(im, (W, H), method=Image.LANCZOS, centering=(0.5, 0.5))

def compose(src, W, H, headline, sub, tag, out, band_h=None):
    im = Image.open(os.path.join(SRC, src)).convert("RGB")
    band_h = band_h or int(H*0.36)
    canvas = Image.new("RGB", (W, H), CREAM)
    img_h = H - band_h
    canvas.paste(cover(im, W, img_h), (0, 0))
    d = ImageDraw.Draw(canvas)
    y0 = img_h; pad = int(W*0.065)
    # row 1: tag pill left, wordmark right
    ft = font(int(W*0.024)); tw = d.textlength(tag, font=ft)
    ty = y0 + int(W*0.035); th = int(W*0.046)
    d.rounded_rectangle([pad, ty, pad+tw+34, ty+th], radius=th//2, fill=(228,239,231))
    d.text((pad+17, ty+int(th*0.22)), tag, font=ft, fill=INK)
    fw = font(int(W*0.026)); ww = d.textlength("Tidy Money", font=fw)
    d.text((W-pad-ww, ty+int(th*0.18)), "Tidy Money", font=fw, fill=SAGE)
    # headline (max 2 lines)
    size = int(W*0.058)
    while True:
        fh = font(size); lines = wrap(d, headline, fh, W-2*pad)
        if len(lines) <= 2 or size < int(W*0.04): break
        size -= 2
    y = ty + th + int(W*0.022)
    for ln in lines:
        d.text((pad, y), ln, font=fh, fill=INK); y += int(size*1.15)
    # sub (max 2 lines)
    fs = font(int(W*0.027), bold=False)
    for ln in wrap(d, sub, fs, W-2*pad)[:2]:
        d.text((pad, y+4), ln, font=fs, fill=(95,95,95)); y += int(W*0.036)
    # price chip bottom-left
    fp = font(int(W*0.027)); ptxt = "$12  ·  one-time  ·  instant download"
    pw = d.textlength(ptxt, font=fp); ph = int(W*0.05)
    py = max(y + int(W*0.03), H - int(W*0.04) - ph)
    d.rounded_rectangle([pad, py, pad+pw+40, py+ph], radius=ph//2, fill=INK)
    d.text((pad+20, py+int(ph*0.2)), ptxt, font=fp, fill=WHITE)
    canvas.save(os.path.join(OUT, out), quality=92)
    print("saved", out)

SPECS = [
    # (src, headline, sub, tag, slug)
    ("img1_laptop.png", "See exactly where your money goes.", "Log spending once. Every tab updates itself. Google Sheets & Excel.", "BUDGET SPREADSHEET", "a1_clarity"),
    ("img2_couple.png", "One budget sheet for both of you.", "Same plan, same numbers, both phones. Google Sheets & Excel.", "FOR COUPLES", "a3_couples"),
    ("img3_desk.png", "A budget that does the math for you.", "Set up in 10 minutes. No app, no subscription.", "NO SUBSCRIPTION", "a2_simple"),
    ("img4_chart.png", "Planned vs. actual, at a glance.", "Budget, savings goals, debt payoff, net worth. One calm file.", "10 TABS, ONE FILE", "a4_pretty"),
]
only = sys.argv[1:] 
for src, h, s, t, slug in SPECS:
    if not os.path.exists(os.path.join(SRC, src)): print("skip (missing)", src); continue
    if only and slug not in only: continue
    compose(src, 1080, 1080, h, s, t, f"{slug}_1x1.jpg")
    compose(src, 1080, 1350, h, s, t, f"{slug}_4x5.jpg", band_h=420)
