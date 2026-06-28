# -*- coding: utf-8 -*-
"""Approximate PIL renderer for a PPTX to catch layout/overflow issues."""
import io, os
from pptx import Presentation
from pptx.util import Emu
from PIL import Image, ImageDraw, ImageFont
from pptx.oxml.ns import qn

PPTX = "/home/user/licenta/Prezentare_Licenta_Template.pptx"
OUT = "/tmp/claude-0/-home-user-licenta/03ba0bac-fff2-5a16-859d-96a3dba91012/scratchpad"
REG = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
BLD = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
ITA = "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf"

prs = Presentation(PPTX)
SW = prs.slide_width; SH = prs.slide_height
PXW = 1600; SCALE = PXW / SW; PXH = int(SH * SCALE)

def font(sz, bold, ital):
    px = max(8, int(sz * SCALE * 12700 / 12700 * (72/72) ))  # sz is pt
    px = max(8, int(sz * SCALE * 914400/72 / 914400 * 72 * (PXW/ (SW/914400*72)) /72))
    return None

def fpx(pt):  # pt -> px
    return max(7, int(pt * PXW / (SW/914400) / 72))

def getfont(pt, bold=False, ital=False):
    path = BLD if bold else (ITA if ital else REG)
    return ImageFont.truetype(path, fpx(pt))

def emu_to_px(v):
    return int(v * SCALE)

def blip_image(shape, part):
    el = shape._element
    blips = el.findall(".//" + qn("a:blip"))
    for b in blips:
        rid = b.get(qn("r:embed"))
        if rid and rid in part.rels:
            return part.rels[rid].target_part.blob
    return None

def draw_shape(d, img, shape, part):
    try:
        l = emu_to_px(shape.left); t = emu_to_px(shape.top)
        w = emu_to_px(shape.width); h = emu_to_px(shape.height)
    except Exception:
        return
    # picture
    if shape.shape_type == 13:
        try:
            blob = shape.image.blob
            pim = Image.open(io.BytesIO(blob)).convert("RGBA")
            pim = pim.resize((max(1,w), max(1,h)))
            img.paste(pim, (l, t), pim)
            return
        except Exception:
            pass
    # autoshape/rect with solid fill or blip fill
    fill_rgb = None
    try:
        if shape.fill.type == 1:
            c = shape.fill.fore_color.rgb
            fill_rgb = (c[0], c[1], c[2])
    except Exception:
        pass
    blob = blip_image(shape, part)
    if blob is not None and shape.shape_type != 13:
        try:
            pim = Image.open(io.BytesIO(blob)).convert("RGBA").resize((max(1,w),max(1,h)))
            img.paste(pim, (l,t), pim)
        except Exception:
            pass
    elif fill_rgb is not None:
        d.rectangle([l, t, l+w, t+h], fill=fill_rgb)
    # group: recurse
    if shape.shape_type == 6:
        return  # skip decorative groups
    # text
    if shape.has_text_frame and shape.text_frame.text.strip():
        ty = t + emu_to_px(Emu(int(0.03*914400)))
        for para in shape.text_frame.paragraphs:
            runs = para.runs
            if not runs:
                continue
            txt = "".join(r.text for r in runs).replace("\n", " ").replace("\v", " ")
            r0 = runs[0]
            pt = r0.font.size.pt if r0.font.size else 18
            bold = bool(r0.font.bold)
            ital = bool(r0.font.italic)
            col = (48,54,66)
            try:
                if r0.font.color and r0.font.color.type is not None:
                    c = r0.font.color.rgb; col = (c[0],c[1],c[2])
            except Exception:
                pass
            f = getfont(pt, bold, ital)
            align = para.alignment
            # wrap
            maxw = w - 4
            words = txt.split(" ")
            line = ""; lines = []
            for wd in words:
                test = (line + " " + wd).strip()
                if d.textlength(test, font=f) <= maxw or not line:
                    line = test
                else:
                    lines.append(line); line = wd
            if line: lines.append(line)
            for ln in lines:
                tw = d.textlength(ln, font=f)
                if align == 2:  # center
                    tx = l + (w - tw)//2
                elif align == 3:  # right
                    tx = l + w - tw
                else:
                    tx = l
                d.text((tx, ty), ln, fill=col, font=f)
                ty += fpx(pt) * 1.18

paths = []
for i, slide in enumerate(prs.slides):
    img = Image.new("RGBA", (PXW, PXH), (250,250,250,255))
    d = ImageDraw.Draw(img)
    for shape in slide.shapes:
        draw_shape(d, img, shape, slide.part)
    d.rectangle([0,0,PXW-1,PXH-1], outline=(180,180,180))
    p = f"{OUT}/slide_{i+1:02d}.png"
    img.convert("RGB").save(p)
    paths.append(p)
    print(p)

# contact sheet 3 cols
cols=3; rows=(len(paths)+cols-1)//cols
tw=PXW//3; th=PXH//3
sheet=Image.new("RGB",(tw*cols, th*rows),(255,255,255))
for i,p in enumerate(paths):
    im=Image.open(p).resize((tw,th))
    sheet.paste(im,((i%cols)*tw,(i//cols)*th))
sheet.save(f"{OUT}/contact_sheet.png")
print("SHEET", f"{OUT}/contact_sheet.png")
