"""Render the masthead of The Blank, Blank University's student newspaper, as a transparent PNG.

    python3 tools/make-the-blank-masthead.py engine/shared-assets/the-blank-masthead.png

Needs Pillow and the macOS fonts Bodoni 72 and Didot. Colours are sampled from Blank University's logo."""
import math, sys
from PIL import Image, ImageDraw, ImageFont

NAVY = (6, 40, 67, 255)   # sampled from Blank University Logo.png
GOLD = (211, 167, 83, 255)
BODONI = "/System/Library/Fonts/Supplemental/Bodoni 72.ttc"
DIDOT = "/System/Library/Fonts/Supplemental/Didot.ttc"

S = 4  # supersample, then shrink for smooth edges
W, H = 1200, 193
im = Image.new("RGBA", (W * S, H * S), (0, 0, 0, 0))
d = ImageDraw.Draw(im)

def rule(y, h):
    d.rectangle([40 * S, y * S, (W - 40) * S, (y + h) * S], fill=NAVY)

def star(cx, cy, r):
    pts = []
    for i in range(10):
        a = -math.pi / 2 + i * math.pi / 5
        rr = r if i % 2 == 0 else r * 0.42
        pts.append(((cx + rr * math.cos(a)) * S, (cy + rr * math.sin(a)) * S))
    d.polygon(pts, fill=GOLD)

def centred(text, font, y, fill, tracking=0):
    widths = [font.getlength(ch) + tracking * S for ch in text]
    x = (W * S - (sum(widths) - tracking * S)) / 2
    for ch, w in zip(text, widths):
        d.text((x, y * S), ch, font=font, fill=fill)
        x += w
    return sum(widths) / S

# Rules above: thin over thick
rule(12, 2); rule(18, 5)

# Title
title = ImageFont.truetype(BODONI, 128 * S, index=2)
tw = title.getlength("The Blank") / S
top = title.getbbox("The Blank")[1] / S
d.text(((W - tw) / 2 * S, (34 - top) * S), "The Blank", font=title, fill=NAVY)

# Tagline between two short rules with stars
tag = ImageFont.truetype(DIDOT, 18 * S, index=2)
tag_w = centred("THE STUDENT NEWSPAPER OF BLANK UNIVERSITY", tag, 142, NAVY, tracking=4)
ty = 153
for side in (-1, 1):
    inner = W / 2 + side * (tag_w / 2 + 22)
    outer = 40 if side < 0 else W - 40
    a, b = sorted((inner + side * 20, outer))
    d.rectangle([a * S, (ty - 1) * S, b * S, (ty + 1) * S], fill=NAVY)
    star(inner + side * 6, ty, 8)

# Rules below: thick over thin
rule(173, 5); rule(181, 2)

out = im.resize((W, H), Image.LANCZOS)
out.save(sys.argv[1])
