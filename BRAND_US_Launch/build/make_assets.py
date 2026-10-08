#!/usr/bin/env python3
"""Fiber strand background for the title and section slides. Dark gradient plus a few thin bright curves with a soft glow."""
import math, random
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter

OUT = Path(__file__).resolve().parent / "assets"
OUT.mkdir(exist_ok=True)
W, H = 1920, 1080
random.seed(7)

def hexrgb(h): return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
BG1, BG2, CARD = hexrgb("0A0F1E"), hexrgb("05070F"), hexrgb("141C36")
CYAN, VIOLET, MAGENTA = hexrgb("2DE2FF"), hexrgb("8B5CF6"), hexrgb("FF4FD8")

def gradient():
    img = Image.new("RGB", (W, H), BG1)
    px = img.load()
    for y in range(H):
        for x in range(W):
            t = (x / W) * 0.55 + (y / H) * 0.45
            px[x, y] = tuple(int(BG1[i] * (1 - t) + BG2[i] * t) for i in range(3))
    # a faint violet bloom bottom right so the strands have something to glow against
    bloom = Image.new("RGB", (W, H), (0, 0, 0))
    bd = ImageDraw.Draw(bloom)
    bd.ellipse([W * 0.55, H * 0.45, W * 1.25, H * 1.35], fill=(28, 18, 60))
    bloom = bloom.filter(ImageFilter.GaussianBlur(220))
    return Image.blend(img, Image.eval(bloom, lambda v: v), 0.0).convert("RGB") if False else Image.fromarray(__import__("numpy").clip(__import__("numpy").array(img).astype(int) + __import__("numpy").array(bloom).astype(int), 0, 255).astype("uint8"))

def strand_points(y0, amp, freq, phase, tilt):
    pts = []
    for x in range(-50, W + 51, 6):
        y = y0 + tilt * (x / W) + amp * math.sin(freq * (x / W) * 2 * math.pi + phase) + 0.35 * amp * math.sin(2.3 * freq * (x / W) * 2 * math.pi + 1.7 * phase)
        pts.append((x, y))
    return pts

def draw_strands(base, strands):
    glow = Image.new("RGB", (W, H), (0, 0, 0))
    gd = ImageDraw.Draw(glow)
    core = Image.new("RGB", (W, H), (0, 0, 0))
    cd = ImageDraw.Draw(core)
    for (y0, amp, freq, phase, tilt, color, width) in strands:
        pts = strand_points(y0, amp, freq, phase, tilt)
        gd.line(pts, fill=color, width=width * 6)
        cd.line(pts, fill=color, width=width)
    glow = glow.filter(ImageFilter.GaussianBlur(26))
    core = core.filter(ImageFilter.GaussianBlur(0.8))
    import numpy as np
    a = np.array(base).astype(int) + (np.array(glow).astype(int) * 0.55).astype(int) + np.array(core).astype(int)
    return Image.fromarray(np.clip(a, 0, 255).astype("uint8"))

def title_bg():
    base = gradient()
    strands = [
        (H * 0.62, 90, 0.9, 0.4, 140, CYAN, 2),
        (H * 0.70, 70, 1.1, 2.1, 110, VIOLET, 2),
        (H * 0.78, 110, 0.7, 1.2, 60, CYAN, 1),
        (H * 0.86, 60, 1.3, 3.0, 40, MAGENTA, 1),
        (H * 0.55, 50, 1.6, 0.9, 200, VIOLET, 1),
        (H * 0.92, 80, 0.8, 2.6, -30, CYAN, 1),
    ]
    img = draw_strands(base, strands)
    # darken the upper left where the title sits
    shade = Image.new("L", (W, H), 0)
    sd = ImageDraw.Draw(shade)
    sd.rectangle([0, 0, W * 0.75, H * 0.62], fill=70)
    shade = shade.filter(ImageFilter.GaussianBlur(160))
    dark = Image.new("RGB", (W, H), BG2)
    img = Image.composite(dark, img, shade.point(lambda v: int(v * 0.9)))
    img.save(OUT / "fiber_bg.png", optimize=True)

def section_bg():
    base = gradient()
    strands = [
        (H * 0.30, 70, 1.0, 1.1, 260, CYAN, 1),
        (H * 0.20, 60, 1.4, 2.4, 300, VIOLET, 1),
        (H * 0.40, 90, 0.8, 0.3, 220, CYAN, 2),
        (H * 0.15, 50, 1.2, 1.9, 380, MAGENTA, 1),
    ]
    img = draw_strands(base, strands)
    shade = Image.new("L", (W, H), 0)
    sd = ImageDraw.Draw(shade)
    sd.rectangle([0, H * 0.35, W * 0.7, H], fill=70)
    shade = shade.filter(ImageFilter.GaussianBlur(160))
    dark = Image.new("RGB", (W, H), BG2)
    img = Image.composite(dark, img, shade.point(lambda v: int(v * 0.9)))
    img.save(OUT / "fiber_section_bg.png", optimize=True)

if __name__ == "__main__":
    title_bg(); section_bg()
    print("wrote", OUT / "fiber_bg.png", OUT / "fiber_section_bg.png")
