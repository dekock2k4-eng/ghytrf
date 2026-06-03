"""
generate_logo.py — render the SheetMind brand mark + Office icon set as PNGs.

Mark: a rounded-square tile with an emerald→teal gradient (a refined nod to
Excel's green) and a crisp white "spark" glyph (AI) over a subtle data baseline.
Rendered at 4x and downscaled for clean anti-aliased edges.
"""
import math
import os

from PIL import Image, ImageDraw, ImageFilter

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "addin_web", "assets")
os.makedirs(OUT, exist_ok=True)

SS = 4  # supersample factor


def _lerp(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def _gradient(size, top, mid, bot):
    img = Image.new("RGB", (size, size))
    px = img.load()
    for y in range(size):
        t = y / (size - 1)
        if t < 0.5:
            c = _lerp(top, mid, t * 2)
        else:
            c = _lerp(mid, bot, (t - 0.5) * 2)
        for x in range(size):
            px[x, y] = c
    return img


def _rounded_mask(size, radius):
    m = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return m


def _spark(cx, cy, outer, inner, rot=0):
    pts = []
    for i in range(8):
        ang = math.radians(rot + i * 45)
        r = outer if i % 2 == 0 else inner
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    return pts


def make_tile(size, transparent_bg=True):
    S = size * SS
    # Gradient tile
    grad = _gradient(S, (110, 231, 183), (16, 185, 129), (13, 148, 136))  # mint→emerald→teal
    radius = int(S * 0.24)
    mask = _rounded_mask(S, radius)
    tile = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    tile.paste(grad, (0, 0), mask)

    d = ImageDraw.Draw(tile)

    # Soft top gloss for depth
    gloss = Image.new("L", (S, S), 0)
    gd = ImageDraw.Draw(gloss)
    gd.ellipse([-S * 0.3, -S * 0.7, S * 1.3, S * 0.45], fill=60)
    gloss = gloss.filter(ImageFilter.GaussianBlur(S * 0.04))
    white = Image.new("RGBA", (S, S), (255, 255, 255, 255))
    # Apply the gloss only inside the rounded tile.
    tile.paste(white, (0, 0), Image.composite(gloss, Image.new("L", (S, S), 0), mask))

    d = ImageDraw.Draw(tile)
    cx, cy = S * 0.5, S * 0.46

    # Data baseline: two soft rounded bars under the spark
    bar_w, bar_h = S * 0.42, S * 0.052
    by = S * 0.74
    for i, w in enumerate((bar_w, bar_w * 0.62)):
        x0 = cx - bar_w / 2
        y0 = by + i * (bar_h + S * 0.045)
        d.rounded_rectangle([x0, y0, x0 + w, y0 + bar_h], radius=bar_h / 2,
                            fill=(255, 255, 255, 200))

    # Main spark (white)
    main = _spark(cx, cy, S * 0.26, S * 0.072)
    d.polygon(main, fill=(255, 255, 255, 255))
    # Small accent spark top-right
    acc = _spark(S * 0.72, S * 0.26, S * 0.085, S * 0.026)
    d.polygon(acc, fill=(255, 255, 255, 235))

    tile = tile.resize((size, size), Image.LANCZOS)
    if not transparent_bg:
        bg = Image.new("RGBA", (size, size), (255, 255, 255, 255))
        bg.alpha_composite(tile)
        return bg
    return tile


def make_wordmark(width=720, height=200):
    """Logo tile + 'SheetMind' wordmark — used in the task-pane header (rendered
    larger; the UI also has a CSS wordmark, this PNG is a fallback/share asset)."""
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    tile = make_tile(int(height * 0.78))
    ty = (height - tile.height) // 2
    img.alpha_composite(tile, (8, ty))
    return img


if __name__ == "__main__":
    # Office icon set
    for s in (16, 32, 64, 80, 128, 300):
        make_tile(s).save(os.path.join(OUT, f"icon-{s}.png"))
    # Hi-res logo + wordmark
    make_tile(512).save(os.path.join(OUT, "logo.png"))
    make_wordmark().save(os.path.join(OUT, "wordmark.png"))
    print("Wrote logos to", OUT)
    print("  ", ", ".join(sorted(os.listdir(OUT))))
