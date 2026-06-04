"""Generate icon.ico from owl_source.png.

Crops the surrounding white space off the owl, composites it onto a
white circle with a thin border, scales the owl up to fill the frame, and
writes a multi-resolution icon.ico (16/32/48/64/128/256).
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageOps

HERE = Path(__file__).parent
SRC = HERE / "owl_source.png"
OUT = HERE / "icon.ico"

# Render everything at a high master resolution, then downsample to each
# icon size so edges stay crisp.
S = 1024
SS = 4  # supersample factor for the circle/border anti-aliasing

OWL_COLOR = (55, 55, 55, 255)
BORDER_COLOR = (175, 178, 182, 255)   # neutral gray ring, reads on light & dark
WHITE = (255, 255, 255, 255)

# How much of the circle's inner diameter the owl's HEIGHT should occupy.
OWL_FILL = 0.78
BORDER_FRAC = 0.022   # border thickness as a fraction of icon size


def build_master():
    # --- crop owl tight to its ink ---
    src = Image.open(SRC).convert("L")
    bbox = ImageOps.invert(src).getbbox()
    owl_l = src.crop(bbox)                 # grayscale, dark owl on white
    owl_alpha = ImageOps.invert(owl_l)     # owl -> opaque, background -> clear

    ow, oh = owl_l.size

    # --- white circle + border on a transparent square (supersampled) ---
    big = S * SS
    canvas = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    draw = ImageDraw.Draw(canvas)

    margin = max(1, int(big * 0.012))      # tiny gap so the ring isn't clipped
    border_w = max(1, int(big * BORDER_FRAC))
    outer = (margin, margin, big - margin, big - margin)
    inner = (margin + border_w, margin + border_w,
             big - margin - border_w, big - margin - border_w)
    draw.ellipse(outer, fill=BORDER_COLOR)  # ring
    draw.ellipse(inner, fill=WHITE)         # white face

    canvas = canvas.resize((S, S), Image.LANCZOS)

    # --- size & place the owl inside the white face ---
    inner_d = (S - 2 * (S * margin // big) - 2 * (S * border_w // big))
    inner_d = S * (inner[2] - inner[0]) // big
    target_h = int(inner_d * OWL_FILL)
    scale = target_h / oh
    target_w = max(1, int(ow * scale))
    target_h = max(1, int(oh * scale))

    color_layer = Image.new("RGBA", (target_w, target_h), OWL_COLOR)
    mask = owl_alpha.resize((target_w, target_h), Image.LANCZOS)

    px = (S - target_w) // 2
    py = (S - target_h) // 2
    canvas.paste(color_layer, (px, py), mask)
    return canvas


def main():
    master = build_master()
    master.save(HERE / "icon_preview.png")  # for eyeballing the result
    sizes = [16, 32, 48, 64, 128, 256]
    frames = [master.resize((n, n), Image.LANCZOS) for n in sizes]
    frames[-1].save(OUT, format="ICO",
                    sizes=[(n, n) for n in sizes],
                    append_images=frames[:-1])
    print("wrote", OUT)


if __name__ == "__main__":
    main()
