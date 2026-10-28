"""Generate the add-in icons (16/32/80 px PNGs). Run once: python3 tools/make_icons.py"""

import os
from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join(os.path.dirname(__file__), "..", "assets")
FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
BG = (15, 108, 189, 255)      # Office blue
FG = (255, 255, 255, 255)


def make(size: int) -> Image.Image:
    scale = 4
    img = Image.new("RGBA", (size * scale, size * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    radius = int(size * scale * 0.22)
    draw.rounded_rectangle([0, 0, size * scale - 1, size * scale - 1], radius=radius, fill=BG)

    label = "OC" if size >= 32 else "O"
    try:
        font = ImageFont.truetype(FONT_PATH, int(size * scale * (0.52 if len(label) == 2 else 0.72)))
    except OSError:
        font = ImageFont.load_default()
    box = draw.textbbox((0, 0), label, font=font)
    x = (size * scale - (box[2] - box[0])) / 2 - box[0]
    y = (size * scale - (box[3] - box[1])) / 2 - box[1]
    draw.text((x, y), label, font=font, fill=FG)
    return img.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for px in (16, 32, 80):
        make(px).save(os.path.join(OUT, f"icon-{px}.png"))
        print(f"wrote icon-{px}.png")
