"""Generate placeholder iLA icons (swap for real brand icons later)."""
from PIL import Image, ImageDraw, ImageFont
FONT = "/usr/share/fonts/truetype/google-fonts/Poppins-Bold.ttf"
ACCENT = (15, 82, 186)

def colour_icon(size, path):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=size // 5, fill=ACCENT)
    font = ImageFont.truetype(FONT, int(size * 0.42))
    d.text((size / 2, size / 2), "iLA", font=font, fill="white", anchor="mm")
    img.save(path)

def outline_icon(path):
    # 32x32, white on transparent (Teams/M365 requirement)
    img = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([1, 1, 30, 30], radius=6, outline="white", width=2)
    d.text((16, 16), "iLA", font=ImageFont.truetype(FONT, 12), fill="white", anchor="mm")
    img.save(path)

for s in (16, 32, 64, 80, 128):
    colour_icon(s, f"web/assets/icon-{s}.png")
colour_icon(192, "package/color.png")
outline_icon("package/outline.png")
print("icons done")
