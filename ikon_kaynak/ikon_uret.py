"""Temizlik Birimi Sunum Asistanı ikonu: ISS ovali + Vega Taurus yıldızı."""
import math
from PIL import Image, ImageDraw, ImageFilter, ImageChops

S = 1024  # çizim çözünürlüğü
NAVY = (0, 45, 86)
VEGA = {"teal": (60, 188, 203), "magenta": (228, 6, 107), "purple": (110, 45, 140), "yellow": (250, 200, 30)}
ISS = r"C:/Users/Lenovo/Desktop/ISS.jpg"
OUT = r"C:/Users/Lenovo/Downloads/Mert (Claude)/bin/temizlik-sunum"


def iss_oval():
    im = Image.open(ISS).convert("RGB")
    w, h = im.size
    # beyaz olmayan + içi doldurulmuş (harfler) = oval maskesi
    g = im.convert("L").point(lambda v: 255 if v > 200 else 0)
    ff = g.copy()
    ImageDraw.floodfill(ff, (0, 0), 128)
    mask = ff.point(lambda v: 0 if v == 128 else 255).filter(ImageFilter.GaussianBlur(0.8))
    out = im.convert("RGBA"); out.putalpha(mask)
    return out.crop(mask.getbbox())


def sparkle(d, cx, cy, r, color, rot=0, thin=0.30):
    pts = []
    for i in range(8):
        ang = math.radians(rot + i * 45)
        rr = r if i % 2 == 0 else r * thin
        pts.append((cx + rr * math.cos(ang), cy + rr * math.sin(ang)))
    d.polygon(pts, fill=color)


def rounded(size, rad):
    m = Image.new("L", (size, size), 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, size - 1, size - 1), rad, fill=255)
    return m


def main():
    pad = 40
    card = S - 2 * pad
    base = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # gölge
    sh = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle((pad, pad + 14, S - pad, S - pad + 14), 190, fill=(0, 20, 40, 110))
    base.alpha_composite(sh.filter(ImageFilter.GaussianBlur(18)))

    # beyaz kart, hafif dikey gradyan
    grad = Image.new("RGBA", (card, card))
    gd = ImageDraw.Draw(grad)
    for y in range(card):
        t = y / card
        c = tuple(int(255 - 18 * t) for _ in range(3))
        gd.line((0, y, card, y), fill=c + (255,))
    # alt lacivert bant + Vega renk şeridi
    band_h = int(card * 0.20)
    gd.rectangle((0, card - band_h, card, card), fill=NAVY + (255,))
    stripe = 18
    cols = [VEGA["teal"], VEGA["purple"], VEGA["magenta"], VEGA["yellow"]]
    seg = card / 4
    for i, c in enumerate(cols):
        gd.rectangle((int(i * seg), card - band_h - stripe, int((i + 1) * seg), card - band_h), fill=c + (255,))
    # bantta sunum (bar grafik) simgesi
    bx, by = card // 2, card - band_h // 2
    bw = 34
    for i, hgt in enumerate((50, 80, 110)):
        x0 = bx - 75 + i * 55
        gd.rounded_rectangle((x0, by + 55 - hgt, x0 + bw, by + 55), 6, fill=(255, 255, 255, 255))
    grad.putalpha(rounded(card, 190))
    base.alpha_composite(grad, (pad, pad))

    # ISS ovali
    oval = iss_oval()
    ow = int(card * 0.66)
    oh = int(oval.height * ow / oval.width)
    oval = oval.resize((ow, oh), Image.LANCZOS)
    ox = pad + int(card * 0.06)
    oy = pad + int((card - band_h - stripe - oh) / 2) + 20
    base.alpha_composite(oval, (ox, oy))

    # Vega yıldız kümesi (sağ üst, diyagonal dizilim)
    st = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    sd = ImageDraw.Draw(st)
    cx, cy = pad + card * 0.80, pad + card * 0.30
    sparkle(sd, cx + 55, cy - 80, 48, VEGA["yellow"], rot=15)
    sparkle(sd, cx + 20, cy - 18, 62, VEGA["purple"], rot=10)
    sparkle(sd, cx - 55, cy + 15, 92, VEGA["teal"], rot=-10)
    sparkle(sd, cx - 10, cy + 90, 96, VEGA["magenta"], rot=20)
    # beyaz diyagonal kesik çizgi (orijinal logodaki gibi)
    cut = Image.new("L", (S, S), 0)
    ImageDraw.Draw(cut).line((cx + 125, cy - 160, cx - 125, cy + 205), fill=255, width=13)
    a = st.getchannel("A")
    st.putalpha(ImageChops.subtract(a, cut))
    # yıldızlara beyaz kontur (oval üstünde ayrışsın)
    halo = Image.new("RGBA", (S, S), (255, 255, 255, 0))
    halo.putalpha(st.getchannel("A").filter(ImageFilter.MaxFilter(25)))
    base.alpha_composite(halo)
    base.alpha_composite(st)

    big = base
    big.resize((512, 512), Image.LANCZOS).save(OUT + "/ikon_kaynak/ikon_onizleme.png")
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    big.resize((256, 256), Image.LANCZOS).save(OUT + "/ikon.ico", sizes=sizes)
    print("ok")


if __name__ == "__main__":
    main()
