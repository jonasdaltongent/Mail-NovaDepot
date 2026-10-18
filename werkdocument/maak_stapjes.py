#!/usr/bin/env python3
"""
maak_stapjes.py — het beeld bij het stapje "Typ het codewoord" (assets/stapjes/s5-codewoord.png)

Gebruik:  python3 maak_stapjes.py <leeg.pdf>     (vereist PyMuPDF en Pillow)

  leeg.pdf    CO1_De-zakelijke-e-mail zoals Google Documenten het toont (Drive: Bestand › Downloaden ›
              PDF-document, of de Drive API files.export).

Sinds 15-10-2026 begint *Nalezen* met het codewoord van de zelftest (_afspraken/lespagina.md, De laatste stap).
Zelfde werkwijze als MWW les 3: een uitsnede op ware grootte met één rood kader.
"""
import os
import sys

import fitz
from PIL import Image, ImageDraw

HIER = os.path.dirname(os.path.abspath(__file__))
UIT = os.path.join(HIER, "..", "assets", "stapjes")
SCHAAL = 2 * 96 / 72          # eerst op dubbele grootte tekenen, dan verkleinen: scherpe letters
ROOD = (197, 34, 31)
VAK = (0xF3, 0xF5, 0xFC)      # HEX_VAK in huisstijl.py
KADER = (0xEE, 0xF0, 0xF9)    # HEX_KADER
KOP = (0xE4, 0xE8, 0xF6)      # HEX_KOP: kopregel van een tabel
VENSTERKOP = (0xF2, 0xF6, 0xFC)


def kleur(fill):
    return tuple(round(c * 255) for c in fill) if fill else None


def vlakken(page, rgb):
    """De gevulde rechthoeken met deze kleur, van boven naar onder en van links naar rechts."""
    uit = []
    for d in page.get_drawings():
        k = kleur(d.get("fill"))
        if k and all(abs(a - b) <= 2 for a, b in zip(k, rgb)):
            uit.append(fitz.Rect(d["rect"]))
    uit.sort(key=lambda r: (round(r.y0), r.x0))
    return uit


def tekst(page, s, nr=0):
    treffers = page.search_for(s)
    if len(treffers) <= nr:
        raise SystemExit(f"'{s}' niet gevonden op bladzijde {page.number + 1}")
    return treffers[nr]


def bladzijde_met(doc, s):
    for p in doc:
        if p.search_for(s):
            return p
    raise SystemExit(f"'{s}' staat in geen enkele bladzijde")


def bewaar(page, naam, uitsnede, kaders=(), boven=None, onder=None):
    """Snijdt uitsnede (in pdf-punten) uit, tekent de rode kaders en bewaart op ware grootte.
    boven/onder: wat boven of onder die hoogte valt, wordt wit (geen halve regels van de buurrij)."""
    uitsnede = fitz.Rect(uitsnede) & page.rect
    pix = page.get_pixmap(matrix=fitz.Matrix(SCHAAL, SCHAAL), clip=uitsnede, alpha=False)
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    d = ImageDraw.Draw(im)
    if boven is not None:
        d.rectangle([0, 0, im.width, (boven - uitsnede.y0) * SCHAAL], fill="white")
    if onder is not None:
        d.rectangle([0, (onder - uitsnede.y0) * SCHAAL, im.width, im.height], fill="white")
    for k in kaders:
        k = fitz.Rect(k)
        box = [(k.x0 - uitsnede.x0) * SCHAAL - 6, (k.y0 - uitsnede.y0) * SCHAAL - 6,
               (k.x1 - uitsnede.x0) * SCHAAL + 6, (k.y1 - uitsnede.y0) * SCHAAL + 6]
        wit = [box[0] - 4, box[1] - 4, box[2] + 4, box[3] + 4]
        d.rounded_rectangle(wit, radius=16, outline=(255, 255, 255), width=4)
        d.rounded_rectangle(box, radius=12, outline=ROOD, width=5)
    im = im.resize((im.width // 2, im.height // 2), Image.LANCZOS)
    im.save(os.path.join(UIT, naam + ".png"), optimize=True)
    print("gemaakt:", naam + ".png", im.size)


def rand(r, x=8, y=8):
    return fitz.Rect(r.x0 - x, r.y0 - y, r.x1 + x, r.y1 + y)


def band(r, links, rechts):
    """Dezelfde hoogte als r, over de volle breedte van links tot rechts."""
    return fitz.Rect(links, r.y0, rechts, r.y1)


def main(leeg):
    os.makedirs(UIT, exist_ok=True)
    doc = fitz.open(leeg)
    p = bladzijde_met(doc, "Het codewoord")
    kop, code = tekst(p, "Nalezen"), tekst(p, "Het codewoord")
    vak = [v for v in vlakken(p, VAK) if v.y0 > code.y1][0]
    koppen = [k for k in vlakken(p, KOP) if k.y0 > vak.y1]
    onder = koppen[0].y1 + 30 if koppen else vak.y1             # nog een stukje van de controletabel eronder
    bewaar(p, "s5-codewoord", rand(fitz.Rect(vak.x0, kop.y0, vak.x1, onder), 8, 10),
           [fitz.Rect(vak.x0, code.y0 - 2, vak.x1, vak.y1)], onder=onder)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    main(sys.argv[1])
