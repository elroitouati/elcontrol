# -*- coding: utf-8 -*-
"""בונה את כל העבודות (או חלק): python3 build_all.py [id-prefix ...]"""
import os, sys, subprocess, glob
from PIL import Image
from songs import SONGS, STYLES
from layouts import Builder
from names import NAMES, SINGULAR

HERE = os.path.dirname(os.path.abspath(__file__))
SK = "/root/.claude/skills/synced/08c638a7-ddb8-46c5-be86-9b61f06d3064_3c3b1d69-4ee7-425a-92f6-908dec8fa724/pptx"


def render(folder, name, dpi=80):
    pptx = os.path.join(folder, name + ".pptx")
    subprocess.run(["python3", f"{SK}/scripts/office/soffice.py", "--headless", "--convert-to", "pdf", "--outdir", folder, pptx],
                   check=True, capture_output=True)
    prev = os.path.join(folder, "previews")
    os.makedirs(prev, exist_ok=True)
    for f in glob.glob(os.path.join(prev, "*.png")):
        os.remove(f)
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), os.path.join(folder, name + ".pdf"), os.path.join(prev, "slide")], check=True)
    fs = sorted(glob.glob(os.path.join(prev, "slide-*.png")))
    ims = [Image.open(f) for f in fs]
    w, h = ims[0].size
    cols = 3
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * w + (cols + 1) * 12, rows * h + (rows + 1) * 12), (230, 230, 230))
    for i, im in enumerate(ims):
        sheet.paste(im, (12 + (i % cols) * (w + 12), 12 + (i // cols) * (h + 12)))
    sheet.save(os.path.join(prev, "overview.jpg"), quality=85)
    return fs


if __name__ == "__main__":
    want = sys.argv[1:]
    LYR = {}
    if os.environ.get("LYRICS_DIR"):
        sys.path.insert(0, os.environ["LYRICS_DIR"])
        from lyrics_data import LYRICS as LYR
    OUT = os.environ.get("OUT_DIR")
    for song in SONGS:
        if want and not any(song["id"].startswith(w) for w in want):
            continue
        folder = os.path.join(OUT or HERE, song["id"])
        os.makedirs(folder, exist_ok=True)
        out = os.path.join(folder, song["id"] + ".pptx")
        song = dict(song, names=None if os.environ.get("PUBLIC") else NAMES.get(song["id"][:2]),
                    lyrics=LYR.get(song["id"][:2]))
        if song["id"][:2] in SINGULAR:
            song.update(SINGULAR[song["id"][:2]], singular=True)
        Builder(song, STYLES[song["style"]]).build(out)
        r = subprocess.run(["python3", f"{SK}/scripts/office/validate.py", out], capture_output=True, text=True)
        print(song["id"], "|", r.stdout.strip().splitlines()[-1] if r.stdout.strip() else r.stderr[-300:])
        render(folder, song["id"])
