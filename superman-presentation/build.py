# -*- coding: utf-8 -*-
"""בונה את מצגת "סופרמן" – הערכה חלופית בספרות, שכבה ט'.
הרצה: python3 build.py  →  superman.pptx (+ רקעים ב-assets/bg)
"""
import copy, math, os, random
from PIL import Image, ImageDraw, ImageFilter, ImageEnhance
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
BG_DIR = os.path.join(ASSETS, "bg")
os.makedirs(BG_DIR, exist_ok=True)

# ---------- palette ----------
BG = (11, 14, 22)          # כחול-לילה עמוק
PANEL = RGBColor(0x16, 0x1C, 0x2B)
GOLD = RGBColor(0xD1, 0xA9, 0x5E)   # זהב מעומעם
WHITE = RGBColor(0xF3, 0xEF, 0xE7)
MUTED = RGBColor(0xA4, 0xAC, 0xBD)
FONT = "Heebo"
FONT_BLACK = "Heebo Black"
FONT_LIGHT = "Heebo Light"

W, H = 13.333, 7.5
TOTAL = 11


# ---------- backgrounds: אור שבוקע מחושך ----------
def make_bg(name, light, photo=None, photo_box=None, photo_dark=0.55, blur=0):
    """light: 0..1 – עוצמת האור. המצגת מתבהרת בהדרגה, כמו מסע השיר."""
    PW, PH = 1920, 1080
    img = Image.new("RGB", (PW, PH), BG)
    if photo:
        ph = Image.open(photo).convert("RGB")
        bx, by, bw, bh = photo_box
        r = max(bw / ph.width, bh / ph.height)
        ph = ph.resize((int(ph.width * r) + 1, int(ph.height * r) + 1), Image.LANCZOS)
        ph = ph.crop(((ph.width - bw) // 2, (ph.height - bh) // 2, (ph.width - bw) // 2 + bw, (ph.height - bh) // 2 + bh))
        if blur:
            ph = ph.filter(ImageFilter.GaussianBlur(blur))
        ph = ImageEnhance.Brightness(ph).enhance(photo_dark)
        # מסכת דהייה רכה לכל הכיוונים
        mask = Image.new("L", (bw, bh), 0)
        md = ImageDraw.Draw(mask)
        steps = 120
        for i in range(steps):
            a = int(255 * (i / steps) ** 1.6)
            md.rectangle([i * bw // (2 * steps), i * bh // (2 * steps), bw - i * bw // (2 * steps), bh - i * bh // (2 * steps)], fill=a)
        mask = mask.filter(ImageFilter.GaussianBlur(40))
        img.paste(ph, (bx, by), mask)

    glow = Image.new("RGB", (PW, PH), (0, 0, 0))
    g = ImageDraw.Draw(glow)
    # הילה חמה מהפינה הימנית-עליונה
    cx, cy = PW + 80, -160
    for r in range(1500, 0, -12):
        t = 1 - r / 1500
        k = (t ** 2.4) * light
        g.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(int(205 * k), int(160 * k), int(95 * k)))
    # קרן אור אלכסונית – הגלימה/האור שחודר
    beam = Image.new("L", (PW, PH), 0)
    b = ImageDraw.Draw(beam)
    b.polygon([(PW, -40), (PW - 230, -40), (PW * 0.18, PH + 40), (PW * 0.48, PH + 40)], fill=int(34 * light + 6))
    beam = beam.filter(ImageFilter.GaussianBlur(90))
    glow = Image.composite(Image.new("RGB", (PW, PH), (215, 175, 110)), glow, beam) if False else glow
    img = Image.blend(img, Image.new("RGB", (PW, PH), (0, 0, 0)), 0)
    from PIL import ImageChops
    img = ImageChops.add(img, glow)
    beam_rgb = Image.merge("RGB", [beam.point(lambda v: int(v * 0.85)), beam.point(lambda v: int(v * 0.68)), beam.point(lambda v: int(v * 0.45))])
    img = ImageChops.add(img, beam_rgb)
    # קו גלימה עדין – קשת רחבה בתחתית
    cape = Image.new("L", (PW, PH), 0)
    c = ImageDraw.Draw(cape)
    pts = []
    for i in range(0, 101):
        x = PW * i / 100
        y = PH * 0.94 - math.sin(i / 100 * math.pi * 0.9 + 0.2) * PH * 0.16 - (i / 100) * PH * 0.08
        pts.append((x, y))
    c.line(pts, fill=int(70 * (0.5 + light / 2)), width=3)
    cape = cape.filter(ImageFilter.GaussianBlur(1.2))
    cape_rgb = Image.merge("RGB", [cape.point(lambda v: int(v * 0.82)), cape.point(lambda v: int(v * 0.66)), cape.point(lambda v: int(v * 0.37))])
    img = ImageChops.add(img, cape_rgb)
    # גרעין עדין נגד פסים
    random.seed(7)
    noise = Image.effect_noise((PW, PH), 6).convert("RGB")
    img = ImageChops.add(img, ImageEnhance.Brightness(noise).enhance(0.05))
    out = os.path.join(BG_DIR, name + ".png")
    img.save(out, optimize=True)
    return out


# ---------- text helpers (RTL מלא) ----------
def _rpr(run, size, color, bold=False, font=FONT, italic=False):
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", "he-IL")
    rPr.set("altLang", "en-US")
    rPr.set("sz", str(int(size * 100)))
    rPr.set("b", "1" if bold else "0")
    if italic:
        rPr.set("i", "1")
    for tag in ("a:solidFill", "a:latin", "a:ea", "a:cs"):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    fill = etree.SubElement(rPr, qn("a:solidFill"))
    clr = etree.SubElement(fill, qn("a:srgbClr"))
    clr.set("val", str(color))
    for tag in ("a:latin", "a:cs"):
        el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)


def add_text(slide, x, y, w, h, paras, anchor="t", name=None):
    """paras: [ {runs:[(text, opts)], align:'r', space_after:pt, line:1.0} ]"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    bp = tf._txBody.find(qn("a:bodyPr"))
    for k in ("lIns", "rIns", "tIns", "bIns"):
        bp.set(k, "0")
    bp.set("rtlCol", "1")
    bp.set("anchor", {"t": "t", "m": "ctr", "b": "b"}[anchor])
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        pPr = para._p.get_or_add_pPr()
        pPr.set("algn", p.get("align", "r"))
        pPr.set("rtl", "0" if p.get("ltr") else "1")
        if p.get("line"):
            ls = etree.SubElement(pPr, qn("a:lnSpc"))
            etree.SubElement(ls, qn("a:spcPct")).set("val", str(int(p["line"] * 100000)))
        if p.get("space_after") is not None:
            sa = etree.SubElement(pPr, qn("a:spcAft"))
            etree.SubElement(sa, qn("a:spcPts")).set("val", str(int(p["space_after"] * 100)))
        etree.SubElement(pPr, qn("a:buNone"))
        for text, o in p["runs"]:
            r = para.add_run()
            r.text = text
            _rpr(r, o.get("size", 22), o.get("color", WHITE), o.get("bold", False), o.get("font", FONT), o.get("italic", False))
        # end-paragraph props so empty lines keep size
        endp = etree.SubElement(para._p, qn("a:endParaRPr"))
        endp.set("lang", "he-IL")
        endp.set("sz", str(int(p["runs"][0][1].get("size", 22) * 100)) if p["runs"] else "2200")
    return tb


def P(*runs, align="r", space_after=None, line=None, ltr=False):
    return {"runs": list(runs), "align": align, "space_after": space_after, "line": line, "ltr": ltr}


def R(text, size=22, color=WHITE, bold=False, font=FONT, italic=False):
    return (text, dict(size=size, color=color, bold=bold, font=font, italic=italic))


def rect(slide, x, y, w, h, fill=None, line=None, radius=None, line_w=1.0, name=None, transparency=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    if radius:
        s.adjustments[0] = radius
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
        if transparency:
            sf = s.fill._xPr.find(qn("a:solidFill"))
            clr = sf.find(qn("a:srgbClr"))
            etree.SubElement(clr, qn("a:alpha")).set("val", str(int((100 - transparency) * 1000)))
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    s.text_frame.text = ""
    return s


def oval(slide, x, y, d, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background(); s.shadow.inherit = False
    return s


def picture_fill(slide, path, x, y, w, h, name=None):
    im = Image.open(path)
    iw, ih = im.size
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    box_r, img_r = w / h, iw / ih
    if img_r > box_r:  # תמונה רחבה – חותכים צדדים
        c = (1 - box_r / img_r) / 2
        pic.crop_left = pic.crop_right = c
    else:
        c = (1 - img_r / box_r)
        pic.crop_top = c * 0.25
        pic.crop_bottom = c * 0.75
    if name:
        pic.name = name
    return pic


def set_bg(slide, path):
    pic = slide.shapes.add_picture(path, 0, 0, Inches(W), Inches(H))
    pic.name = "Background"
    sp = slide.shapes._spTree
    sp.remove(pic._element)
    sp.insert(2, pic._element)


def add_fade(slide):
    sld = slide._element
    trans = etree.fromstring(
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="slow"><p:fade/></p:transition>')
    clr = sld.find(qn("p:clrMapOvr"))
    if clr is not None:
        clr.addnext(trans)
    else:
        sld.find(qn("p:cSld")).addnext(trans)


def chrome(slide, n, tag=None, title=None):
    """תגית 'עמוד X' בפינה + מספר שקף + כותרת"""
    if tag:
        rect(slide, 0.6, 0.5, 1.75, 0.5, line=GOLD, radius=0.5, line_w=1.25, name="Rubric tag")
        add_text(slide, 0.6, 0.5, 1.75, 0.5, [P(R(tag, 16, GOLD, True), align="ctr")], anchor="m")
    add_text(slide, 0.6, 6.92, 1.2, 0.35, [P(R(f"{n} / {TOTAL}", 12, MUTED), align="l", ltr=True)], anchor="m", name="Slide number")
    if title:
        add_text(slide, 2.7, 0.38, 10.03, 0.85, [P(R(title, 44, WHITE, True))], anchor="m", name="Title")
    add_fade(slide)


# ================= content =================
LYRICS = [
    ["שראתה בתוכי כוח על", "שהוא גם קביים גם כנפיים", "שלימדה אותי שבשבר מר", "מסתתרים להם גם שמיים",
     "ולא חשוב כמה חטפתי", "העיקר לא הרמתי ידיים"],
    ["ולאט שיר ועוד שיר", "למדתי לעצמי להסביר", "שכל מה שחוויתי", "נועד לי", "כדי שאוכל לבסוף להבין"],
    ["שרק מי שצעד באפלה ובחושך", "יודע באמת מהו אור", "שכל תפילה בה ביקשתי לשכוח", "רק גרמה לי לשוב ולזכור"],
    ["ואחרי שנפצעתי", "היו לי לילות", "ששאגתי מתוך החזה", "התחלתי לנהל עם אלוהים חשבונות", "אם תיתן לי ככה",
     "אני אתן לך את זה"],
    ["אך בסתר ליבי השבור השלם", "לא היו לי ספקות", "כבר לא מפחיד אותי למות", "יותר מפחיד לא לחיות"],
    ["נזרקתי לחיים ידיי היו ריקות", "שרדתי כל אוויר דליל", "מחנק בעליות", "אחרי החושך שעברתי", "אלף תהומות",
     "כבר לא מפחיד אותי למות", "יותר מפחיד לא לחיות"],
    ["זה לכל מי שחשד בי שאני אחר", "שקיבלתי את האור שלי מן ההפקר", "שהמלאכים ירדו", "השתחוו, אמרו ׳אמן",
     "הוא לא אדם רגיל", "הוא סופר-מן׳"],
]
CHORUS = {"כבר לא מפחיד אותי למות", "יותר מפחיד לא לחיות", "הוא סופר-מן׳"}


def stanza_paras(stanzas, size=22):
    out = []
    for si, st in enumerate(stanzas):
        for li, line in enumerate(st):
            col = GOLD if line in CHORUS else WHITE
            out.append(P(R(line, size, col, line in CHORUS), line=1.02,
                         space_after=14 if (li == len(st) - 1 and si < len(stanzas) - 1) else 0))
    return out


def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    blank = prs.slide_layouts[6]
    sec = []

    bgs = {}
    lv = [0.30, 0.34, 0.38, 0.42, 0.42, 0.48, 0.48, 0.54, 0.58, 0.66, 1.0, 0.6]
    for i, l in enumerate(lv):
        if i == 10:
            bgs[i] = make_bg(f"bg{i:02d}", l, photo=os.path.join(ASSETS, "amedi_stage.jpg"),
                             photo_box=(0, 0, 1920, 1080), photo_dark=0.55, blur=14)
        else:
            bgs[i] = make_bg(f"bg{i:02d}", l)

    def new(i):
        s = prs.slides.add_slide(blank)
        set_bg(s, bgs[i])
        return s

    # ---- 0. שער ----
    s = new(0)
    add_text(s, 0.6, 0.55, 12.13, 0.5, [P(R("הערכה חלופית בספרות – שכבה ט׳", 20, GOLD, True))])
    add_text(s, 0.6, 1.55, 12.13, 2.2, [P(R("סופרמן", 130, WHITE, font=FONT_BLACK), line=0.9)], anchor="m", name="Title")
    add_text(s, 0.6, 3.75, 12.13, 0.8, [P(R("עידן עמדי", 40, GOLD, False, FONT_LIGHT))])
    add_text(s, 0.6, 4.55, 12.13, 0.6, [P(R("שיר על פציעה, שיקום ובחירה בחיים", 24, MUTED))])
    add_text(s, 0.6, 5.85, 12.13, 0.5, [P(R("מגישים", 16, MUTED))])
    add_text(s, 0.6, 6.2, 12.13, 0.6, [P(R("אלרואי טואטי  ·  יואב טימסטיט  ·  אלינדב בק", 24, WHITE, True))])
    add_fade(s)

    # ---- 1. חלק א' ----
    s = new(1)
    chrome(s, 2, "חלק א׳", "הקטע שבחרנו")
    # כרטיס הציטוט – מימין
    rect(s, 7.05, 1.6, 5.68, 4.95, fill=PANEL, radius=0.06, name="Quote card")
    add_text(s, 11.55, 1.45, 1.0, 1.2, [P(R("״", 96, GOLD, font=FONT_BLACK))], name="Quote mark")
    add_text(s, 7.45, 2.35, 4.9, 3.6, [
        P(R("אך בסתר ליבי ", 27, WHITE, True), R("השבור השלם", 27, GOLD, True), space_after=6),
        P(R("לא היו לי ספקות", 27, WHITE, True), space_after=18),
        P(R("כבר לא מפחיד אותי ", 27, WHITE, True), R("למות", 27, GOLD, True), space_after=6),
        P(R("יותר מפחיד לא ", 27, WHITE, True), R("לחיות", 27, GOLD, True)),
    ], anchor="m")
    add_text(s, 7.45, 5.95, 4.9, 0.4, [P(R("עידן עמדי, ״סופרמן״", 16, MUTED))])
    # למה + חיבור אישי – משמאל
    add_text(s, 0.6, 1.6, 6.05, 2.4, [
        P(R("למה בחרנו בקטע", 24, GOLD, True), space_after=6),
        P(R("זה רגע השיא של השיר. אחרי הפציעה, הדובר מגלה שהפחד הגדול הוא לא מהמוות – אלא מלוותר על החיים. הצירוף ״השבור השלם״ מראה שאפשר להיות פגוע ועדיין שלם.", 22, WHITE), line=1.05),
    ])
    add_text(s, 0.6, 4.2, 6.05, 2.4, [
        P(R("החיבור האישי שלנו", 24, GOLD, True), space_after=6),
        P(R("גדלנו בתוך המלחמה הזאת, וראינו סביבנו אנשים שיצאו למילואים וחזרו אחרים. הקטע לימד אותנו שגבורה היא לא להיות בלתי שביר – אלא לבחור לחיות גם אחרי שנשברים.", 22, WHITE), line=1.05),
    ])

    # ---- 2. עמוד 1 – האמן ----
    s = new(2)
    chrome(s, 3, "עמוד 1", "האמן: עידן עמדי")
    rect(s, 0.78, 1.78, 4.1, 4.7, line=GOLD, line_w=1.0, name="Photo frame")
    picture_fill(s, os.path.join(ASSETS, "amedi_portrait_seli_ben_arie_ccbysa4.jpg"), 0.6, 1.6, 4.1, 4.7, name="Artist photo")
    add_text(s, 0.6, 6.6, 4.3, 0.3, [P(R("צילום: סלי בן אריה, 2020 · CC BY-SA 4.0", 11, MUTED), align="l")])
    bio = [
        [R("נולד ב-1988 בירושלים", 22, GOLD, True), R(" וגדל בשכונת פסגת זאב.", 22)],
        [R("שירת כלוחם ומפקד", 22, GOLD, True), R(" בחיל ההנדסה הקרבית, ובמילואים ביחידת דובדבן.", 22)],
        [R("ב״כוכב נולד״ 2010", 22, GOLD, True), R(" הגיע לגמר עם ״כאב של לוחמים״ – שיר שכתב על חוויותיו כלוחם.", 22)],
        [R("שחקן ויוצר:", 22, GOLD, True), R(" גילם את שגיא צור בסדרה ״פאודה״ והוציא שישה אלבומי אולפן.", 22)],
        [R("עשייה ציבורית:", 22, GOLD, True), R(" נפצע קשה בעזה במילואים, שיתף את הציבור במסע השיקום – ובסוף 2025 חזר לשרת.", 22)],
    ]
    y = 1.6
    for row in bio:
        oval(s, 12.55, y + 0.17, 0.16, GOLD)
        add_text(s, 5.35, y, 7.0, 0.95, [P(*row, line=1.0)])
        y += 1.0

    # ---- 3+4. עמוד 2 – מילות השיר ----
    for k, (right, left) in enumerate([(LYRICS[0:2], LYRICS[2:4]), (LYRICS[4:6], LYRICS[6:7])]):
        s = new(3 + k)
        chrome(s, 4 + k, "עמוד 2", "מילות השיר" + (" – חלק א׳" if k == 0 else " – חלק ב׳"))
        add_text(s, 6.95, 1.55, 5.78, 5.3, stanza_paras(right), name="Lyrics right")
        add_text(s, 0.6, 1.55, 5.9, 5.3, stanza_paras(left), name="Lyrics left")
        add_text(s, 0.6, 6.55 if k else 6.55, 12.13, 0.35, [P(R("מילים: עידן עמדי · מקור הטקסט: LyricFind", 12, MUTED), align="l")])

    # ---- 5+6. עמוד 3 – מושגים ----
    concepts = [
        dict(name="מטאפורה", n=1, ord="ראשון",
             definition="חיבור בין שני מושגים מעולמות שונים, בלי מילת דימוי (כמו, כְּ־), כדי ליצור משמעות חדשה.",
             quote=[[R("״שראתה בתוכי כוח על", 30, WHITE, True)],
                    [R("שהוא גם ", 30, WHITE, True), R("קביים", 30, GOLD, True), R(" גם ", 30, WHITE, True), R("כנפיים", 30, GOLD, True), R("״", 30, WHITE, True)]],
             how="הכוח שאשתו ראתה בו הוא ״קביים״ – תמיכה למי שנפצע ולא יכול ללכת לבד – וגם ״כנפיים״ – מה שמרים ומאפשר לעוף.",
             why="שיקום הוא לא רק לעמוד שוב על הרגליים, אלא גם לחלום ולהמריא מחדש. האהבה נותנת לו את שניהם."),
        dict(name="ניגוד", n=2, ord="שני",
             definition="העמדה של שני דברים הפוכים זה מול זה, כדי לחדד את ההבדל ביניהם ולהבליט רעיון.",
             quote=[[R("״שרק מי שצעד ב", 30, WHITE, True), R("אפלה", 30, GOLD, True), R(" וב", 30, WHITE, True), R("חושך", 30, GOLD, True)],
                    [R("יודע באמת מהו ", 30, WHITE, True), R("אור", 30, GOLD, True), R("״", 30, WHITE, True)]],
             how="החושך הוא הפציעה והלילות הקשים שאחריה; האור הוא החזרה לחיים. הניגוד חוזר גם בפזמון: ״למות״ מול ״לחיות״.",
             why="רק מי שעבר את הרע מעריך באמת את הטוב. זה המסע של כל השיר – מחושך לאור."),
    ]
    for k, c in enumerate(concepts):
        s = new(5 + k)
        chrome(s, 6 + k, "עמוד 3", f"מושג {c['ord']}: {c['name']}")
        # עמודה ימנית: שם + הגדרה
        add_text(s, 8.75, 1.6, 3.98, 1.0, [P(R(c["name"], 50, GOLD, font=FONT_BLACK))], anchor="m")
        add_text(s, 8.75, 2.65, 3.98, 1.6, [P(R("הגדרה: ", 22, MUTED, True), R(c["definition"], 22, WHITE), line=1.05)])
        # כרטיס ציטוט
        rect(s, 0.6, 1.6, 7.75, 2.65, fill=PANEL, radius=0.06, name="Quote card")
        add_text(s, 0.95, 1.6, 7.1, 0.6, [P(R("הציטוט מהשיר", 16, MUTED, True))], anchor="b")
        add_text(s, 0.95, 2.2, 7.1, 1.85, [P(*q, line=1.05, space_after=4) for q in c["quote"]], anchor="m")
        # שתי עמודות הסבר
        rect(s, 6.82, 4.6, 5.91, 2.15, fill=PANEL, radius=0.06, transparency=40)
        rect(s, 0.6, 4.6, 5.91, 2.15, fill=PANEL, radius=0.06, transparency=40)
        add_text(s, 7.1, 4.78, 5.35, 1.85, [P(R("איך המושג בא לידי ביטוי?", 22, GOLD, True), space_after=4),
                                           P(R(c["how"], 22, WHITE), line=1.02)])
        add_text(s, 0.88, 4.78, 5.35, 1.85, [P(R("מה זה תורם למשמעות?", 22, GOLD, True), space_after=4),
                                            P(R(c["why"], 22, WHITE), line=1.02)])

    # ---- 7. עמוד 4א – מילים ----
    s = new(7)
    chrome(s, 8, "עמוד 4א", "מילים שמעוררות את המלחמה")
    add_text(s, 0.6, 1.45, 12.13, 0.6, [P(R("מילים וצירופים מהשיר שמחזירים אותנו אל שדה הקרב ואל הפציעה", 22, MUTED))])
    words = ["נפצעתי", "חטפתי", "לא הרמתי ידיים", "קביים",
             "שאגתי מתוך החזה", "אוויר דליל", "מחנק בעליות", "החושך שעברתי"]
    cw, ch, gap = 2.8, 1.95, 0.31
    for i, wd in enumerate(words):
        col, row = i % 4, i // 4
        x = W - 0.6 - cw - col * (cw + gap)  # מימין לשמאל
        yy = 2.3 + row * (ch + 0.3)
        rect(s, x, yy, cw, ch, fill=PANEL, radius=0.08, name=f"Word card {i + 1}")
        add_text(s, x + cw - 0.8, yy + 0.18, 0.6, 0.4, [P(R(f"{i + 1:02d}", 14, GOLD, True), ltr=True)])
        add_text(s, x + 0.15, yy + 0.35, cw - 0.3, ch - 0.45, [P(R(f"״{wd}״", 28, WHITE, True), align="ctr", line=0.95)], anchor="m")

    # ---- 8. עמוד 4ב – אירועים ----
    s = new(8)
    chrome(s, 9, "עמוד 4ב", "האירועים שהמילים מעוררות")
    picture_fill(s, os.path.join(ASSETS, "amedi_handshake.jpg"), 0.6, 1.6, 5.2, 3.75, name="Event photo")
    add_text(s, 0.6, 5.5, 5.2, 1.3, [P(R("עידן עמדי במדי מילואים, בלחיצת יד עם מפקד בכיר. מאחורי ה״סופרמן״ שבשיר עומד לוחם אמיתי.", 18, MUTED), line=1.0)])
    events = [
        [R("7 באוקטובר 2023:", 22, GOLD, True), R(" פורצת מלחמת ״חרבות ברזל״, ועמדי מגויס למילואים.", 22)],
        [R("לוחם הנדסה:", 22, GOLD, True), R(" נלחם עם יחידתו, גם בתוך רצועת עזה.", 22)],
        [R("8 בינואר 2024:", 22, GOLD, True), R(" נפצע קשה מפיצוץ בזמן פעולה להשמדת מנהרה.", 22)],
        [R("בית החולים שיבא:", 22, GOLD, True), R(" מגיע מורדם ומונשם, עם רסיסים בכל הגוף.", 22)],
        [R("השיקום:", 22, GOLD, True), R(" יותר משנה של החלמה – ומתוכה נולד ״סופרמן״.", 22)],
    ]
    # ציר זמן אנכי
    rect(s, 12.62, 1.85, 0.025, 4.2, fill=GOLD, transparency=60, name="Timeline")
    y = 1.6
    for row in events:
        oval(s, 12.53, y + 0.2, 0.2, GOLD)
        add_text(s, 6.35, y, 6.0, 0.95, [P(*row, line=1.0)])
        y += 1.0

    # ---- 9. סיכום ----
    s = new(10)
    add_text(s, 0.6, 0.55, 12.13, 0.5, [P(R("סיכום", 20, GOLD, True), align="ctr")])
    add_text(s, 0.6, 1.7, 12.13, 3.6, [
        P(R("ה״סופרמן״ של עידן עמדי לא עף, ולא בלתי פגיע.", 36, WHITE, True), align="ctr", space_after=18, line=1.05),
        P(R("הוא נפצע, נשבר – ובחר לקום ולחיות.", 40, GOLD, True), align="ctr", space_after=18, line=1.05),
        P(R("זו הגבורה האמיתית.", 40, WHITE, False, FONT_LIGHT), align="ctr", line=1.05),
    ], anchor="m")
    add_text(s, 0.6, 6.92, 1.2, 0.35, [P(R(f"10 / {TOTAL}", 12, MUTED), align="l", ltr=True)], anchor="m")
    add_fade(s)

    # ---- 10. מקורות ----
    s = new(11)
    chrome(s, 11, None, "מקורות")
    src = [
        ("ויקיפדיה העברית, ״עידן עמדי״", "he.wikipedia.org/wiki/עידן_עמדי"),
        ("Wikipedia, ״Idan Amedi״", "en.wikipedia.org/wiki/Idan_Amedi"),
        ("ynet, דיווח על פציעתו של עידן עמדי ברצועת עזה, 8.1.2024", "ynet.co.il/news/article/bjhwsjfut"),
        ("ice, כתבה על יציאת הסינגל ״סופרמן״, 16.12.2024", "ice.co.il/culture/news/article/1042515"),
        ("כאן חדשות, כתבה על ״כאב של לוחמים״", "kan.org.il/content/kan-news/culture/680628"),
        ("מילות השיר: עידן עמדי, ״סופרמן״ (2024), באמצעות LyricFind", ""),
        ("הגדרות המושגים: דף לימוד בספרות, אתר דעת", "daat.ac.il"),
        ("תמונת האמן: סלי בן אריה, ויקישיתוף, CC BY-SA 4.0", "commons.wikimedia.org"),
    ]
    paras = []
    for t, u in src:
        runs = [R(t, 18, WHITE, True)]
        if u:
            runs.append(R("  ·  ", 18, MUTED))
            runs.append(R(u, 16, MUTED))
        paras.append(P(*runs, space_after=10, line=1.0))
    add_text(s, 0.6, 1.6, 12.13, 5.2, paras)

    out = os.path.join(HERE, "superman.pptx")
    prs.save(out)
    print("saved", out)


if __name__ == "__main__":
    build()
