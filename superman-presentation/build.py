# -*- coding: utf-8 -*-
"""בונה את מצגת "סופרמן" – הערכה חלופית בספרות, שכבה ט'.
הרצה:  python3 build.py [night|paper|doc]   (ברירת מחדל: night)
"""
import math, os, random, sys
from PIL import Image, ImageDraw, ImageFilter, ImageEnhance, ImageChops, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
W, H = 13.333, 7.5
M = 0.7  # שוליים

STYLES = {
    # קולנועי: כחול-לילה, אור זהוב שמתחזק לאורך המצגת
    "night": dict(bg=(22, 29, 45), text="F4F0E8", muted="B6BDCC", accent="DDB46A", hair="47506A",
                  title="Heebo", display="Heebo Black", body="Heebo", light="Heebo Light", gray_photos=False),
    # ספר שירה: דף לבן, אות עברית קלאסית, אדום-גלימה
    "paper": dict(bg=(246, 245, 242), text="1C1C20", muted="5E616A", accent="9E3328", hair="CFCAC0",
                  title="Frank Ruhl Libre Black", display="Frank Ruhl Libre Black", body="Heebo", light="Frank Ruhl Libre", gray_photos=True),
    # דוקומנטרי: שחור-לבן, טיפוגרפיה גדולה, נגיעה אחת של אדום
    "doc": dict(bg=(14, 14, 15), text="F2F2F0", muted="A2A2A2", accent="D0503F", hair="3C3C3C",
                title="Heebo Black", display="Heebo Black", body="Heebo", light="Heebo Light", gray_photos=True),
}


# ---------- רקעים ----------
def _grain(img, amt):
    noise = Image.effect_noise(img.size, 8).convert("RGB")
    return ImageChops.add(img, ImageEnhance.Brightness(noise).enhance(amt))


def _photo_layer(img, photo, dark, blur, gray):
    PW, PH = img.size
    ph = Image.open(photo).convert("RGB")
    r = max(PW / ph.width, PH / ph.height)
    ph = ph.resize((int(ph.width * r) + 1, int(ph.height * r) + 1), Image.LANCZOS)
    ph = ph.crop(((ph.width - PW) // 2, (ph.height - PH) // 2, (ph.width - PW) // 2 + PW, (ph.height - PH) // 2 + PH))
    if gray:
        ph = ImageOps.grayscale(ph).convert("RGB")
    ph = ph.filter(ImageFilter.GaussianBlur(blur))
    ph = ImageEnhance.Brightness(ph).enhance(dark)
    return Image.blend(img, ph, 0.85)


def make_bg(style, name, light, photo=None):
    st = STYLES[style]
    PW, PH = 1920, 1080
    img = Image.new("RGB", (PW, PH), st["bg"])
    if photo:
        img = _photo_layer(img, photo, 0.62 if style == "night" else 0.5, 14, st["gray_photos"])
    if style in ("night", "doc"):
        glow = Image.new("RGB", (PW, PH), (0, 0, 0))
        g = ImageDraw.Draw(glow)
        cx, cy = PW + 80, -160
        tint = (215, 168, 100) if style == "night" else (200, 200, 200)
        amt = light if style == "night" else light * 0.45
        for r in range(1600, 0, -12):
            t = 1 - r / 1600
            k = (t ** 2.2) * amt
            g.ellipse([cx - r, cy - r, cx + r, cy + r], fill=tuple(int(c * k) for c in tint))
        img = ImageChops.add(img, glow)
        beam = Image.new("L", (PW, PH), 0)
        ImageDraw.Draw(beam).polygon([(PW, -40), (PW - 260, -40), (PW * 0.2, PH + 40), (PW * 0.5, PH + 40)],
                                     fill=int(30 * amt + 6))
        beam = beam.filter(ImageFilter.GaussianBlur(100))
        mul = (0.85, 0.68, 0.45) if style == "night" else (0.7, 0.7, 0.7)
        img = ImageChops.add(img, Image.merge("RGB", [beam.point(lambda v, m=m: int(v * m)) for m in mul]))
        img = _grain(img, 0.05)
    else:
        # נייר: גרעין עדין והצללה רכה בשוליים
        img = ImageChops.subtract(img, ImageEnhance.Brightness(Image.effect_noise((PW, PH), 10).convert("RGB")).enhance(0.035))
        vig = Image.new("L", (PW, PH), 0)
        ImageDraw.Draw(vig).ellipse([-PW * 0.25, -PH * 0.35, PW * 1.25, PH * 1.35], fill=255)
        vig = vig.filter(ImageFilter.GaussianBlur(220))
        img = Image.composite(img, ImageEnhance.Brightness(img).enhance(0.93), vig)
    d = os.path.join(ASSETS, "bg", style)
    os.makedirs(d, exist_ok=True)
    out = os.path.join(d, name + ".png")
    img.save(out, optimize=True)
    return out


def photo_variant(style, path):
    if not STYLES[style]["gray_photos"]:
        return path
    d = os.path.join(ASSETS, "bg", style)
    os.makedirs(d, exist_ok=True)
    out = os.path.join(d, "gray_" + os.path.basename(path))
    im = ImageOps.grayscale(Image.open(path).convert("RGB"))
    ImageEnhance.Contrast(im).enhance(1.12).save(out, quality=92)
    return out


# ---------- טקסט RTL ----------
class Deck:
    def __init__(self, style):
        s = STYLES[style]
        self.TEXT, self.MUTED, self.ACC, self.HAIR = (RGBColor.from_string(s[k]) for k in ("text", "muted", "accent", "hair"))
        self.F_TITLE, self.F_DISPLAY, self.F_BODY, self.F_LIGHT = s["title"], s["display"], s["body"], s["light"]

    def R(self, text, size=22, color=None, bold=False, font=None):
        return (text, dict(size=size, color=color or self.TEXT, bold=bold, font=font or self.F_BODY))

    @staticmethod
    def P(*runs, align="r", space_after=None, line=None, ltr=False):
        return dict(runs=list(runs), align=align, space_after=space_after, line=line, ltr=ltr)

    @staticmethod
    def _rpr(run, o):
        rPr = run._r.get_or_add_rPr()
        rPr.set("lang", "he-IL"); rPr.set("altLang", "en-US")
        rPr.set("sz", str(int(o["size"] * 100)))
        rPr.set("b", "1" if o["bold"] else "0")
        fill = etree.SubElement(rPr, qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr")).set("val", str(o["color"]))
        for tag in ("a:latin", "a:cs"):
            etree.SubElement(rPr, qn(tag)).set("typeface", o["font"])

    def text(self, slide, x, y, w, h, paras, anchor="t", name=None):
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
            pPr.set("algn", p["align"])
            pPr.set("rtl", "0" if p["ltr"] else "1")
            if p["line"]:
                ls = etree.SubElement(pPr, qn("a:lnSpc"))
                etree.SubElement(ls, qn("a:spcPct")).set("val", str(int(p["line"] * 100000)))
            if p["space_after"] is not None:
                sa = etree.SubElement(pPr, qn("a:spcAft"))
                etree.SubElement(sa, qn("a:spcPts")).set("val", str(int(p["space_after"] * 100)))
            etree.SubElement(pPr, qn("a:buNone"))
            for t, o in p["runs"]:
                r = para.add_run(); r.text = t; self._rpr(r, o)
            e = etree.SubElement(para._p, qn("a:endParaRPr"))
            e.set("lang", "he-IL"); e.set("sz", str(int(p["runs"][0][1]["size"] * 100)))
        return tb

    def hair(self, slide, x, y, w):
        s = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
        s.line.color.rgb = self.HAIR
        s.line.width = Pt(0.75)
        st = s._element.find(qn("p:style"))
        if st is not None:
            s._element.remove(st)
        return s

    def title(self, slide, t):
        self.text(slide, M, 0.5, W - 2 * M, 0.95, [self.P(self.R(t, 44, self.TEXT, False, self.F_TITLE))], anchor="m", name="Title")

    def pageno(self, slide, n, x=M):
        self.text(slide, x, 6.95, 1.0, 0.3, [self.P(self.R(str(n), 12, self.MUTED), align="l", ltr=True)], anchor="m", name="Slide number")


def picture_fill(slide, path, x, y, w, h, focus=0.3, name=None):
    iw, ih = Image.open(path).size
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    box_r, img_r = w / h, iw / ih
    if img_r > box_r:
        c = 1 - box_r / img_r
        pic.crop_left, pic.crop_right = c * 0.5, c * 0.5
    else:
        c = 1 - img_r / box_r
        pic.crop_top, pic.crop_bottom = c * focus, c * (1 - focus)
    if name:
        pic.name = name
    return pic


def set_bg(slide, path):
    pic = slide.shapes.add_picture(path, 0, 0, Inches(W), Inches(H))
    pic.name = "Background"
    tree = slide.shapes._spTree
    tree.remove(pic._element); tree.insert(2, pic._element)


def add_fade(slide):
    t = etree.fromstring('<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="slow"><p:fade/></p:transition>')
    slide._element.find(qn("p:clrMapOvr")).addnext(t)


# ================= תוכן =================
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


def build(style="night"):
    D = Deck(style)
    R, P = D.R, D.P
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    blank = prs.slide_layouts[6]
    levels = [0.42, 0.46, 0.50, 0.54, 0.54, 0.58, 0.58, 0.62, 0.66, 1.0]
    stage = os.path.join(ASSETS, "amedi_stage.jpg")
    portrait = photo_variant(style, os.path.join(ASSETS, "amedi_portrait_seli_ben_arie_ccbysa4.jpg"))
    handshake = photo_variant(style, os.path.join(ASSETS, "amedi_handshake.jpg"))

    def new(i, photo=None):
        s = prs.slides.add_slide(blank)
        set_bg(s, make_bg(style, f"bg{i:02d}", levels[i], None if style == "paper" else photo))
        add_fade(s)
        return s

    # ---- 1. שער ----
    s = new(0)
    D.text(s, M, 0.6, W - 2 * M, 0.5, [P(R("הערכה חלופית בספרות · שכבה ט׳", 20, D.ACC, True))])
    D.text(s, M, 1.45, W - 2 * M, 2.4, [P(R("סופרמן", 138, D.TEXT, False, D.F_DISPLAY), line=0.9)], anchor="m", name="Title")
    D.text(s, M, 3.85, W - 2 * M, 0.8, [P(R("עידן עמדי", 40, D.ACC, False, D.F_LIGHT))])
    D.text(s, M, 4.65, W - 2 * M, 0.6, [P(R("שיר על פציעה, שיקום ובחירה בחיים", 24, D.MUTED))])
    D.text(s, M, 6.0, W - 2 * M, 0.6, [P(R("אלרואי טואטי   ·   יואב טימסטיט   ·   אלינדב בק", 24, D.TEXT, True))])

    # ---- 2. חלק א' ----
    s = new(1)
    D.title(s, "הקטע שבחרנו")
    D.text(s, 7.3, 1.55, 5.33, 1.3, [P(R("״", 120, D.ACC, False, D.F_DISPLAY), line=0.8)], name="Quote mark")
    D.text(s, 7.3, 2.65, 5.33, 3.0, [
        P(R("אך בסתר ליבי ", 30, D.TEXT, True), R("השבור השלם", 30, D.ACC, True), space_after=4, line=1.05),
        P(R("לא היו לי ספקות", 30, D.TEXT, True), space_after=20, line=1.05),
        P(R("כבר לא מפחיד אותי ", 30, D.TEXT, True), R("למות", 30, D.ACC, True), space_after=4, line=1.05),
        P(R("יותר מפחיד לא ", 30, D.TEXT, True), R("לחיות", 30, D.ACC, True), line=1.05),
    ])
    D.text(s, 7.3, 5.75, 5.33, 0.4, [P(R("עידן עמדי, ״סופרמן״", 16, D.MUTED))])
    D.text(s, M, 1.8, 6.0, 2.3, [
        P(R("למה בחרנו בקטע", 24, D.ACC, True), space_after=6),
        P(R("כאן השיר מגיע לשיא. אחרי הפציעה עמדי כבר לא מפחד מהמוות, הוא מפחד לבזבז את החיים. אהבנו את הצירוף ״השבור השלם״, כי הוא מראה שאפשר להיות פצוע ועדיין שלם מבפנים.", 22), line=1.08),
    ])
    D.text(s, M, 4.35, 6.0, 2.3, [
        P(R("החיבור האישי שלנו", 24, D.ACC, True), space_after=6),
        P(R("גדלנו בתוך המלחמה הזאת, וראינו סביבנו אנשים שיצאו למילואים וחזרו אחרים. הקטע לימד אותנו שגבורה היא לא להיות בלתי שביר, אלא לבחור לחיות גם אחרי שנשברים.", 22), line=1.08),
    ])
    D.pageno(s, 2)

    # ---- 3. האמן ----
    s = new(2)
    picture_fill(s, portrait, 0, 0, 4.75, H, focus=0.2, name="Artist photo")
    D.text(s, 0.25, 7.1, 4.3, 0.3, [P(R("צילום: סלי בן אריה, 2020 · CC BY-SA 4.0", 10, RGBColor(0xDD, 0xDD, 0xDD)), align="l")], anchor="m")
    D.text(s, 5.35, 0.5, W - M - 5.35, 0.95, [P(R("עידן עמדי", 44, D.TEXT, False, D.F_TITLE))], anchor="m", name="Title")
    bio = [
        ("1988", "נולד בירושלים וגדל בשכונת פסגת זאב. בצבא שירת כלוחם ומפקד בהנדסה הקרבית."),
        ("2010", "הגיע לגמר ״כוכב נולד״ עם ״כאב של לוחמים״, שיר שכתב על החוויות שלו כלוחם."),
        ("2017", "התחיל לשחק את שגיא צור בסדרה ״פאודה״, לצד הקריירה במוזיקה."),
        ("2024", "נפצע קשה בעזה בזמן שירות מילואים, ושיתף את הציבור במסע השיקום שלו."),
        ("2025", "הוציא את האלבום ״סופרמן״, ובסוף אותה שנה חזר לשרת במילואים."),
    ]
    y = 1.7
    for i, (yr, t) in enumerate(bio):
        if i:
            D.hair(s, 5.35, y - 0.12, W - M - 5.35)
        D.text(s, 11.35, y, 1.28, 0.5, [P(R(yr, 26, D.ACC, False, D.F_LIGHT))])
        D.text(s, 5.35, y + 0.02, 5.8, 0.9, [P(R(t, 22), line=1.0)])
        y += 1.04
    D.pageno(s, 3, x=5.35)

    # ---- 4–5. מילות השיר ----
    def stanza_paras(stanzas):
        out = []
        for si, st in enumerate(stanzas):
            for li, line in enumerate(st):
                ch = line in CHORUS
                out.append(P(R(line, 22, D.ACC if ch else D.TEXT, ch), line=1.02,
                             space_after=14 if (li == len(st) - 1 and si < len(stanzas) - 1) else 0))
        return out

    for k, (right, left) in enumerate([(LYRICS[0:2], LYRICS[2:4]), (LYRICS[4:6], LYRICS[6:7])]):
        s = new(3 + k)
        D.title(s, "מילות השיר" if k == 0 else "מילות השיר (המשך)")
        D.text(s, 6.95, 1.75, W - M - 6.95, 5.1, stanza_paras(right), name="Lyrics right")
        D.text(s, M, 1.75, 5.9, 5.1, stanza_paras(left), name="Lyrics left")
        D.text(s, 2.0, 6.95, W - M - 2.0, 0.3, [P(R("מילים: עידן עמדי", 12, D.MUTED))], anchor="m")
        D.pageno(s, 4 + k)

    # ---- 6–7. מושגים ----
    concepts = [
        dict(label="מושג ראשון", name="מטאפורה",
             definition="חיבור בין שני דברים מעולמות שונים, בלי מילת דימוי כמו ״כמו״ או ״כְּ״, כדי ליצור משמעות חדשה.",
             quote=[[R("שראתה בתוכי כוח על", 32, D.TEXT, True)],
                    [R("שהוא גם ", 32, D.TEXT, True), R("קביים", 32, D.ACC, True), R(" גם ", 32, D.TEXT, True), R("כנפיים", 32, D.ACC, True)]],
             how="הכוח שאשתו ראתה בו הוא גם ״קביים״, משהו שתומך במי שנפצע ולא מסוגל ללכת לבד, וגם ״כנפיים״, משהו שמרים אותו למעלה.",
             why="השיקום הוא לא רק לחזור ללכת, הוא גם לחזור לחלום. האהבה נתנה לו את שני הדברים."),
        dict(label="מושג שני", name="ניגוד",
             definition="העמדה של שני דברים הפוכים זה מול זה, כדי לחדד את ההבדל ביניהם ולהבליט רעיון.",
             quote=[[R("שרק מי שצעד ב", 32, D.TEXT, True), R("אפלה", 32, D.ACC, True), R(" וב", 32, D.TEXT, True), R("חושך", 32, D.ACC, True)],
                    [R("יודע באמת מהו ", 32, D.TEXT, True), R("אור", 32, D.ACC, True)]],
             how="החושך הוא הפציעה והלילות הקשים, והאור הוא החזרה לחיים. אותו ניגוד חוזר בפזמון: ״למות״ מול ״לחיות״.",
             why="את הטוב מעריכים באמת רק אחרי שעוברים את הרע. כך בנוי כל השיר: הוא עובר מהשבר אל האור."),
    ]
    for k, c in enumerate(concepts):
        s = new(5 + k)
        D.text(s, 8.6, 0.6, W - M - 8.6, 0.4, [P(R(c["label"], 20, D.MUTED, True))])
        D.text(s, 8.6, 1.0, W - M - 8.6, 1.2, [P(R(c["name"], 60, D.ACC, False, D.F_DISPLAY))], anchor="m", name="Title")
        D.text(s, 8.6, 2.35, W - M - 8.6, 1.9, [P(R(c["definition"], 22, D.TEXT), line=1.06)])
        D.text(s, M, 0.75, 7.3, 1.1, [P(R("״", 96, D.ACC, False, D.F_DISPLAY), line=0.8)])
        D.text(s, M, 1.75, 7.3, 2.3, [P(*q, line=1.08, space_after=4) for q in c["quote"]], anchor="t")
        D.hair(s, M, 4.45, W - 2 * M)
        D.text(s, 6.95, 4.7, W - M - 6.95, 2.2, [P(R("איך זה מופיע בשיר", 22, D.ACC, True), space_after=6),
                                                  P(R(c["how"], 22), line=1.04)])
        D.text(s, M, 4.7, 5.6, 2.2, [P(R("מה זה מוסיף למשמעות", 22, D.ACC, True), space_after=6),
                                      P(R(c["why"], 22), line=1.04)])
        D.pageno(s, 6 + k)

    # ---- 8. מילים מהמלחמה: קומפוזיציה טיפוגרפית ----
    s = new(7)
    D.title(s, "מילים שמעוררות את המלחמה")
    D.text(s, M, 1.45, W - 2 * M, 0.5, [P(R("מילים וצירופים מתוך השיר שמחזירים אותנו אל הקרב ואל הפציעה", 22, D.MUTED))])
    # (מילה, גודל, קצה ימני, y, מודגש)
    words = [
        ("נפצעתי", 66, 12.63, 2.15, True), ("לא הרמתי ידיים", 40, 8.4, 2.45, False),
        ("חטפתי", 44, 3.6, 2.35, False),
        ("שאגתי מתוך החזה", 38, 12.63, 3.65, False), ("קביים", 58, 6.6, 3.45, True),
        ("אוויר דליל", 34, 2.9, 3.75, False),
        ("מחנק בעליות", 36, 11.2, 5.05, False), ("החושך שעברתי", 52, 6.3, 4.85, True),
    ]
    for wd, sz, xr, y, hi in words:
        bw = min(6.0, xr - M)
        D.text(s, xr - bw, y, bw, sz / 72 * 1.4,
               [P(R(wd, sz, D.ACC if hi else D.TEXT, True, D.F_DISPLAY if hi else D.F_BODY))])
    D.pageno(s, 8)

    # ---- 9. אירועים ----
    s = new(8)
    D.title(s, "האירועים שהמילים מעוררות")
    picture_fill(s, handshake, M, 1.75, 5.0, 3.55, focus=0.3, name="Event photo")
    D.text(s, M, 5.45, 5.0, 1.3, [P(R("עידן עמדי במדי מילואים, בלחיצת יד עם מפקד בכיר. מאחורי ה״סופרמן״ שבשיר עומד לוחם אמיתי.", 18, D.MUTED), line=1.02)])
    events = [
        ("אוקטובר 2023", "פורצת מלחמת ״חרבות ברזל״, ועמדי מגויס למילואים."),
        ("עזה", "כלוחם הנדסה, הוא נלחם בתוך רצועת עזה."),
        ("8 בינואר 2024", "נפצע קשה מפיצוץ בזמן פעולה להשמדת מנהרה."),
        ("שיבא", "מגיע לבית החולים מורדם ומונשם, עם רסיסים בכל הגוף."),
        ("השנה שאחרי", "יותר משנה של שיקום. מתוכה נולד השיר ״סופרמן״."),
    ]
    y = 1.7
    for i, (lbl, t) in enumerate(events):
        if i:
            D.hair(s, 6.25, y - 0.1, W - M - 6.25)
        D.text(s, 10.55, y, W - M - 10.55, 0.9, [P(R(lbl, 20, D.ACC, True), line=1.0)])
        D.text(s, 6.25, y, 4.15, 0.9, [P(R(t, 22), line=1.0)])
        y += 1.02
    D.pageno(s, 9)

    # ---- 10. סיכום ----
    s = new(9, photo=stage)
    D.text(s, M, 1.55, W - 2 * M, 4.2, [
        P(R("סופרמן בקומיקס אף פעם לא נפצע.", 40, D.TEXT, True), align="ctr", space_after=16, line=1.05),
        P(R("עידן עמדי נפצע, נשבר וקם.", 44, D.ACC, False, D.F_DISPLAY), align="ctr", space_after=16, line=1.05),
        P(R("בעינינו, זה סופרמן אמיתי.", 40, D.TEXT, False, D.F_LIGHT), align="ctr", line=1.05),
    ], anchor="m", name="Title")
    D.pageno(s, 10)

    suffix = "" if style == "night" else f"_{style}"
    out = os.path.join(HERE, f"superman{suffix}.pptx")
    prs.save(out)
    print("saved", out)
    return out


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "night")
