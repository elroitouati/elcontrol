# -*- coding: utf-8 -*-
"""מייצר את דף בחירת הפרויקטים. python3 make_picker.py <out_dir>
TAKEN: מספר פרויקט -> שמות (מתעדכן כשתלמידים בוחרים)."""
import sys, os, re, html
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from songs import SONGS

TAKEN = {}  # למשל: {"03": "דנה כהן ויואב לוי"}

STYLE = {"flag": "דגל: לבן וכחול, אות עברית קלאסית", "sunrise": "שחור עם אופק בוער", "space": "מינימליסטי, הרבה מקום ריק",
         "field": "שדה ירוק עם כלניות", "bloom": "ירוק לילה וכלנית אדומה", "notebook": "דף מחברת וכתב יד",
         "concrete": "פוסטר בטון וצהוב שלטים", "dusk": "שחר במדבר, אינדיגו", "lighthouse": "מגדלור: קרן אור בחושך",
         "diary": "יומן אישי, כתב יד"}
REL = {"03": "12 בספטמבר 2024"}
TABS = [("cover", "שער"), ("artist", "האמן"), ("inside", "הקטע שבחרנו"), ("concept", "ניתוח מושג")]

e = html.escape


def release(s, n):
    m = re.search(r"יצא ב-(.+)$", s["credits"])
    return m.group(1) if m else REL.get(n, "")


def project(i, s):
    n = f"{i:02d}"
    taken = TAKEN.get(n)
    status = f'<span class="chip taken">נבחר · {e(taken)}</span>' if taken else '<span class="chip free">פנוי</span>'
    tabs = "".join(
        f'<button type="button" class="tab" role="tab" aria-selected="{"true" if k == 0 else "false"}" data-src="images/{n}-{t}.jpg" data-alt="שקף {lab} בפרויקט {n}">{lab}</button>'
        for k, (t, lab) in enumerate(TABS))
    msg = f"בחרנו בפרויקט {n}: ״{s['title']}״ של {s['artist_short']}. השמות שלנו: "
    ex = "<br>".join(e(l) for l in s["excerpt"][:2])
    btn = ('<p class="takenmsg">הפרויקט הזה כבר נבחר.</p>' if taken else
           f'<button type="button" class="copy" data-msg="{e(msg)}" data-note="note{n}">העתקת הודעת בחירה</button>'
           f'<span class="note" id="note{n}" role="status"></span>')
    return f'''
<article class="proj{' is-taken' if taken else ''}" id="p{n}">
  <div class="viewer">
    <img class="big" src="images/{n}-cover.jpg" alt="שקף השער בפרויקט {n}" width="1100" height="619" loading="lazy">
    <div class="tabs" role="tablist" aria-label="שקפים מפרויקט {n}">{tabs}</div>
  </div>
  <div class="meta">
    <div class="top"><span class="code">{n}</span>{status}</div>
    <h2>{e(s["title"])}</h2>
    <p class="artist">{e(s["artist"])}</p>
    <p class="about">{e(s["tagline"])}. יצא ב-{e(release(s, n))}.</p>
    <blockquote>{ex}</blockquote>
    <dl>
      <div><dt>עיצוב</dt><dd>{e(STYLE[s["style"]])}</dd></div>
      <div><dt>מושגים</dt><dd>{e(" · ".join(c["name"] for c in s["concepts"]))}</dd></div>
    </dl>
    <div class="pick">{btn}</div>
  </div>
</article>'''


def build(out_dir):
    mosaic = "".join(
        f'<a class="tile{" is-taken" if TAKEN.get(f"{i:02d}") else ""}" href="#p{i:02d}"><img src="images/{i:02d}-cover.jpg" alt="" width="1100" height="619" loading="lazy">'
        f'<span class="tnum">{i:02d}</span><span class="sr">{e(s["title"])}</span></a>' for i, s in enumerate(SONGS, 1))
    free = len(SONGS) - len(TAKEN)
    page = f'''<title>בחירת פרויקט בספרות</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Frank+Ruhl+Libre:wght@500;800&family=Assistant:wght@400;600;700&display=swap">
<style>
/* Layout: a dark gallery band with all ten covers, then one row per project: slide viewer beside its details */
:root{{
  --bg:#F2F3F5; --surface:#FFFFFF; --ink:#14161B; --muted:#5B6170; --line:#D8DBE1;
  --accent:#1D4FBF; --on-accent:#FFFFFF; --band:#0F1218; --band-ink:#F3F4F6; --band-muted:#A6ACB8;
  --free:#1C7A4C; --taken:#B4462D;
  --display:"Frank Ruhl Libre", "David Libre", Georgia, serif;
  --body:"Assistant", "Heebo", Arial, sans-serif;
}}
@media (prefers-color-scheme: dark){{:root:not([data-theme="light"]){{
  --bg:#101318; --surface:#181B22; --ink:#ECEEF2; --muted:#A2A8B5; --line:#2A2F39; --accent:#86A8FF; --on-accent:#0D1220;
  --band:#07090C; --free:#5CC596; --taken:#F08A70; color-scheme:dark}}}}
:root[data-theme="dark"]{{--bg:#101318; --surface:#181B22; --ink:#ECEEF2; --muted:#A2A8B5; --line:#2A2F39; --accent:#86A8FF; --on-accent:#0D1220;
  --band:#07090C; --free:#5CC596; --taken:#F08A70; color-scheme:dark}}
html{{direction:rtl; scroll-behavior:smooth}}
@media (prefers-reduced-motion:reduce){{html{{scroll-behavior:auto}} *{{transition:none!important}}}}
body{{background:var(--bg); color:var(--ink); font-family:var(--body); font-size:17px; line-height:1.55}}
.inner{{max-width:1180px; margin:0 auto; padding-left:max(16px, env(safe-area-inset-left, 0px)); padding-right:max(16px, env(safe-area-inset-right, 0px))}}
.band{{background:var(--band); color:var(--band-ink)}}
.band .inner{{padding-block:44px 40px; display:grid; gap:22px}}
.kicker{{margin:0; color:var(--band-muted); font-weight:600; letter-spacing:.03em}}
h1{{margin:0; font-family:var(--display); font-weight:800; font-size:clamp(40px,7vw,76px); line-height:1; text-wrap:balance}}
.lead{{margin:0; max-width:60ch; color:var(--band-muted); font-size:18px}}
.count{{color:var(--band-ink); font-weight:700}}
.mosaic{{display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:10px}}
.tile{{position:relative; display:block; color:var(--band-ink); text-decoration:none; border-radius:4px; overflow:hidden; outline-offset:3px}}
.tile img{{display:block; width:100%; height:auto; aspect-ratio:16/9; object-fit:cover; transition:transform .25s ease, filter .25s ease}}
.tile:hover img{{transform:scale(1.04)}}
.tile:focus-visible{{outline:3px solid var(--band-ink)}}
.tile.is-taken img{{filter:grayscale(1) brightness(.55)}}
.tnum{{position:absolute; bottom:6px; left:8px; font-weight:700; font-size:13px; background:rgba(0,0,0,.55); padding:1px 7px; border-radius:3px; font-variant-numeric:tabular-nums}}
.sr{{position:absolute; width:1px; height:1px; overflow:hidden; clip:rect(0 0 0 0)}}
.tname{{position:absolute; bottom:6px; right:8px; left:8px; font-size:13px; font-weight:600; text-shadow:0 1px 6px rgba(0,0,0,.8); white-space:nowrap; overflow:hidden; text-overflow:ellipsis}}
.steps{{margin:0; padding:0; list-style:none; display:flex; flex-wrap:wrap; gap:8px 26px; color:var(--band-ink)}}
.steps b{{font-family:var(--display); font-size:22px; margin-left:6px}}
main .inner{{padding-block:10px 60px}}
.proj{{display:grid; grid-template-columns:minmax(0,1.45fr) minmax(0,1fr); gap:34px; padding-block:40px; border-bottom:1px solid var(--line); scroll-margin-top:12px}}
.viewer{{display:grid; gap:10px; min-width:0}}
.big{{display:block; width:100%; height:auto; aspect-ratio:16/9; object-fit:cover; border:1px solid var(--line); box-shadow:0 18px 40px -26px rgba(0,0,0,.55)}}
.tabs{{display:flex; flex-wrap:wrap; gap:6px}}
.tab{{font:inherit; font-size:15px; border:1px solid var(--line); background:var(--surface); color:var(--ink); padding:6px 12px; border-radius:999px; cursor:pointer}}
.tab[aria-selected="true"]{{background:var(--ink); color:var(--bg); border-color:var(--ink)}}
.tab:focus-visible, .copy:focus-visible{{outline:3px solid var(--accent); outline-offset:2px}}
.meta{{display:flex; flex-direction:column; gap:6px; min-width:0}}
.top{{display:flex; align-items:center; gap:12px}}
.code{{font-family:var(--display); font-weight:800; font-size:34px; line-height:1; color:var(--accent); font-variant-numeric:tabular-nums}}
.chip{{font-size:13px; font-weight:700; padding:3px 10px; border-radius:999px; border:1.5px solid currentColor}}
.chip.free{{color:var(--free)}}
.chip.taken{{color:var(--taken)}}
h2{{margin:4px 0 0; font-family:var(--display); font-weight:800; font-size:clamp(30px,4vw,44px); line-height:1.05; text-wrap:balance}}
.artist{{margin:0; font-weight:700; color:var(--muted)}}
.about{{margin:6px 0 0}}
blockquote{{margin:10px 0 4px; padding:0 14px 0 0; border-right:3px solid var(--line); font-family:var(--display); font-size:19px; line-height:1.45}}
dl{{margin:8px 0 0; display:grid; gap:6px}}
dl div{{display:grid; grid-template-columns:76px minmax(0,1fr); gap:10px}}
dt{{color:var(--muted); font-size:14px; padding-top:2px}}
dd{{margin:0}}
.pick{{margin-top:16px; display:flex; flex-wrap:wrap; align-items:center; gap:12px}}
.copy{{font:inherit; font-weight:700; border:0; border-radius:6px; padding:11px 18px; background:var(--accent); color:var(--on-accent); cursor:pointer}}
.copy:hover{{filter:brightness(1.08)}}
.note{{font-size:14px; color:var(--muted)}}
.takenmsg{{margin:0; color:var(--taken); font-weight:600}}
.proj.is-taken .big{{filter:grayscale(.6)}}
footer{{padding-top:26px; color:var(--muted); font-size:14px; max-width:70ch}}
@media (max-width:900px){{.mosaic{{grid-template-columns:repeat(2,minmax(0,1fr))}} .proj{{grid-template-columns:minmax(0,1fr); gap:18px}}}}
</style>
<section class="band">
  <div class="inner">
    <p class="kicker">הערכה חלופית בספרות · שכבה ט׳</p>
    <h1>בוחרים פרויקט</h1>
    <p class="lead">עשרה פרויקטים מוכנים. כל אחד עוסק בשיר אחר מתקופת המלחמה, ובנוי בעיצוב אחר לגמרי. <span class="count">{free} מתוך {len(SONGS)} פנויים.</span></p>
    <nav class="mosaic" aria-label="כל הפרויקטים">{mosaic}</nav>
    <ol class="steps"><li><b>1</b>לוחצים על פרויקט ומדפדפים בשקפים</li><li><b>2</b>בוחרים אחד</li><li><b>3</b>שולחים לאלרואי את מספר הפרויקט ואת השמות</li></ol>
  </div>
</section>
<main><div class="inner">
{"".join(project(i, s) for i, s in enumerate(SONGS, 1))}
<footer>בכל פרויקט תשעה שקפים לפי המחוון: הקטע שבחרנו, האמן, מילות השיר, שני מושגים, מילים מהמלחמה, האירועים וסיכום. השמות נכנסים לשער אחרי הבחירה. כל התמונות ברישיון חופשי, עם קרדיט לצלם.</footer>
</div></main>
<script>
document.querySelectorAll(".proj").forEach(function(p){{
  var big = p.querySelector(".big");
  p.querySelectorAll(".tab").forEach(function(t){{
    t.addEventListener("click", function(){{
      p.querySelectorAll(".tab").forEach(function(o){{ o.setAttribute("aria-selected", o === t ? "true" : "false"); }});
      big.src = t.getAttribute("data-src"); big.alt = t.getAttribute("data-alt");
    }});
  }});
}});
document.querySelectorAll(".copy").forEach(function(btn){{
  btn.addEventListener("click", function(){{
    var note = document.getElementById(btn.getAttribute("data-note"));
    var msg = btn.getAttribute("data-msg");
    function fallback(){{ note.textContent = msg; var r = document.createRange(); r.selectNodeContents(note); var s = window.getSelection(); s.removeAllRanges(); s.addRange(r); }}
    try {{ navigator.clipboard.writeText(msg).then(function(){{ note.textContent = "הועתק. מוסיפים את השמות ושולחים לאלרואי."; }}, fallback); }}
    catch (err) {{ fallback(); }}
  }});
}});
</script>
'''
    open(os.path.join(out_dir, "project-picker.html"), "w", encoding="utf-8").write(page)
    return page


SITE = "https://elroitouati.github.io/elcontrol/"


def standalone(page, out_dir):
    """גרסה עצמאית לאתר (GitHub Pages) עם תצוגה מקדימה לוואטסאפ."""
    head = ('<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
            '<meta name="description" content="עשרה פרויקטים מוכנים להערכה חלופית בספרות, שכבה ט׳. מדפדפים, בוחרים ושולחים.">'
            '<meta property="og:type" content="website"><meta property="og:locale" content="he_IL">'
            '<meta property="og:title" content="בוחרים פרויקט · הערכה חלופית בספרות">'
            '<meta property="og:description" content="עשרה שירים, עשרה עיצובים. בוחרים אחד ושולחים לאלרואי.">'
            f'<meta property="og:image" content="{SITE}og.jpg"><meta property="og:image:width" content="1200"><meta property="og:image:height" content="630">'
            f'<meta property="og:url" content="{SITE}"><meta name="twitter:card" content="summary_large_image">'
            '<style>*,*::before,*::after{box-sizing:border-box}:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}'
            'body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>')
    i = page.index("</title>") + len("</title>")
    out = head + page[:i] + page[i:].replace("<style>", "<style>", 1)
    out = out.replace('<meta name="claude:viewport-fit-sides" content="cover">', "")
    out = out.replace("<section class=\"band\">", "</head><body><section class=\"band\">", 1) + "</body></html>"
    open(os.path.join(out_dir, "index.html"), "w", encoding="utf-8").write(out)


if __name__ == "__main__":
    p = build(sys.argv[1])
    if len(sys.argv) > 2:
        standalone(p, sys.argv[2])
