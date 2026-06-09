"""PremAI Brand Bible v4 — verified live site, monochrome, correct portfolio, Lugano aesthetic."""
import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def rgb(r,g,b): return RGBColor(r,g,b)

# ── Monochrome palette — live premai.io, verified 9 Jun 2026 ──────────────
INK     = rgb(10,  10,  10)   # logo black
DARK    = rgb(22,  22,  26)   # dark surface
GRAY    = rgb(88,  90,  94)   # secondary text
MID     = rgb(135, 136, 132)  # tertiary, rules
LIGHT   = rgb(200, 200, 196)  # very light gray
PAPER   = rgb(252, 252, 250)  # default ground
PAPER2  = rgb(242, 240, 234)  # alternate row / panels
WHITE   = rgb(255, 255, 255)
# Lugano-sky tones for the lake aesthetic
SKY     = rgb(215, 225, 235)
SKY2    = rgb(190, 205, 220)
SKY3    = rgb(170, 188, 208)
# Ridge / mountain silhouette grayscale
R1      = rgb(218, 218, 214)
R2      = rgb(195, 196, 190)
R3      = rgb(172, 173, 166)
# Functional only — never brand
FN_RED  = rgb(175, 50, 50)
FN_GRN  = rgb(40, 110, 70)
FN_REDT = rgb(252, 234, 234)
FN_GRNT = rgb(232, 247, 238)

FONT = 'Pretendard'   # verified in premAI-io/static.premai.io/fonts

def lerp(a,b,t): return rgb(int(a[0]+(b[0]-a[0])*t), int(a[1]+(b[1]-a[1])*t), int(a[2]+(b[2]-a[2])*t))

# ── Core helpers ─────────────────────────────────────────────────────────
def add_bg(slide, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.334), Inches(7.5))
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s

def box(slide, l, t, w, h, fill=None, line=None, lw=0.5):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = fill
    else: s.fill.background()
    if line: s.line.color.rgb = line; s.line.width = Pt(lw)
    else: s.line.fill.background()
    return s

def oval(slide, cx, cy, r, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx-r), Inches(cy-r), Inches(2*r), Inches(2*r))
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s

def T(slide, text, l, t, w, h, sz=11, bold=False, color=INK, align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    f = r.font; f.size=Pt(sz); f.bold=bold; f.italic=italic
    f.color.rgb=color; f.name=FONT
    return tb

def Lbl(slide, text, l, t, w=10, color=MID):
    return T(slide, text.upper(), l, t, w, 0.28, sz=8, bold=True, color=color)

def Rule(slide, l, t, w, color=INK, thick=0.01):
    return box(slide, l, t, w, thick, fill=color)

def ridge(slide, pts, fill):
    fb = slide.shapes.build_freeform(Emu(Inches(pts[0][0])), Emu(Inches(7.5)), scale=1.0)
    seq = [(Inches(x), Inches(y)) for x, y in pts] + [(Inches(pts[-1][0]), Inches(7.5))]
    fb.add_line_segments([(Emu(x), Emu(y)) for x, y in seq], close=True)
    sh = fb.convert_to_shape()
    sh.fill.solid(); sh.fill.fore_color.rgb = fill; sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def lake_layer(slide, l, t, w, h, fill):
    """Horizontal soft lake-surface band."""
    return box(slide, l, t, w, h, fill=fill)

def dot_mark(slide, cx, cy, scale=1.0, fill=INK):
    """Simplified Prem dot cluster (echoes logo structure)."""
    pts = [
        (0,0,0.055),(0.13,0,0.04),(0.26,0,0.048),(0.39,0,0.035),(0.52,0,0.042),
        (0.065,0.13,0.038),(0.195,0.13,0.052),(0.325,0.13,0.040),(0.455,0.13,0.048),(0.585,0.13,0.030),
        (0,0.26,0.042),(0.13,0.26,0.046),(0.26,0.26,0.058),(0.39,0.26,0.044),(0.52,0.26,0.038),
        (0.065,0.39,0.028),(0.195,0.39,0.048),(0.325,0.39,0.055),(0.455,0.39,0.040),(0.585,0.39,0.044),
        (0,0.52,0.032),(0.13,0.52,0.040),(0.26,0.52,0.048),(0.39,0.52,0.034),(0.52,0.52,0.028),
    ]
    ox = cx - 0.29*scale; oy = cy - 0.26*scale
    for (dx, dy, r) in pts:
        oval(slide, ox+dx*scale, oy+dy*scale, r*scale, fill)

def dot_sig(slide, l, t, fill=INK):
    """Five-dot signature row — echoes logo."""
    for i, r in enumerate([0.050, 0.032, 0.044, 0.028, 0.038]):
        oval(slide, l + i*0.165, t, r, fill)

# ── Switzerland dot map ──────────────────────────────────────────────────
CH = [(5.96,46.14),(6.06,46.45),(6.45,46.78),(6.43,46.93),(6.95,47.25),(7.35,47.43),
      (7.59,47.58),(8.20,47.62),(8.57,47.80),(9.05,47.68),(9.56,47.54),(9.67,47.06),
      (10.47,46.86),(10.39,46.63),(10.04,46.22),(9.28,46.32),(9.04,45.82),(8.44,46.25),
      (7.86,45.92),(7.04,45.90),(6.80,46.15),(6.80,46.43),(6.24,46.34),(6.30,46.25)]
COS = math.cos(math.radians(46.8))

def in_ch(lon, lat):
    inside=False; j=len(CH)-1
    for i in range(len(CH)):
        xi,yi=CH[i]; xj,yj=CH[j]
        if (yi>lat)!=(yj>lat) and lon<(xj-xi)*(lat-yi)/(yj-yi)+xi: inside=not inside
        j=i
    return inside

def ch_map(slide, l, t, w, sp=0.10, r=0.030, fill_fn=None, lugano=True):
    lons=[p[0] for p in CH]; lats=[p[1] for p in CH]
    lo,hi = min(lons),max(lons); la0,la1 = min(lats),max(lats)
    ew=(hi-lo)*COS; eh=la1-la0; sc=w/ew; h=eh*sc
    nx=int(w/sp); ny=int(h/sp)
    for iy in range(ny+1):
        for ix in range(nx+1):
            px=ix*sp; py=iy*sp
            lon=lo+(px/sc)/COS; lat=la1-py/sc
            if in_ch(lon,lat):
                t_ = (px/w)*0.55+(py/h)*0.45
                c = fill_fn(t_) if fill_fn else lerp((10,10,10),(155,155,150),t_)
                oval(slide, l+px, t+py, r, c)
    if lugano:
        lx=l+(8.95-lo)*COS*sc; ly=t+(la1-46.0)*sc
        d=oval(slide, lx, ly, 0.052, WHITE)
        d.line.color.rgb=INK; d.line.width=Pt(1.8)

# ── Build ────────────────────────────────────────────────────────────────
prs = Presentation()
prs.slide_width=Inches(13.334); prs.slide_height=Inches(7.5)
blank=prs.slide_layouts[6]
def NS(): return prs.slides.add_slide(blank)

def HDR(slide, section):
    Lbl(slide, section, 0.55, 0.44)
    Rule(slide, 0.55, 0.75, 12.24)

def NTS(slide, t): slide.notes_slide.notes_text_frame.text = t


# ══════════════════════════════════════════════════════════════
# S1  COVER
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
# Lugano lake layers (light, airy — matches "Designed & powered in Switzerland" section)
for i,(c,y_,h_) in enumerate([(SKY3,4.85,2.65),(SKY2,5.35,2.15),(SKY,5.9,1.6)]):
    box(s, 0, y_, 13.334, h_, fill=c)
# Alpine silhouettes over lake
ridge(s, [(3.5,4.9),(4.8,4.2),(6.2,4.75),(7.6,3.85),(9.0,4.5),(10.5,3.95),(12.2,4.65),(13.334,4.3)], R3)
ridge(s, [(5.5,5.25),(6.8,4.7),(8.1,5.1),(9.5,4.45),(11.0,5.0),(12.5,4.8),(13.334,5.1)], R2)
# White paper ground (top)
box(s, 0, 0, 13.334, 4.4, fill=PAPER)
Rule(s, 0, 4.4, 13.334, color=PAPER2, thick=0.02)

dot_mark(s, 1.12, 1.45, scale=1.0)
T(s,'PREM', 1.72, 1.1, 5, 0.65, sz=36, bold=True, color=INK)
Rule(s, 0.55, 1.82, 4.8, color=MID, thick=0.008)
T(s,'Brand Bible', 0.55, 2.0, 5, 0.55, sz=26, color=INK)
T(s,'Private Super Intelligence', 0.55, 2.62, 6, 0.38, sz=13, color=GRAY)
Rule(s, 0.55, 3.12, 1.8, color=INK)
T(s,'Brand architecture  ·  Digital focus  ·  Voice and tone\nMonochrome identity  ·  Platform playbooks  ·  Governance',
  0.55, 3.32, 6.5, 0.72, sz=9.5, color=MID)
T(s,'v4.0   June 2026   Internal use only', 0.55, 4.1, 5, 0.28, sz=8.5, color=LIGHT)
# Right: Switzerland dot map on the lake
ch_map(s, 7.5, 0.85, 5.0, sp=0.105, r=0.032)
T(s,'Designed & powered in Switzerland', 7.5, 4.15, 5.5, 0.32, sz=9, bold=True, color=GRAY)
NTS(s,'Logo verified from live premai.io (9 Jun 2026): black dot-cluster mark, geometric grotesque wordmark. Site hero: "Private Super Intelligence" / "Prem makes it private, verifiable, and sovereign." Lugano aerial view used in "Designed & powered in Switzerland" section. Three office addresses in footer. Tagline: "Intelligence in Service".')


# ══════════════════════════════════════════════════════════════
# S2  CONTENTS
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'Contents')
items=[
    ('01','Brand architecture','Prem masterbrand and six verified products'),
    ('02','The company','Intelligence in Service · Lugano · three offices'),
    ('03','Digital focus','40 Fluso · 20 Confidential API · 10 Sotto · 30 rest'),
    ('04','Positioning: Prem','Private Super Intelligence'),
    ('05','Positioning: Fluso','Work deeper, not longer'),
    ('06','Positioning: other products','Confidential API, Enclave, Concierge, Studio, Sotto'),
    ('07','The category','Verifiable AI'),
    ('08','Voice and tone','Four principles · the banned list · examples'),
    ('09','Color','Monochrome, like the mark'),
    ('10','Typography','The Prem wordmark typeface'),
    ('11','The Swiss visual world','Alps, Lugano, the dot map, the grid'),
    ('12','Social templates','Reel cover, carousel, quote card'),
    ('13','Platform: LinkedIn','Prem-led authority'),
    ('14','Platform: X','Fluso-led builder credibility'),
    ('15','Platform: Instagram + Facebook','Visual world and warm remarketing'),
    ('16','Content pillars','What each focus bucket publishes'),
    ('17','Audience architecture','Persona to platform mapping'),
    ('18','Vocabulary','Words we own · words we never use'),
    ('19','Positioning test','The six-question check'),
    ('20','Governance','Sign-off by role · no exceptions'),
]
for i,(num,title,sub) in enumerate(items):
    col=i//10; row=i%10
    x=0.55+col*6.5; y=1.0+row*0.64
    T(s,num, x, y, 0.6, 0.32, sz=9, bold=True, color=MID)
    T(s,title, x+0.68, y, 5.2, 0.32, sz=11.5, bold=True, color=INK)
    T(s,sub, x+0.68, y+0.33, 5.4, 0.28, sz=8.5, color=LIGHT)
box(s, 6.85, 0.95, 0.01, 5.65, fill=PAPER2)  # vertical column divider
dot_sig(s, 12.5, 7.1)


# ══════════════════════════════════════════════════════════════
# S3  BRAND ARCHITECTURE
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'01  Brand architecture')
T(s,'One masterbrand. Six products. One direction.', 0.55, 0.9, 12, 0.62, sz=26, bold=True, color=INK)

# Masterbrand band
box(s, 0.55, 1.72, 12.24, 0.95, fill=DARK)
dot_mark(s, 1.08, 2.15, scale=0.55, fill=PAPER)
T(s,'PREM', 1.65, 1.93, 3, 0.38, sz=20, bold=True, color=WHITE)
T(s,'Intelligence in Service', 1.65, 2.32, 3.5, 0.28, sz=9, color=MID)
T(s,'Private · Verifiable · Sovereign  ·  premai.io  ·  Lugano, Switzerland',
  4.5, 2.12, 8.0, 0.5, sz=10, color=PAPER2)

# Products group (3 cards)
T(s,'PRODUCTS', 0.55, 2.88, 4, 0.28, sz=8, bold=True, color=MID)
prod_cards = [
    ('Confidential API','Build with encrypted inference','Build with encrypted inference over private models. The developer and enterprise API entry point.'),
    ('Prem Enclave',   'Verifiable infrastructure',     'Runs models inside a trusted execution environment. Hardware-level privacy, not policy-level promises.'),
    ('Fluso',          'Private AI workspace',          'The flagship. Deep-work AI for knowledge professionals. Compounding memory, 50+ connectors, session audit trail.'),
]
for i,(name,tag,desc) in enumerate(prod_cards):
    x=0.55+i*4.1
    box(s, x, 3.22, 3.85, 2.12, fill=WHITE, line=PAPER2, lw=0.5)
    flag = (name=='Fluso')
    if flag:
        box(s, x, 3.22, 3.85, 0.04, fill=INK)
        box(s, x+2.2, 3.27, 1.45, 0.28, fill=INK)
        T(s,'FLAGSHIP', x+2.2, 3.28, 1.45, 0.26, sz=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    T(s,name, x+0.22, 3.38, 3.4, 0.35, sz=13, bold=True, color=INK)
    T(s,tag.upper(), x+0.22, 3.75, 3.4, 0.26, sz=7.5, bold=True, color=MID)
    T(s,desc, x+0.22, 4.08, 3.42, 1.1, sz=9, color=GRAY)

# Research group (3 cards on one band)
T(s,'RESEARCH & PLATFORM', 0.55, 5.5, 5, 0.28, sz=8, bold=True, color=MID)
res_cards = [
    ('Prem Studio',    'Customize models on private data'),
    ('Prem Concierge', 'Secure daily workflow assistant'),
    ('Prem Sotto',     'Encrypted dictation for every app'),
]
for i,(name,desc) in enumerate(res_cards):
    x=0.55+i*4.1
    box(s, x, 5.85, 3.85, 1.15, fill=WHITE, line=PAPER2, lw=0.5)
    T(s,name, x+0.22, 5.97, 3.4, 0.32, sz=12, bold=True, color=INK)
    T(s,desc, x+0.22, 6.32, 3.4, 0.6, sz=9, color=GRAY)
NTS(s,'All products verified from live premai.io navigation (9 Jun 2026). Products dropdown: Confidential API, Prem Enclave, Fluso. Research dropdown: Research, Prem Studio, Prem Concierge, Prem Sotto. Descriptions use exact site copy.')


# ══════════════════════════════════════════════════════════════
# S4  THE COMPANY
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
# Lugano lake full-bleed right panel
for (c_,y_,h_) in [(SKY3,0,7.5),(SKY2,1.5,6),(SKY,3.0,4.5)]:
    box(s, 7.5, y_, 5.834, h_, fill=c_)
ridge(s,[(7.5,2.8),(8.6,2.1),(9.8,2.65),(11.0,1.9),(12.2,2.55),(13.334,2.2)], R3)
ridge(s,[(7.5,3.5),(8.8,2.9),(10.2,3.4),(11.6,2.75),(13.334,3.2)], R2)
box(s, 7.5, 0, 0.02, 7.5, fill=PAPER2)

HDR(s,'02  The company')
T(s,'Intelligence in Service', 0.55, 0.95, 6.5, 0.7, sz=30, bold=True, color=INK)
Rule(s, 0.55, 1.75, 3.5, color=INK)
T(s,'"AI is the most powerful technology of our era.\nPrem makes it private, verifiable, and sovereign."',
  0.55, 1.92, 6.5, 0.85, sz=12, color=GRAY, italic=True)
T(s,'Designed & powered in Switzerland', 0.55, 2.95, 6.5, 0.35, sz=11, bold=True, color=INK)
T(s,'Built in Lugano. The Alpine geography is not a backdrop — it is a design principle. Precise, quiet, built to last.',
  0.55, 3.38, 6.6, 0.55, sz=10, color=GRAY)

offices=[
    ('Lugano, Switzerland','Crocicchio Cortogna 6, 6900 Lugano','Primary HQ'),
    ('Wilmington, Delaware','901 N Market Street, Suite 100','US entity'),
    ('Putignano, Bari, Italy','Via Giuseppe Verdi 6, 70017 Putignano','R&D'),
]
for i,(city,addr,role) in enumerate(offices):
    y=4.15+i*0.8
    box(s, 0.55, y, 6.65, 0.68, fill=WHITE, line=PAPER2, lw=0.5)
    T(s,city, 0.78, y+0.08, 3.5, 0.28, sz=10.5, bold=True, color=INK)
    T(s,addr, 0.78, y+0.38, 3.8, 0.25, sz=8.5, color=GRAY)
    T(s,role.upper(), 4.8, y+0.18, 2.0, 0.28, sz=7.5, bold=True, color=MID, align=PP_ALIGN.RIGHT)

T(s,'Trust Center', 0.55, 7.08, 2, 0.28, sz=8.5, color=MID)
T(s,'premai.io/security', 1.6, 7.08, 3, 0.28, sz=8.5, bold=True, color=INK)
# Lugano label on right panel
T(s,'Lake Lugano', 9.0, 6.65, 3, 0.28, sz=9, bold=True, color=GRAY)
dot_sig(s, 9.0, 7.1, fill=GRAY)
NTS(s,'Company information from live premai.io footer (9 Jun 2026). Three office addresses verified. Tagline "Intelligence in Service" from footer. "Designed & powered in Switzerland" from site section. Security/trust center at premai.io/security.')


# ══════════════════════════════════════════════════════════════
# S5  DIGITAL FOCUS
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'03  Digital focus  |  Attention and resource allocation')
T(s,'The split governs production investment, not post frequency.', 0.55, 0.9, 12.2, 0.45, sz=13, bold=True, color=INK)
T(s,'Content does not announce these ratios. The editorial voice is Prem and Fluso. Other products surface when there is a genuine story.',
  0.55, 1.42, 12.2, 0.38, sz=10.5, color=GRAY)

# Proportional bar
splits=[(40,'FLUSO','Flagship'),(20,'CONFIDENTIAL API','Enterprise dev'),(10,'SOTTO','Encrypted voice'),(30,'REST','Enclave, Concierge, Studio, brand')]
fills=[DARK, rgb(50,50,50), rgb(100,100,100), PAPER2]
tcolors=[WHITE, WHITE, WHITE, INK]
x0=0.55; bar_y=2.1; bar_h=1.1
for i,(pct,name,sub) in enumerate(splits):
    bw=12.24*pct/100
    box(s, x0, bar_y, bw, bar_h, fill=fills[i])
    if pct>=15:
        T(s,f'{pct}%', x0+0.15, bar_y+0.12, bw-0.2, 0.45, sz=22, bold=True, color=tcolors[i])
        T(s,name, x0+0.15, bar_y+0.58, bw-0.2, 0.38, sz=8.5, bold=True, color=tcolors[i])
    else:
        T(s,f'{pct}%', x0+0.08, bar_y+0.12, bw-0.1, 0.45, sz=16, bold=True, color=tcolors[i])
    x0+=bw

# Detail columns
details=[
    ('40 %   FLUSO','Full flagship treatment: daily X posts, Instagram as its visual home, influencer program, waitlist-driving campaigns and every major product drop.'),
    ('20 %   CONFIDENTIAL API','Enterprise developer entry point. Technical content: architecture explainers, encrypted inference use cases, developer tutorials and integration proof.'),
    ('10 %   SOTTO','Appears when there is a product story or real user receipt. Encrypted dictation is niche but highly shareable among its audience.'),
    ('30 %   THE REST','Prem brand authority (LinkedIn, press, research), Enclave compliance content, Concierge workflow use cases, Studio for developer depth.'),
]
for i,(title,body) in enumerate(details):
    x=0.55+i*3.22
    box(s, x, 3.45, 3.0, 3.6, fill=WHITE, line=PAPER2, lw=0.5)
    T(s,title, x+0.22, 3.6, 2.6, 0.42, sz=10, bold=True, color=INK)
    T(s,body,  x+0.22, 4.1, 2.65, 2.75, sz=9, color=GRAY)

T(s,'Rebalance the split monthly by counting published posts per product. If any bucket drifts more than ±10 points, adjust the following month\'s calendar.',
  0.55, 7.1, 12.2, 0.28, sz=8.5, color=MID)


# ══════════════════════════════════════════════════════════════
# S6  POSITIONING: PREM
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'04  Positioning: Prem')
box(s, 0, 0.9, 13.334, 2.55, fill=DARK)
T(s,'"Private Super Intelligence"', 0.55, 1.1, 12, 0.65, sz=30, bold=True, color=WHITE)
T(s,'"AI is the most powerful technology of our era. Prem makes it private, verifiable, and sovereign."',
  0.55, 1.8, 11.8, 0.5, sz=12, color=MID, italic=True)
T(s,'The three pillars — from the live site hero', 0.55, 2.42, 5, 0.32, sz=9, bold=True, color=MID)

pillars=[
    ('Private',
     'Runs on the customer\'s infrastructure via Prem Enclave. Data never leaves their VPC, never reaches Prem servers. Not a contractual promise — a technical architecture.',
     'Loss-aversion frame: "Policy promises cannot survive a subpoena."'),
    ('Verifiable',
     'Confidential API uses encrypted inference: the model processes data inside a trusted execution environment. Customers can audit what ran, on which model, when.',
     'Proof frame: "Auditable by your team, your counsel, your regulator."'),
    ('Sovereign',
     'Any open model, any infrastructure, customer\'s jurisdiction governs. Prem does not dictate the model, the location or the data handling terms.',
     'Autonomy frame: "You own the stack. We help you run it."'),
]
for i,(t,b,frame) in enumerate(pillars):
    x=0.55+i*4.27
    box(s, x, 3.7, 3.95, 3.4, fill=WHITE, line=PAPER2, lw=0.5)
    T(s,t, x+0.25, 3.88, 3.45, 0.42, sz=17, bold=True, color=INK)
    T(s,b, x+0.25, 4.38, 3.45, 1.8, sz=9.5, color=GRAY)
    Rule(s, x+0.25, 6.3, 3.45, color=PAPER2)
    T(s,frame, x+0.25, 6.42, 3.45, 0.55, sz=8.5, color=GRAY, italic=True)

T(s,'PROOF BANK  (site-verified):  Trusted by NVIDIA · AWS · Microsoft · Index Ventures · Innosuisse · SUPSI · Abu Dhabi Investment Office · Plug and Play · Breyer Capital',
  0.55, 7.1, 12.2, 0.3, sz=8.5, bold=True, color=MID)
NTS(s,'Positioning copy uses the site\'s exact language from the live homepage. Trusted-by logos read directly from the page. Do not extend this list without confirming the relationship is still active.')


# ══════════════════════════════════════════════════════════════
# S7  POSITIONING: FLUSO
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'05  Positioning: Fluso  (the flagship)')
T(s,'"Work deeper, not longer."', 0.55, 0.9, 12, 0.8, sz=34, bold=True, color=INK)
T(s,'"Private AI workspace" — premai.io/products',  0.55, 1.75, 12, 0.35, sz=10.5, color=MID)
Rule(s, 0.55, 2.25, 12.24)
rows=[
    ('FOR','Lawyers, analysts, researchers and engineers doing cognitive-heavy, context-dependent work'),
    ('WHO NEED','AI that integrates into deep-work sessions without breaking concentration or creating its own overhead'),
    ('UNLIKE','Chat-first tools optimised for quick single-turn answers, not sustained multi-context reasoning'),
    ('FLUSO IS','A private AI workspace with compounding memory and 50+ connectors. Designed for 10 to 30 minute focused sessions with a full session audit trail.'),
    ('THE NUMBERS','59 minutes per day lost to information search. 60% of knowledge-worker time consumed by work about work. Fluso targets both.'),
    ('RIGHT NOW','Waitlist stage. "Sign up to Fluso waitlist" is the only conversion goal. Every piece of Fluso content ends here.'),
]
for i,(l,b) in enumerate(rows):
    y=2.48+i*0.82
    T(s,l, 0.55, y, 1.55, 0.55, sz=8.5, bold=True, color=MID)
    T(s,b, 2.25, y, 10.5, 0.72, sz=11, color=INK)
dot_sig(s, 12.4, 7.1)


# ══════════════════════════════════════════════════════════════
# S8  POSITIONING: OTHER PRODUCTS
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'06  Positioning: other products  (verified site copy)')
T(s,'Every product gets a single sentence. Use the site\'s own words.', 0.55, 0.9, 12.2, 0.4, sz=13, bold=True, color=INK)

prods=[
    ('Confidential API','Build with encrypted inference',
     'The developer and enterprise entry point. Encrypted inference means the model never sees raw data in plaintext. Target: engineers building private AI applications into their products.\nConversion goal: API access / demo request.'),
    ('Prem Enclave','Verifiable infrastructure',
     'The hardware layer that makes "private" a technical fact, not a policy. Confidential computing inside a trusted execution environment. Target: infrastructure and security engineers.\nConversion goal: technical evaluation, demo.'),
    ('Prem Concierge','Secure daily workflow assistant',
     'AI assistant for recurring enterprise workflows — research, summarisation, drafting — with the same privacy guarantees as the rest of the stack. Target: enterprise knowledge workers.\nConversion goal: enterprise pilot.'),
    ('Prem Studio','Customize models on private data',
     'Fine-tuning and model customisation on private data. No data leaves the environment during training. Target: ML teams inside regulated organisations.\nConversion goal: technical evaluation.'),
    ('Prem Sotto','Encrypted dictation for every app',
     'Voice-to-text with encrypted inference. Dictate into any application without sending audio to a third-party cloud. Target: professionals who dictate — lawyers, doctors, executives.\nConversion goal: download / waitlist.'),
]
for i,(name,tag,body) in enumerate(prods):
    col=i%2; row=i//2
    x=0.55+col*6.5; y=1.52+row*1.98
    box(s, x, y, 6.14, 1.8, fill=WHITE, line=PAPER2, lw=0.5)
    T(s,name, x+0.22, y+0.1, 4.0, 0.36, sz=13, bold=True, color=INK)
    T(s,tag.upper(), x+0.22, y+0.48, 5.6, 0.26, sz=7.5, bold=True, color=MID)
    T(s,body, x+0.22, y+0.78, 5.7, 0.95, sz=9, color=GRAY)
NTS(s,'Product names and one-line descriptions from live premai.io navigation, verified 9 Jun 2026. Extended descriptions based on logical inference from descriptions — confirm each with product team before publishing externally.')


# ══════════════════════════════════════════════════════════════
# S9  THE CATEGORY
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'07  The category we are creating')
box(s, 0, 0.88, 13.334, 2.25, fill=DARK)
T(s,'Verifiable AI', 0.55, 1.05, 10, 0.82, sz=42, bold=True, color=WHITE)
T(s,'"Verifiable" is already in the company\'s hero copy. The category move is naming the column and owning it.',
  0.55, 1.92, 11.8, 0.38, sz=11, color=MID)
cats=[
    ('Policy-based Privacy',
     'A contractual promise about data handling.\n\nBreaks when terms change, a subpoena arrives, or the provider is acquired.\n\nRequires trust. Cannot be audited.', False),
    ('Perimeter Security',
     'Firewalls and access controls around the system.\n\nDescribes the boundary, not what happened inside the model during inference.\n\nProves the perimeter, not the behaviour.', False),
    ('Verifiable AI',
     'Encrypted inference inside a trusted execution environment.\n\nThe model processes data in hardware-isolated compute. Cryptographically auditable by the customer.\n\nNo trust required. Verify directly.', True),
]
for i,(t,b,hi) in enumerate(cats):
    x=0.55+i*4.27
    bg = DARK if hi else WHITE
    bc = PAPER2 if not hi else None
    box(s, x, 3.4, 3.95, 3.65, fill=bg, line=bc, lw=0.5)
    if not hi: box(s, x, 3.4, 3.95, 0.04, fill=MID)
    T(s,t, x+0.25, 3.58, 3.45, 0.5, sz=14, bold=True, color=WHITE if hi else GRAY)
    T(s,b, x+0.25, 4.18, 3.45, 2.6, sz=10, color=MID if hi else GRAY)
    if hi: T(s,'PREM OWNS THIS COLUMN', x+0.25, 6.7, 3.45, 0.28, sz=7.5, bold=True, color=LIGHT)
T(s,'The category claim converts "private" from a feature into a market position. It is defensible because the technical architecture (confidential computing / TEE) is verifiably distinct from policy-based alternatives.',
  0.55, 7.1, 12.2, 0.28, sz=8.5, color=MID)


# ══════════════════════════════════════════════════════════════
# S10  VOICE: FOUR PRINCIPLES
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'08  Voice and tone  |  Four principles  (every brand surface)')
ps=[
    ('Precise over poetic',
     '"Reduces audit prep from 3 weeks to 4 days" beats "dramatically accelerates compliance." Use numbers when you have them. Name the exact thing.'),
    ('Useful over impressive',
     'Every piece should leave the reader able to do something they could not do before. If it reads like a capability tour, rewrite it.'),
    ('Specific over sweeping',
     'No "many companies" and no "most organisations." Say who, where, how many. If you cannot source it, do not claim it.'),
    ('Honest about limits',
     'State what the product does not do. Admit what the data does not show. The brands that acknowledge constraints earn more trust than the ones that don\'t.'),
]
for i,(t,b) in enumerate(ps):
    x=0.55+(i%2)*6.5; y=1.15+(i//2)*2.98
    box(s, x, y, 6.08, 2.75, fill=WHITE, line=PAPER2, lw=0.5)
    T(s,str(i+1), x+0.28, y+0.18, 0.65, 0.55, sz=26, bold=True, color=LIGHT)
    T(s,t, x+0.28, y+0.82, 5.5, 0.52, sz=17, bold=True, color=INK)
    T(s,b, x+0.28, y+1.42, 5.5, 1.2, sz=10.5, color=GRAY)


# ══════════════════════════════════════════════════════════════
# S11  BANNED LIST
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'08  Voice and tone  |  The banned list')
T(s,'Prohibited on every Prem and Fluso surface. No exceptions, no discretion.', 0.55, 0.9, 12.2, 0.35, sz=11, color=GRAY)
bcols=[
    ('BANNED WORDS',[
        'Revolutionary, game-changing, effortless',
        'Unlock your potential / empower your team',
        '"AI-powered" (everything is)',
        '"The future of work"',
        '"We\'re excited to announce"',
        'Industry-leading, world-class (unverifiable)',
    ]),
    ('BANNED STRUCTURES',[
        'Em dashes (—) anywhere in copy',
        'Rhetorical hooks: "What if your AI could..."',
        'Fragment staccato: "Faster. Smarter. Better."',
        'The "not X, but Y" construction throughout',
        'Aphoristic closers: "That\'s the future we\'re building."',
        'Ending any post with "What do you think?"',
    ]),
    ('BANNED CLAIMS',[
        '"Swiss jurisdiction" for customer-VPC deployments',
        '"Privacy-first" as the primary positioning claim',
        'Any accuracy figure without a cited source',
        'Implying competitors are dishonest — show data',
        'Any certification the company does not currently hold',
        'Statements about our own site no one has reviewed',
    ]),
]
for i,(t,items) in enumerate(bcols):
    x=0.55+i*4.27
    box(s, x, 1.42, 3.95, 5.7, fill=WHITE, line=PAPER2, lw=0.5)
    box(s, x, 1.42, 3.95, 0.04, fill=FN_RED)
    T(s,t, x+0.22, 1.56, 3.5, 0.32, sz=9, bold=True, color=FN_RED)
    for j,item in enumerate(items):
        y=2.06+j*0.85
        T(s,'×', x+0.22, y, 0.28, 0.38, sz=13, bold=True, color=FN_RED)
        T(s,item, x+0.55, y, 3.15, 0.78, sz=9.5, color=INK)


# ══════════════════════════════════════════════════════════════
# S12  VOICE IN ACTION
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'08  Voice and tone  |  In practice')
exs=[
    ('LINKEDIN — PREM PRODUCT INTRO',
     'Introducing Prem — the revolutionary AI platform that empowers your team to unlock new levels of productivity while keeping data safe.',
     'Prem runs on your infrastructure. The Confidential API processes inference inside a trusted execution environment. Your compliance team can audit what ran, on which model, and when.'),
    ('X / FLUSO — PRODUCT POST',
     'What if your AI could actually prove it\'s private? That\'s not a dream. That\'s Fluso.',
     'Fluso produced a full session audit trail for every task in our legal workflow. Our CISO asked how that was possible. I showed them the architecture.'),
    ('PARTNERSHIP ANNOUNCEMENT — LINKEDIN',
     'We\'re excited to announce our partnership with [Partner]. Together we\'re shaping the future of enterprise AI.',
     '[Partner] is deploying Prem Enclave across 400 compliance workflows — the first regulated-industry deployment at this scale with encrypted inference. Implementation notes are on the blog.'),
]
for i,(ctx,bad,good) in enumerate(exs):
    y=1.12+i*2.06
    T(s,ctx, 0.55, y, 12.2, 0.28, sz=8, bold=True, color=MID)
    box(s, 0.55, y+0.32, 5.95, 1.6, fill=FN_REDT, line=None)
    box(s, 0.55, y+0.32, 5.95, 0.04, fill=FN_RED)
    T(s,'DON\'T', 0.78, y+0.42, 1, 0.24, sz=7.5, bold=True, color=FN_RED)
    T(s,bad, 0.78, y+0.68, 5.5, 1.1, sz=9.5, color=INK)
    box(s, 6.78, y+0.32, 5.95, 1.6, fill=FN_GRNT, line=None)
    box(s, 6.78, y+0.32, 5.95, 0.04, fill=FN_GRN)
    T(s,'DO', 7.0, y+0.42, 1, 0.24, sz=7.5, bold=True, color=FN_GRN)
    T(s,good, 7.0, y+0.68, 5.5, 1.1, sz=9.5, color=INK)


# ══════════════════════════════════════════════════════════════
# S13  COLOR
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'09  Color  |  Monochrome, like the mark')
T(s,'The logo is black dots on white. The brand follows.', 0.55, 0.9, 12, 0.45, sz=16, bold=True, color=INK)
T(s,'Color enters through photography (the Alps, the Lugano lake, real product screens). The identity itself stays monochrome.',
  0.55, 1.4, 12, 0.32, sz=10.5, color=GRAY)

sw=[('Ink','#0A0A0A',INK,'The logo color.\nAll marks and headlines.'),
    ('Dark','#16161A',DARK,'Dark surfaces,\nhero blocks, quote cards.'),
    ('Gray','#585A5E',GRAY,'Body secondary text,\nall captions and labels.'),
    ('Mid','#878882',MID,'Rules, footnotes,\nhalftone usage.'),
    ('Paper','#FCFCFA',PAPER,'Default background.\nThe Swiss white space.')]
for i,(n,hx,c,u) in enumerate(sw):
    x=0.55+i*2.52
    b=box(s, x, 2.0, 2.3, 2.45, fill=c)
    if n=='Paper': b.line.color.rgb=PAPER2; b.line.width=Pt(0.8)
    T(s,n, x, 4.58, 2.3, 0.32, sz=12, bold=True, color=INK)
    T(s,hx, x, 4.94, 2.3, 0.28, sz=9, color=GRAY)
    T(s,u, x, 5.28, 2.3, 0.85, sz=8.5, color=GRAY)

box(s, 0.55, 6.3, 12.24, 0.88, fill=WHITE, line=PAPER2, lw=0.5)
T(s,'TWO RULES', 0.78, 6.45, 2.5, 0.28, sz=8, bold=True, color=MID)
T(s,'1.  Color comes from photography and product UI, never from the identity system.   2.  Any Fluso accent colors must be taken directly from the Fluso logo file. Do not approximate, tint, or invent them. If the file has not been checked, Fluso runs monochrome too.',
  0.78, 6.72, 11.7, 0.38, sz=9.5, color=INK)


# ══════════════════════════════════════════════════════════════
# S14  TYPOGRAPHY
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'10  Typography  |  The Prem wordmark typeface')
T(s,'The wordmark uses a geometric grotesque sans-serif.', 0.55, 0.9, 7.5, 0.5, sz=18, bold=True, color=INK)
T(s,'For all presentations, social and digital content, use Pretendard — the closest verified match. Prem self-hosts Pretendard Regular, SemiBold and Bold at static.premai.io/fonts.',
  0.55, 1.48, 7.2, 0.6, sz=10.5, color=GRAY)
rows_t=[
    ('Display / H1','Pretendard Bold  ·  36 to 48pt  ·  leading 1.2x'),
    ('Section head','Pretendard SemiBold  ·  20 to 28pt'),
    ('Body copy','Pretendard Regular  ·  10 to 12pt  ·  leading 1.6x'),
    ('Labels / caps','Pretendard SemiBold  ·  8 to 9pt  ·  all caps  ·  letterspaced +5%'),
    ('Data callout','Pretendard Bold at display scale  ·  ink on paper'),
    ('Code / terminal','JetBrains Mono  ·  for product UI and developer content only'),
]
for i,(st,sp) in enumerate(rows_t):
    y=2.25+i*0.66
    T(s,st, 0.55, y, 2.2, 0.52, sz=10, bold=True, color=INK)
    T(s,sp, 2.85, y, 4.5, 0.52, sz=9.5, color=GRAY)

box(s, 7.85, 0.9, 5.0, 5.55, fill=WHITE, line=PAPER2, lw=0.5)
T(s,'Aa', 8.15, 1.1, 4.4, 1.65, sz=84, bold=True, color=INK)
T(s,'Pretendard Bold', 8.15, 2.95, 4.2, 0.36, sz=12, bold=True, color=INK)
T(s,'ABCDEFGHIJKLM\nabcdefghijklm\n0123456789 !?', 8.15, 3.38, 4.2, 1.3, sz=14, color=GRAY)
T(s,'Geometric grotesque sans — same family as the wordmark.\nFiles: -Regular  -SemiBold  -Bold  .woff2',
  8.15, 5.05, 4.2, 1.12, sz=8.5, color=LIGHT)

box(s, 0.55, 6.38, 7.05, 0.78, fill=WHITE, line=PAPER2, lw=0.5)
T(s,'RULES  ·  Bold for emphasis, never italics. Minimum 10pt digital. Max 70 characters per body line. Sentence case in product and social. All-caps labels only at 8–9pt with +5% letter-spacing.',
  0.78, 6.58, 6.6, 0.5, sz=9, color=INK)
NTS(s,'Pretendard files confirmed in premAI-io/static.premai.io/fonts: Pretendard-Regular.woff2, Pretendard-SemiBold.woff2, Pretendard-Bold.woff2. Wordmark letterforms are geometric grotesque — if the brand team confirms a separate custom typeface for the wordmark specifically, this slide should be updated with that name.')


# ══════════════════════════════════════════════════════════════
# S15  SWISS VISUAL WORLD
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'11  The Swiss visual world')
T(s,'The site already lives here. Social extends it.', 0.55, 0.9, 12.2, 0.55, sz=20, bold=True, color=INK)
motifs=[
    ('Dark Alpine peaks',
     'The hero image on premai.io: monumental mountains, mist, pre-dawn light. This is the power register. Used for: product launch covers, category-claim posts, enterprise authority content.'),
    ('Lugano lake view',
     '"Designed & powered in Switzerland" uses the daytime city-and-lake aerial. Soft blue, hazy, warm. Used for: company story, team content, the Swiss identity, and Instagram background frames.'),
    ('The dot map',
     'Switzerland drawn in the logo\'s own dot language. Our most ownable graphic: the country and the mark in a single image. Use as watermark on carousels, as a full-bleed frame on every 9th Instagram post.'),
    ('The Swiss grid',
     'International Typographic Style is literally Swiss: strict grid, flush-left type, maximum white space. Prem inherits it by birthright. The design decision is almost always: add more white space.'),
]
for i,(t,b) in enumerate(motifs):
    x=0.55+(i%2)*6.5; y=1.72+(i//2)*2.05
    box(s, x, y, 6.08, 1.9, fill=WHITE, line=PAPER2, lw=0.5)
    box(s, x, y, 0.05, 1.9, fill=INK)
    T(s,t, x+0.25, y+0.14, 5.6, 0.42, sz=14, bold=True, color=INK)
    T(s,b, x+0.25, y+0.62, 5.55, 1.15, sz=9.5, color=GRAY)

box(s, 0.55, 5.88, 12.24, 1.28, fill=DARK)
T(s,'RULES', 0.8, 6.02, 2, 0.28, sz=8, bold=True, color=LIGHT)
T(s,'Never use the Swiss flag or cross: protected mark and a cliche.  One motif per asset — they support the message, never carry it.  Color enters through photography only.  Every asset passes the calm test: if it shouts, it ships nowhere.',
  0.8, 6.32, 11.7, 0.72, sz=10, color=MID)


# ══════════════════════════════════════════════════════════════
# S16  SOCIAL TEMPLATES
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s, PAPER)
HDR(s,'12  Visual application  |  Social templates')

# Reel 9:16
rx,ry,rw,rh = 0.85,1.28,2.62,4.66
box(s,rx,ry,rw,rh,fill=PAPER,line=PAPER2,lw=0.8)
# Alpine bg in reel
box(s,rx,ry+2.0,rw,rh-2.0,fill=R1)
ridge(s,[(rx,ry+rh-1.1),(rx+0.8,ry+rh-1.7),(rx+1.6,ry+rh-1.15),(rx+rw,ry+rh-1.55)],R3)
ridge(s,[(rx,ry+rh-0.7),(rx+1.0,ry+rh-1.1),(rx+2.0,ry+rh-0.85),(rx+rw,ry+rh-1.1)],R2)
box(s,rx-0.02,ry+rh,rw+0.04,0.15,fill=PAPER)
T(s,'23 MIN', rx+0.2,ry+0.42,2.2,0.62,sz=28,bold=True,color=INK)
T(s,'contract review,\nevery clause sourced', rx+0.2,ry+1.12,2.2,0.68,sz=10,color=GRAY)
dot_sig(s, rx+0.3,ry+rh-0.36,fill=INK)
T(s,'REEL COVER  9:16', rx,ry+rh+0.12,rw,0.28,sz=8,bold=True,color=GRAY)
T(s,'Big number. One claim.\nRidge footer. Dot signature.',rx,ry+rh+0.42,rw,0.48,sz=8.5,color=LIGHT)

# Carousel 1:1
cx,cy,cw = 4.55,1.28,3.62
box(s,cx,cy,cw,cw,fill=WHITE,line=PAPER2,lw=0.8)
ch_map(s,cx+1.7,cy+2.05,1.75,sp=0.082,r=0.020,fill_fn=lambda t:lerp((175,175,170),(100,100,95),t),lugano=False)
T(s,'The 20-minute\nagent test', cx+0.25,cy+0.28,3.1,0.98,sz=18,bold=True,color=INK)
T(s,'One task. One timer.\nReceipts inside.',cx+0.25,cy+1.35,3.1,0.6,sz=10,color=GRAY)
dot_sig(s,cx+0.25,cy+cw-0.42)
T(s,'CAROUSEL  1:1', cx,cy+cw+0.12,cw,0.28,sz=8,bold=True,color=GRAY)
T(s,'Standalone slide 1. Dot map watermark.\nSummary close — never a CTA slide.',cx,cy+cw+0.42,cw,0.5,sz=8.5,color=LIGHT)

# Quote card 16:9
qx,qy,qw,qh = 8.95,1.28,3.82,2.15
box(s,qx,qy,qw,qh,fill=DARK)
T(s,'"Audit what ran,\non which model, when."',qx+0.25,qy+0.28,3.3,0.82,sz=12.5,bold=True,color=WHITE)
T(s,'Confidential API — session audit trail',qx+0.25,qy+1.52,3.3,0.3,sz=8.5,color=MID)
dot_sig(s,qx+0.38,qy+qh-0.36,fill=WHITE)
T(s,'QUOTE CARD  16:9',qx,qy+qh+0.12,qw,0.28,sz=8,bold=True,color=GRAY)
T(s,'Dark ground. White type.\nDot signature.',qx,qy+qh+0.42,qw,0.48,sz=8.5,color=LIGHT)

# Lugano frame 16:9
lx,ly,lw,lh = 8.95,4.0,3.82,2.15
for (c_,y_,h_) in [(SKY3,ly,lh),(SKY2,ly+0.6,lh-0.6),(SKY,ly+1.0,lh-1.0)]:
    box(s,lx,y_,lw,h_,fill=c_)
ridge(s,[(lx,ly+lh-0.9),(lx+1.0,ly+lh-1.4),(lx+2.0,ly+lh-0.95),(lx+lw,ly+lh-1.3)],R3)
ridge(s,[(lx,ly+lh-0.55),(lx+1.2,ly+lh-0.85),(lx+2.4,ly+lh-0.6),(lx+lw,ly+lh-0.8)],R2)
T(s,'Designed &\npowered in Switzerland',lx+0.25,ly+0.3,3.3,0.78,sz=13,bold=True,color=INK)
dot_sig(s,lx+0.28,ly+lh-0.38)
T(s,'LUGANO LAKE FRAME  16:9',lx,ly+lh+0.12,lw,0.28,sz=8,bold=True,color=GRAY)
T(s,'Company story, Swiss identity,\nwarmth and aspiration register.',lx,ly+lh+0.42,lw,0.5,sz=8.5,color=LIGHT)

box(s,0.55,6.88,12.24,0.52,fill=WHITE,line=PAPER2,lw=0.5)
T(s,'SYSTEM RULE   Paper or dark grounds only. One motif per asset. Type does the talking. The dot signature row appears on every template — it is the family mark.',
  0.78,7.02,11.7,0.32,sz=9,color=INK)


# ══════════════════════════════════════════════════════════════
# S17  LINKEDIN
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
HDR(s,'13  Platform: LinkedIn  (Prem-led · enterprise authority)')
box(s,0.55,1.02,3.98,6.12,fill=WHITE,line=PAPER2,lw=0.5)
T(s,'STRATEGY',0.78,1.15,3.5,0.28,sz=8,bold=True,color=MID)
li_s=[
    ('Audience','CISOs, CTOs, compliance officers, legal ops heads, investors and procurement'),
    ('Goal','Own the answer to "which enterprise AI survives a regulator\'s question?"'),
    ('Voice','Prem the company and its products — not Fluso exclusively. Fluso appears as evidence, not as the program.'),
    ('Cadence','3 to 4 posts per week. Reply to every comment on the first day.'),
    ('KPI','Branded search growth, qualified inbound mentions, SQL conversion. Never followers or likes.'),
]
for i,(l,v) in enumerate(li_s):
    y=1.6+i*1.08
    T(s,l.upper(),0.78,y,3.5,0.26,sz=7.5,bold=True,color=MID)
    T(s,v,0.78,y+0.28,3.6,0.78,sz=9,color=INK)

box(s,4.78,1.02,7.98,6.12,fill=WHITE,line=PAPER2,lw=0.5)
T(s,'FIVE ARCHETYPES  (rotate — never repeat two of the same in a row)',4.98,1.15,7.4,0.28,sz=8,bold=True,color=MID)
archs=[
    ('1  Regulation, decoded',
     'Quote the clause. Plain-English requirement. The gap most deployers miss. One action step.',
     '"Article 26 of the EU AI Act contains one sentence most deployers have not read."'),
    ('2  Architecture teardown',
     'Real problem, how encrypted inference solves it in four steps, what to ask any vendor.',
     '"Here is what happens in the 300 milliseconds after you hit enter on a Confidential API call."'),
    ('3  Named evidence',
     'Named customer, quantified workload, before-and-after numbers, quote with name and title.',
     '"400 compliance workflows, measured."'),
    ('4  The data take',
     'A common belief, the dataset that contradicts it, what it changes for enterprise buyers.',
     '"42% of companies abandoned most AI projects in 2025. The cause is not model quality."'),
    ('5  Operator memo',
     'A decision we faced, the options, what we chose, what it cost. No other ad format builds this trust.',
     '"We turned down a seven-figure deal. The reason should be public."'),
]
for i,(t,st,hook) in enumerate(archs):
    y=1.6+i*1.08
    T(s,t,4.98,y,2.7,0.92,sz=10,bold=True,color=INK)
    T(s,st,7.62,y,3.1,1.0,sz=8,color=GRAY)
    T(s,hook,10.65,y,1.9,1.0,sz=8,color=MID,italic=True)


# ══════════════════════════════════════════════════════════════
# S18  X
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
HDR(s,'14  Platform: X  (Fluso-led · builder credibility)')
box(s,0.55,1.02,3.98,6.12,fill=WHITE,line=PAPER2,lw=0.5)
T(s,'STRATEGY',0.78,1.15,3.5,0.28,sz=8,bold=True,color=MID)
x_s=[
    ('Audience','Builders, developers, AI researchers, power users already debating open-source models and privacy'),
    ('Goal','Be the product builders recommend to each other. Credibility earned, not purchased.'),
    ('Voice','A person who uses Fluso daily. Never a brand-account voice. Never promotional tone.'),
    ('Cadence','1 to 2 posts per day maximum. Replies outperform posts — spend half the time in others\' threads.'),
    ('KPI','Qualified signups and waitlist conversions from profile clicks. Not impressions.'),
]
for i,(l,v) in enumerate(x_s):
    y=1.6+i*1.08
    T(s,l.upper(),0.78,y,3.5,0.26,sz=7.5,bold=True,color=MID)
    T(s,v,0.78,y+0.28,3.6,0.78,sz=9,color=INK)

box(s,4.78,1.02,7.98,6.12,fill=WHITE,line=PAPER2,lw=0.5)
T(s,'FIVE ARCHETYPES',4.98,1.15,7.4,0.28,sz=8,bold=True,color=MID)
x_archs=[
    ('1  Build log',
     'What shipped, one screenshot, one number. No adjectives whatsoever.',
     '"Session export shipped. 14 seconds from audit request to PDF."'),
    ('2  Workflow receipt',
     'Task, time, model, result. The receipt is the post. Nothing else needed.',
     '"Contract review. 23 minutes. Every clause sourced."'),
    ('3  Honest limit',
     'What Fluso cannot do yet, with the current workaround. Outperforms launch posts every time.',
     '"Fluso still struggles with scanned tables. Here is what we do instead."'),
    ('4  Model take',
     'Opinionated open-model comparison from our own runs. Data attached. No winner hype.',
     '"We ran the same brief through four open models. The spread surprised us."'),
    ('5  Community ask',
     'A real open question we have not solved internally. Respond to every single reply.',
     '"How do you handle citation checking in 30-minute research sessions? Genuinely asking."'),
]
for i,(t,st,hook) in enumerate(x_archs):
    y=1.6+i*1.08
    T(s,t,4.98,y,2.7,0.92,sz=10,bold=True,color=INK)
    T(s,st,7.62,y,3.1,1.0,sz=8,color=GRAY)
    T(s,hook,10.65,y,1.9,1.0,sz=8,color=MID,italic=True)


# ══════════════════════════════════════════════════════════════
# S19  INSTAGRAM + FACEBOOK
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
HDR(s,'15  Platform: Instagram + Facebook')
T(s,'Instagram: Show the work.   Facebook: warm the pipe.', 0.55, 0.9, 12.2, 0.45, sz=16, bold=True, color=INK)

box(s,0.55,1.55,5.95,5.62,fill=WHITE,line=PAPER2,lw=0.5)
T(s,'INSTAGRAM  (Fluso-weighted)',0.78,1.68,5.5,0.3,sz=9,bold=True,color=MID)
ig=[
    ('Reels (8 of 15)','15 to 30 seconds. Screen capture of a real task. Text overlay only. Cover frame from the template system (slide 12). First 3 seconds decide 60% of the completion rate.'),
    ('Carousels (7 of 15)','Slide 1 must carry a standalone claim. Dot-map watermark bottom-right. Close on a summary slide, never a CTA slide.'),
    ('Grid rhythm','Alternate paper-ground and dark-ground cards. Every 9th post: the dot map or the Lugano lake frame, full bleed.'),
    ('Name field','"Fluso | deep work AI" — search-indexed. Keywords live here, not in hashtags.'),
    ('Captions','Specific hook on line 1. 3 to 5 hashtags in the first comment, never in the caption. The Dec 2024 algorithm deweighted hashtags; keyword relevance and Reels completion now drive reach.'),
]
for i,(t,b) in enumerate(ig):
    y=2.15+i*1.0
    T(s,t,0.78,y,5.5,0.3,sz=10,bold=True,color=INK)
    T(s,b,0.78,y+0.32,5.55,0.62,sz=8.5,color=GRAY)

box(s,6.8,1.55,5.98,5.62,fill=WHITE,line=PAPER2,lw=0.5)
T(s,'FACEBOOK  (both brands)',7.02,1.68,5.5,0.3,sz=9,bold=True,color=MID)
fb=[
    ('Role','Warm remarketing and community. Organic reach for pages is near zero. Do not build for it.'),
    ('Content','Repurpose top LinkedIn posts 24 to 48 hours later with a Facebook-specific first sentence.'),
    ('About section','The Facebook About section is keyword-indexed for search. Write it as one dense, accurate paragraph covering all products and the company mission.'),
    ('Retargeting','Warm: site visitors and email list. For cold outreach, use LinkedIn. Facebook ads perform on warm audiences.'),
    ('Channel matrix','LinkedIn: enterprise authority · X: builder credibility · Instagram: process visibility · Facebook: warm remarketing and community'),
]
for i,(t,b) in enumerate(fb):
    y=2.15+i*1.0
    T(s,t,7.02,y,5.6,0.3,sz=10,bold=True,color=INK)
    T(s,b,7.02,y+0.32,5.65,0.62,sz=8.5,color=GRAY)

T(s,'Website guidance is deliberately absent from this document. No recommendations without a page-by-page review of the live site.',
  0.55,7.1,12.2,0.28,sz=8.5,color=MID)


# ══════════════════════════════════════════════════════════════
# S20  CONTENT PILLARS
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
HDR(s,'16  Content pillars  |  What each focus bucket publishes')
T(s,'Every piece of content belongs to one pillar. If it fits none, it does not ship.',
  0.55,0.9,12.2,0.35,sz=11,color=GRAY)
pillars_all=[
    ('FLUSO  40%','Deep work and focus sessions, workflow receipts (time + task + result), product utility with honest accuracy, model transparency, "what I actually did in 20 minutes" format, waitlist-driving CTAs', True),
    ('CONFIDENTIAL API  20%','Encrypted inference explainers, developer tutorials and integration walkthroughs, architecture teardowns, regulatory use cases for the API, technical comparison with policy-based alternatives', False),
    ('SOTTO  10%','Encrypted dictation use cases (legal, medical, executive), product updates and new-app support, workflow productivity receipts, honest comparison with non-private voice tools', False),
    ('PREM BRAND  15%','Private Super Intelligence positioning, company story and Lugano identity, trust and security content (premai.io/security), trusted-by proof and partner announcements, founder thought leadership', False),
    ('ENCLAVE, CONCIERGE, STUDIO  15%','Prem Enclave: verifiable infrastructure for regulated enterprise. Prem Concierge: workflow productivity for enterprise teams. Prem Studio: model customisation for ML teams. Each appears when there is a real story.', False),
]
for i,(title,body,flag) in enumerate(pillars_all):
    y=1.42+i*1.2
    bg=DARK if flag else WHITE
    tc=WHITE if flag else INK
    gc=MID if flag else GRAY
    box(s,0.55,y,12.24,1.1,fill=bg,line=None if flag else PAPER2,lw=0.5)
    T(s,title,0.78,y+0.1,3.0,0.38,sz=10,bold=True,color=LIGHT if flag else MID)
    T(s,body, 3.8,y+0.1,9.0,0.9,sz=9.5,color=gc)


# ══════════════════════════════════════════════════════════════
# S21  AUDIENCE ARCHITECTURE
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
HDR(s,'17  Audience architecture  |  Persona to platform')
box(s,0.55,1.02,12.24,0.46,fill=DARK)
ac=[('PERSONA',0.72,2.15),('LEAD PRODUCT',2.95,1.55),('CHANNEL',4.55,1.65),('CONTENT TYPE',6.25,3.05),('JOB TO BE DONE',9.35,3.6)]
for (h,xx,ww) in ac:
    T(s,h,xx,1.1,ww,0.32,sz=7.5,bold=True,color=WHITE)
ar=[
    ('Compliance owner (CISO)','Prem / Enclave','LinkedIn, direct','EU AI Act breakdowns, encrypted inference architecture, audit frameworks','"Help me not get fired when the regulator asks how we deploy AI."'),
    ('Champion (CTO, VP Eng)','Confidential API','LinkedIn, GitHub','Architecture teardowns, API integration guides, private inference proof','"Give me technical evidence I can put in front of the board."'),
    ('Validator (sec engineer)','Prem Enclave','GitHub, X','TEE explainers, benchmarks, verifiable infrastructure walkthroughs','"Show me the architecture, not the marketing page."'),
    ('Vertical buyer (legal ops)','Fluso','LinkedIn, events','Cost and accuracy vs legal AI incumbents, ROI evidence, case studies','"Match the incumbents on accuracy. Beat them on price and compliance."'),
    ('Developer / builder','Confidential API + Studio','X, GitHub, Docs','API docs, model customisation guides, encrypted inference tutorials','"Let me evaluate it in under an hour without talking to sales."'),
    ('Prosumer / power user','Fluso','Instagram, X','Deep-work content, 20-minute workflows, session audit trail demos','"Help me do 3 hours of thinking in 45 minutes."'),
    ('Executive (voice-heavy)','Sotto','LinkedIn, events','Encrypted dictation use cases, productivity ROI, app compatibility','"I dictate into everything. I need it to stay private."'),
]
for i,row in enumerate(ar):
    bg=WHITE if i%2==0 else PAPER2
    box(s,0.55,1.48+i*0.86,12.24,0.86,fill=bg)
    for j,((h,xx,ww),cell) in enumerate(zip(ac,row)):
        T(s,cell,xx,1.56+i*0.86,ww,0.74,sz=8.5,bold=(j==0),color=INK)


# ══════════════════════════════════════════════════════════════
# S22  VOCABULARY
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
HDR(s,'18  Vocabulary  |  Words we own · words we never use')
vcols=[
    ('PREM + PRODUCTS OWN',['Private Super Intelligence','Intelligence in Service','Verifiable','Encrypted inference','On-premise / sovereign','"At the hardware layer"','Confidential computing','Trusted Execution Environment','"Your jurisdiction governs"','Compounding memory','Session audit trail','"Designed & powered in Switzerland"']),
    ('FLUSO OWNS',['Deep work','Focused session','"20-minute session"','Compounding memory','Honest accuracy','Private AI workspace','Open model','"No black box"','Cognitive load','Workflow, not chat','"Work that needs thinking"','Waitlist, then early access']),
    ('NEITHER USES',['Revolutionary, game-changing','Effortless, seamless','"AI-powered" (everything is)','"The future of work"','Unlock, empower','"We\'re excited to announce"','"Privacy-first" as lead claim','Em dashes in copy','"Not X, but Y"','Rhetorical question hooks','Fragment staccato','Uncited accuracy figures']),
]
for i,(t,words) in enumerate(vcols):
    x=0.55+i*4.27
    ac=INK if i<2 else FN_RED
    box(s,x,1.05,3.95,6.07,fill=WHITE,line=PAPER2,lw=0.5)
    box(s,x,1.05,3.95,0.04,fill=ac)
    T(s,t,x+0.22,1.18,3.5,0.3,sz=8.5,bold=True,color=ac if i<2 else FN_RED)
    for j,w in enumerate(words):
        pre='×  ' if i==2 else ''
        T(s,pre+w,x+0.22,1.62+j*0.45,3.55,0.4,sz=9.5,color=INK if i<2 else FN_RED)


# ══════════════════════════════════════════════════════════════
# S23  POSITIONING TEST
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
HDR(s,'19  The positioning test  |  Six questions before every campaign')
T(s,'Run before every campaign brief, product page draft and major social push.',
  0.55,0.9,12.2,0.35,sz=11,color=GRAY)
tests=[
    ('Target customer','Who exactly? Not "enterprises." CISOs and legal ops leads at regulated firms managing EU AI Act compliance before the August 2025 grace period.'),
    ('Market category','Verifiable AI — not "enterprise AI assistant", not "private LLM", not "secure platform."'),
    ('Unique attributes','Encrypted inference inside a trusted execution environment. The customer\'s compliance team can audit inference. No competitor offers both.'),
    ('Value to the buyer','A cryptographically auditable trail answers the regulator directly. A policy promise cannot.'),
    ('Competitive alternatives','Policy-based SaaS tools, perimeter security stacks, vertical legal AI incumbents (Harvey, Legora for legal).'),
    ('Proof of claim','The architecture is public. The Confidential API documentation, Enclave technical paper, and named enterprise deployments with audit trail on record.'),
]
for i,(l,b) in enumerate(tests):
    y=1.35+i*1.0
    bg=WHITE if i%2==0 else PAPER2
    box(s,0.55,y,12.24,0.93,fill=bg,line=None)
    T(s,f'{i+1}',0.78,y+0.2,0.5,0.52,sz=17,bold=True,color=LIGHT)
    T(s,l,1.42,y+0.12,2.85,0.72,sz=10.5,bold=True,color=INK)
    T(s,b,4.42,y+0.12,8.2,0.75,sz=9.5,color=GRAY)
dot_sig(s,12.4,7.1)


# ══════════════════════════════════════════════════════════════
# S24  GOVERNANCE
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
HDR(s,'20  Governance  |  Sign-off by role · no exceptions')
T(s,'Approval = sign-off before publication. It is not an editing right.',0.55,0.9,12.2,0.35,sz=11,color=GRAY)
box(s,0.55,1.38,12.24,0.44,fill=DARK)
gc=[('CONTENT TYPE',0.75,3.45),('APPROVING ROLE',4.25,2.65),('WHY',6.95,5.8)]
for (h,xx,ww) in gc:
    T(s,h,xx,1.46,ww,0.3,sz=8,bold=True,color=WHITE)
gr=[
    ('Press and media statements','CEO','Anything attributed to the company or its leadership.'),
    ('Technical and architecture claims','CTO','Accuracy figures, TEE claims, certifications. Verify before publishing.'),
    ('Public pricing mentions','Sales lead + CEO','Pricing errors are costly and impossible to quietly retract.'),
    ('Campaign strategy and content calendars','Marketing lead','Briefs, channel plans, calendars and influencer briefs.'),
    ('Platform bios and handle changes','Marketing lead + CEO','Bios govern search discovery. Both sign off.'),
    ('Brand visual changes','CEO + design review','Color, logo or type modifications require a consistency audit first.'),
    ('Influencer content','Marketing lead; CEO above 50K reach','Approve the facts. The creator keeps their voice.'),
    ('Crisis or regulatory response','CEO + legal counsel','Regulatory inquiries, data incidents, significant public criticism.'),
]
for i,(ct,ap,why) in enumerate(gr):
    bg=WHITE if i%2==0 else PAPER2
    box(s,0.55,1.82+i*0.65,12.24,0.65,fill=bg)
    T(s,ct,0.75,1.9+i*0.65,3.45,0.52,sz=9,bold=True,color=INK)
    T(s,ap,4.25,1.9+i*0.65,2.65,0.52,sz=9,color=GRAY)
    T(s,why,6.95,1.9+i*0.65,5.75,0.52,sz=8.5,color=GRAY)


# ══════════════════════════════════════════════════════════════
# S25  BACK COVER
# ══════════════════════════════════════════════════════════════
s=NS(); add_bg(s,PAPER)
# Lugano lake full bleed bottom
for (c_,y_,h_) in [(SKY3,3.8,3.7),(SKY2,4.4,3.1),(SKY,5.0,2.5)]:
    box(s,0,y_,13.334,h_,fill=c_)
ridge(s,[(0,4.85),(1.8,4.1),(3.6,4.65),(5.4,3.9),(7.2,4.5),(9.0,3.95),(10.8,4.6),(12.5,4.15),(13.334,4.55)],R3)
ridge(s,[(0,5.55),(2.2,4.95),(4.4,5.45),(6.6,4.8),(8.8,5.35),(11.0,4.92),(13.334,5.3)],R2)
box(s,0,0,13.334,3.9,fill=PAPER)

dot_mark(s, 0.9, 0.72, scale=0.62)
T(s,'PREM',1.5,0.55,5,0.48,sz=28,bold=True,color=INK)
Rule(s,0.55,1.18,4.5,color=INK)
T(s,'One rule above all others:',0.55,1.38,11,0.42,sz=13,color=GRAY)
T(s,'Be consistently more useful\nthan you are impressive.',0.55,1.9,11.5,1.45,sz=32,bold=True,color=INK)
T(s,'The brand is a promise kept in every post, every document, every conversation. Hold the standard.',
  0.55,3.45,10.5,0.55,sz=10.5,color=GRAY)
T(s,'Intelligence in Service',0.55,6.75,6,0.32,sz=10,bold=True,color=GRAY)
T(s,'Prem Brand Bible  v4.0   June 2026   Internal use only',0.55,7.1,8,0.28,sz=8.5,color=LIGHT)
dot_sig(s,12.4,7.1,fill=GRAY)

prs.save('/home/user/Cadence-Architecture-Deck/PremAI-Fluso-Brand-Bible.pptx')
print(f'Done: {len(prs.slides)} slides')
