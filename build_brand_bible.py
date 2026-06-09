"""PremAI + Fluso Brand Bible v2 - Swiss design system."""
import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def rgb(r, g, b): return RGBColor(r, g, b)

# ── Verified palette (logo.svg, static.premai.io, fetched 9 Jun 2026) ──
PERI   = rgb(127, 150, 255)   # 7F96FF gradient stop 0
CORAL  = rgb(245, 142, 142)   # F58E8E gradient stop 0.5
SAND   = rgb(242, 211, 152)   # F2D398 gradient stop 1
# Neutrals (from existing PremAI deck system)
PAPER  = rgb(250, 249, 245)   # FAF9F5
PAPER2 = rgb(236, 234, 225)   # ECEAE1
INK    = rgb(23, 23, 27)      # 17171B
GRAY   = rgb(91, 95, 99)      # 5B5F63
LGRAY  = rgb(154, 154, 147)   # 9A9A93
WHITE  = rgb(255, 255, 255)
# Tints for ridgelines
PERI_T  = rgb(225, 231, 255)
CORAL_T = rgb(252, 230, 230)
SAND_T  = rgb(250, 241, 222)
PERI_T2 = rgb(206, 215, 252)
# Functional only (never brand)
FN_RED   = rgb(186, 60, 60)
FN_GREEN = rgb(46, 125, 80)
FN_RED_T   = rgb(250, 235, 235)
FN_GREEN_T = rgb(233, 245, 238)

FONT = 'Pretendard'

def lerp(c1, c2, t):
    return rgb(int(c1[0]+(c2[0]-c1[0])*t), int(c1[1]+(c2[1]-c1[1])*t), int(c1[2]+(c2[2]-c1[2])*t))

def grad(t):
    """Logo gradient: PERI at 0, CORAL at 0.506, SAND at 1."""
    a, b, c = (127,150,255), (245,142,142), (242,211,152)
    if t <= 0.506: return lerp(a, b, t/0.506)
    return lerp(b, c, (t-0.506)/0.494)

# ── Geometry helpers ───────────────────────────────────────
def add_bg(slide, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.34), Inches(7.5))
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s

def box(slide, left, top, w, h, fill=None, line=None, lw=0.75):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(w), Inches(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = fill
    else: s.fill.background()
    if line: s.line.color.rgb = line; s.line.width = Pt(lw)
    else: s.line.fill.background()
    return s

def dot(slide, cx, cy, r, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx-r), Inches(cy-r), Inches(2*r), Inches(2*r))
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s

def txt(slide, text, left, top, w, h, sz=11, bold=False, color=INK,
        align=PP_ALIGN.LEFT, spacing=None):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    f = r.font; f.size = Pt(sz); f.bold = bold; f.color.rgb = color; f.name = FONT
    return tb

def label(slide, text, left, top, w=8, color=GRAY):
    """Small-caps style section label."""
    return txt(slide, text.upper(), left, top, w, 0.3, sz=9, bold=True, color=color)

def rule(slide, left, top, w, color=INK, weight=0.012):
    return box(slide, left, top, w, weight, fill=color)

def dot_row(slide, left, top, n=5, r=0.045, gap=0.16):
    for i in range(n):
        dot(slide, left + i*gap, top, r, grad(i/(n-1)))

def ridge(slide, pts, fill):
    """Filled mountain silhouette. pts: list of (x_in, y_in) for the ridgeline,
    closed down to the bottom edge of the slide."""
    fb = slide.shapes.build_freeform(Emu(Inches(pts[0][0])), Emu(Inches(7.5)), scale=1.0)
    seq = [(Inches(x), Inches(y)) for x, y in pts] + [(Inches(pts[-1][0]), Inches(7.5))]
    fb.add_line_segments([(Emu(x), Emu(y)) for x, y in seq], close=True)
    sh = fb.convert_to_shape()
    sh.fill.solid(); sh.fill.fore_color.rgb = fill; sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

# ── Switzerland dot map ────────────────────────────────────
CH_POLY = [
    (5.96,46.14),(6.06,46.45),(6.45,46.78),(6.43,46.93),(6.95,47.25),
    (7.35,47.43),(7.59,47.58),(8.20,47.62),(8.57,47.80),(9.05,47.68),
    (9.56,47.54),(9.67,47.06),(10.47,46.86),(10.39,46.63),(10.04,46.22),
    (9.28,46.32),(9.04,45.82),(8.44,46.25),(7.86,45.92),(7.04,45.90),
    (6.80,46.15),(6.80,46.43),(6.24,46.34),(6.30,46.25),
]
COS = math.cos(math.radians(46.8))

def in_poly(x, y, poly):
    inside = False
    j = len(poly) - 1
    for i in range(len(poly)):
        xi, yi = poly[i]; xj, yj = poly[j]
        if (yi > y) != (yj > y) and x < (xj-xi)*(y-yi)/(yj-yi)+xi:
            inside = not inside
        j = i
    return inside

def swiss_dot_map(slide, left, top, width, spacing=0.115, r=0.036,
                  color_fn=None, mark_lugano=True):
    lons = [p[0] for p in CH_POLY]; lats = [p[1] for p in CH_POLY]
    lon0, lon1 = min(lons), max(lons); lat0, lat1 = min(lats), max(lats)
    ew = (lon1 - lon0) * COS          # effective width in degrees
    eh = lat1 - lat0
    scale = width / ew
    height = eh * scale
    nx = int(width / spacing); ny = int(height / spacing)
    for iy in range(ny + 1):
        for ix in range(nx + 1):
            px = ix * spacing; py = iy * spacing
            lon = lon0 + (px / scale) / COS
            lat = lat1 - (py / scale)
            if in_poly(lon, lat, CH_POLY):
                t = (px / width) * 0.75 + (py / height) * 0.25
                c = color_fn(t) if color_fn else grad(t)
                dot(slide, left + px, top + py, r, c)
    if mark_lugano:
        lx = left + (8.95 - lon0) * COS * scale
        ly = top + (lat1 - 46.0) * scale
        dot(slide, lx, ly, 0.055, INK)
        txt(slide, 'Lugano HQ', lx + 0.12, ly - 0.12, 1.2, 0.25, sz=8, bold=True, color=INK)
    return height

# ── Build ──────────────────────────────────────────────────
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]
def ns(): return prs.slides.add_slide(blank)

def header(slide, section, page=None):
    label(slide, section, 0.55, 0.42)
    rule(slide, 0.55, 0.78, 12.23, color=INK)

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ══════════════════════════════════════════════════════════
# S1  COVER
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
ridge(s, [(0,6.30),(1.6,5.62),(3.2,6.18),(4.8,5.35),(6.4,6.10),(8.0,5.50),(9.6,6.18),(11.2,5.72),(13.34,6.28)], PERI_T)
ridge(s, [(0,6.72),(2.2,6.05),(4.4,6.62),(6.6,5.95),(8.8,6.52),(11.0,6.12),(13.34,6.62)], CORAL_T)
ridge(s, [(0,7.05),(2.6,6.55),(5.2,6.98),(7.8,6.42),(10.4,6.92),(13.34,6.62)], SAND_T)

swiss_dot_map(s, 7.35, 1.05, 5.3)

dot_row(s, 0.62, 1.18)
txt(s, 'THE BRAND BIBLE', 0.55, 1.5, 6.6, 1.0, sz=44, bold=True, color=INK)
txt(s, 'Prem + Fluso', 0.57, 2.55, 6.5, 0.6, sz=24, color=GRAY)
rule(s, 0.57, 3.35, 3.2, color=INK)
txt(s, 'Built in Lugano. Verified by design.', 0.57, 3.55, 6.2, 0.4, sz=12, color=GRAY)
txt(s, 'Brand architecture  |  Voice  |  Color and type  |  The Swiss visual world  |  Platform playbooks',
    0.57, 4.05, 6.4, 0.6, sz=9.5, color=LGRAY)
txt(s, 'v2.0   June 2026   Internal use only', 0.57, 4.65, 6, 0.35, sz=9, color=LGRAY)
notes(s, 'Logo gradient verified from logo.svg at static.premai.io (premAI-io/static.premai.io, fetched 9 Jun 2026): #7F96FF, #F58E8E, #F2D398. Lugano HQ address per public company information: Crocicchio Cortogna 6, 6900 Lugano.')


# ══════════════════════════════════════════════════════════
# S2  CONTENTS
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, 'Contents')
sections = [
    ('01','Brand architecture','One family, two identities'),
    ('02','Positioning: Prem','Enterprise AI you can prove'),
    ('03','Positioning: Fluso','Work deeper, not longer'),
    ('04','The category','Verifiable AI'),
    ('05','Voice and tone','Four principles, one banned list'),
    ('06','Color','The verified logo palette'),
    ('07','Typography','Pretendard, the brand typeface'),
    ('08','The Swiss visual world','Mountains, the dot map, calm space'),
    ('09','Visual application','Social templates in practice'),
    ('10','Platform playbooks','LinkedIn, X, Instagram, Facebook'),
    ('11','Content pillars','What each brand publishes'),
    ('12','Audience architecture','Persona to platform'),
    ('13','Vocabulary and governance','Words we own, who signs off'),
]
for i, (num, title, sub) in enumerate(sections):
    col = i // 7
    row = i % 7
    x = 0.55 + col * 6.4
    y = 1.15 + row * 0.86
    txt(s, num, x, y, 0.55, 0.35, sz=11, bold=True, color=GRAY)
    txt(s, title, x+0.62, y, 5.0, 0.35, sz=13, bold=True, color=INK)
    txt(s, sub, x+0.62, y+0.38, 5.2, 0.32, sz=9, color=LGRAY)
dot_row(s, 12.35, 7.05)


# ══════════════════════════════════════════════════════════
# S3  BRAND ARCHITECTURE
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '01  Brand architecture')
txt(s, 'One family. Two distinct identities.', 0.55, 1.0, 11, 0.7, sz=28, bold=True, color=INK)

box(s, 0.55, 1.95, 5.85, 4.95, fill=WHITE)
box(s, 0.55, 1.95, 5.85, 0.06, fill=PERI)
txt(s, 'PREM', 0.85, 2.12, 3.5, 0.5, sz=22, bold=True, color=INK)
txt(s, 'Masterbrand. The enterprise identity.', 0.85, 2.66, 5.2, 0.32, sz=9.5, color=GRAY)
prem_rows = [
    ('Audience','CISOs, CTOs, compliance officers, enterprise procurement'),
    ('Tone','Authoritative, precise, institutional. Peer-to-peer CTO voice.'),
    ('Channels','LinkedIn primary. Press, partner and investor materials.'),
    ('Purpose','Build the category Verifiable AI. Win regulated enterprise.'),
    ('Mark','Prem dot-cluster mark, gradient version, on light backgrounds.'),
]
for i, (l, v) in enumerate(prem_rows):
    y = 3.15 + i * 0.72
    txt(s, l.upper(), 0.85, y, 1.25, 0.3, sz=7.5, bold=True, color=GRAY)
    txt(s, v, 2.15, y, 4.05, 0.62, sz=9.5, color=INK)

box(s, 6.95, 1.95, 5.85, 4.95, fill=WHITE)
box(s, 6.95, 1.95, 5.85, 0.06, fill=CORAL)
txt(s, 'FLUSO', 7.25, 2.12, 3.5, 0.5, sz=22, bold=True, color=INK)
txt(s, 'Product sub-brand. The user-facing identity.', 7.25, 2.66, 5.2, 0.32, sz=9.5, color=GRAY)
fluso_rows = [
    ('Audience','Lawyers, analysts, researchers, developers, prosumers'),
    ('Tone','Direct, builder voice, grounded, honest about limitations'),
    ('Channels','X primary. Instagram, product blog, community.'),
    ('Purpose','Be the product builders recommend to each other.'),
    ('Mark','Fluso dot mark, used exactly as provided. Never recolored.'),
]
for i, (l, v) in enumerate(fluso_rows):
    y = 3.15 + i * 0.72
    txt(s, l.upper(), 7.25, y, 1.25, 0.3, sz=7.5, bold=True, color=GRAY)
    txt(s, v, 8.55, y, 4.05, 0.62, sz=9.5, color=INK)

txt(s, 'The lockup is "Fluso by Prem", never "Prem\'s product". The sub-brand keeps its own personality while borrowing the masterbrand\'s enterprise trust.',
    0.55, 7.02, 12.3, 0.4, sz=9.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S4  POSITIONING: PREM
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '02  Positioning: Prem')
txt(s, '"Enterprise AI you can prove."', 0.55, 1.0, 12, 0.8, sz=32, bold=True, color=INK)
txt(s, 'The only AI platform that ships cryptographic attestation with every inference. A proof, not a policy.',
    0.55, 1.85, 11.5, 0.4, sz=12, color=GRAY)

pillars = [
    ('Proof per inference',
     'Reticle produces a hardware-signed cryptographic attestation at the TEE layer on every inference. Verifiable by any party. No competitor ships this.', PERI),
    ('Model portability',
     'Any open-source model on your own infrastructure. Switch models without migration cost. No lock-in, no forced upgrade path.', CORAL),
    ('Sovereign deployment',
     'Data never leaves the customer VPC. The customer\'s jurisdiction governs. Auditable by their team, their counsel, their regulator.', SAND),
]
for i, (t, b, accent) in enumerate(pillars):
    x = 0.55 + i * 4.27
    box(s, x, 2.55, 3.95, 4.35, fill=WHITE)
    box(s, x, 2.55, 3.95, 0.06, fill=accent)
    dot(s, x+0.42, 3.05, 0.09, accent)
    txt(s, t, x+0.28, 3.35, 3.4, 0.65, sz=16, bold=True, color=INK)
    txt(s, b, x+0.28, 4.15, 3.4, 2.6, sz=10.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S5  POSITIONING: FLUSO
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '03  Positioning: Fluso')
txt(s, '"Work deeper, not longer."', 0.55, 1.0, 12, 0.8, sz=32, bold=True, color=INK)
txt(s, 'Deep-work AI for knowledge professionals who think in 20-minute sessions, not prompt chains.',
    0.55, 1.85, 11.5, 0.4, sz=12, color=GRAY)
rule(s, 0.55, 2.5, 12.23, color=PAPER2, weight=0.02)
details = [
    ('FOR','Lawyers, analysts, researchers and engineers doing cognitive-heavy work'),
    ('WHO NEED','AI that fits into deep-work sessions without breaking concentration or creating new overhead'),
    ('UNLIKE','Chat-first tools optimised for quick answers rather than sustained reasoning sessions'),
    ('FLUSO IS','A private AI workspace with compounding memory, built for 10 to 30 minute focused workflows, with session audit trail and honest accuracy reporting'),
    ('THE PROOF','59 minutes per day lost to information search. 60% of knowledge-worker time spent on work about work. Fluso attacks both numbers.'),
]
for i, (l, b) in enumerate(details):
    y = 2.75 + i * 0.86
    txt(s, l, 0.55, y, 1.6, 0.5, sz=9, bold=True, color=GRAY)
    txt(s, b, 2.3, y, 10.4, 0.75, sz=11.5, color=INK)
dot_row(s, 12.35, 7.1)


# ══════════════════════════════════════════════════════════
# S6  THE CATEGORY
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '04  The category we are creating')
txt(s, 'Verifiable AI', 0.55, 1.0, 10, 0.85, sz=40, bold=True, color=INK)
txt(s, 'Not "Private AI". Not "Secure AI". A new category that Prem names and owns.',
    0.55, 1.9, 11.5, 0.4, sz=12, color=GRAY)
cats = [
    ('Policy-based "Private AI"',
     'A contractual promise that the provider handles data properly.\n\nBreaks when terms change or a subpoena arrives.\n\nCannot be audited. Must be trusted.', False),
    ('Perimeter-based "Secure AI"',
     'Firewalls and access controls around the system\'s exterior.\n\nSays nothing about what happened inside the model during inference.\n\nProves the boundary, not the behaviour.', False),
    ('Verifiable AI',
     'Hardware-signed cryptographic proof per inference, via the open-source Reticle stack.\n\nMathematically checkable by any party.\n\nNo trust required.', True),
]
for i, (t, b, hi) in enumerate(cats):
    x = 0.55 + i * 4.27
    box(s, x, 2.55, 3.95, 4.35, fill=WHITE if not hi else INK)
    box(s, x, 2.55, 3.95, 0.06, fill=LGRAY if not hi else PERI)
    txt(s, t, x+0.28, 2.72, 3.4, 0.65, sz=13, bold=True, color=GRAY if not hi else WHITE)
    txt(s, b, x+0.28, 3.45, 3.4, 3.1, sz=10, color=GRAY if not hi else PAPER2)
    if hi:
        txt(s, 'PREM OWNS THIS COLUMN', x+0.28, 6.5, 3.4, 0.3, sz=8, bold=True, color=PERI)


# ══════════════════════════════════════════════════════════
# S7  VOICE: FOUR PRINCIPLES
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '05  Voice and tone  |  Four principles, both brands')
principles = [
    ('Precise over poetic',
     'Name the exact thing. Use numbers when you have them. "Reduces audit prep from 3 weeks to 4 days" beats "dramatically accelerates compliance."'),
    ('Useful over impressive',
     'Every piece should leave the reader able to do something they could not do before. If it reads like a capability tour, rewrite it.'),
    ('Specific over sweeping',
     'No "many companies". No "most organisations". Say who, where, how many. If you cannot source it, do not claim it.'),
    ('Honest about limits',
     'State what the product does not do. Admit what the data does not show. Readers trust brands that acknowledge constraints.'),
]
for i, (t, b) in enumerate(principles):
    x = 0.55 + (i % 2) * 6.4
    y = 1.15 + (i // 2) * 2.95
    box(s, x, y, 5.85, 2.7, fill=WHITE)
    box(s, x, y, 0.06, 2.7, fill=grad(i/3))
    txt(s, str(i+1), x+0.3, y+0.22, 0.6, 0.5, sz=22, bold=True, color=LGRAY)
    txt(s, t, x+0.3, y+0.8, 5.2, 0.5, sz=17, bold=True, color=INK)
    txt(s, b, x+0.3, y+1.35, 5.25, 1.25, sz=10.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S8  BANNED LIST
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '05  Voice and tone  |  The banned list')
txt(s, 'Prohibited across all Prem and Fluso content. No exceptions.', 0.55, 0.95, 12, 0.35, sz=11, color=GRAY)
cols = [
    ('BANNED WORDS', [
        'Revolutionary, game-changing, effortless',
        'Unlock your potential, empower your team',
        '"AI-powered" (everything is AI-powered)',
        '"The future of work"',
        '"We\'re excited to announce"',
        '"Industry-leading" and other unverifiable superlatives',
    ]),
    ('BANNED STRUCTURES', [
        'Em dashes anywhere in copy',
        'Rhetorical hooks: "What if your AI could..."',
        'Fragment staccato: "Faster. Smarter. Better."',
        'The "not X, but Y" construction',
        'Aphoristic closers: "That\'s the future we\'re building."',
        'Ending posts with "What do you think?"',
    ]),
    ('BANNED CLAIMS', [
        '"Swiss jurisdiction" for customer-VPC deployments',
        '"Privacy-first" as the lead positioning claim',
        'Any accuracy figure without a named source',
        'Implying competitors are dishonest. Show data instead.',
        'Any certification the company does not currently hold',
        'Statements about our own website or product UI no one has checked',
    ]),
]
for i, (t, items) in enumerate(cols):
    x = 0.55 + i * 4.27
    box(s, x, 1.4, 3.95, 5.7, fill=WHITE)
    box(s, x, 1.4, 3.95, 0.06, fill=FN_RED)
    txt(s, t, x+0.25, 1.55, 3.5, 0.35, sz=9.5, bold=True, color=FN_RED)
    for j, item in enumerate(items):
        y = 2.05 + j * 0.83
        txt(s, '×', x+0.25, y, 0.3, 0.35, sz=12, bold=True, color=FN_RED)
        txt(s, item, x+0.6, y, 3.15, 0.75, sz=9.5, color=INK)


# ══════════════════════════════════════════════════════════
# S9  VOICE IN ACTION
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '05  Voice and tone  |  In practice')
examples = [
    ('LINKEDIN PRODUCT INTRO',
     'Introducing Prem, the revolutionary AI platform that empowers your team to unlock new levels of productivity while keeping data safe.',
     'Prem runs on your infrastructure. Every inference produces a cryptographic proof via Reticle. Your compliance team can audit what ran, when, and on which model.'),
    ('X / FLUSO POST',
     'What if your AI could actually prove it\'s private? That\'s not a dream. That\'s Fluso.',
     'Fluso produced a hardware-signed audit trail for every session in our legal workflow. Our CISO asked how. I sent the Reticle docs.'),
    ('PARTNERSHIP ANNOUNCEMENT',
     'We\'re excited to announce our partnership with [Partner]. Together we\'re shaping the future of enterprise AI.',
     '[Partner] is deploying Prem across 400 compliance workflows, the first regulated-industry deployment at this scale with full attestation. Implementation notes on the blog.'),
]
for i, (ctx, bad, good) in enumerate(examples):
    y = 1.15 + i * 2.0
    txt(s, ctx, 0.55, y, 12, 0.28, sz=8, bold=True, color=GRAY)
    box(s, 0.55, y+0.32, 6.0, 1.52, fill=FN_RED_T)
    txt(s, 'DON\'T', 0.78, y+0.42, 1.0, 0.25, sz=7.5, bold=True, color=FN_RED)
    txt(s, bad, 0.78, y+0.7, 5.55, 1.05, sz=9.5, color=INK)
    box(s, 6.78, y+0.32, 6.0, 1.52, fill=FN_GREEN_T)
    txt(s, 'DO', 7.0, y+0.42, 1.0, 0.25, sz=7.5, bold=True, color=FN_GREEN)
    txt(s, good, 7.0, y+0.7, 5.55, 1.05, sz=9.5, color=INK)


# ══════════════════════════════════════════════════════════
# S10  COLOR
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '06  Color  |  The verified logo palette')
txt(s, 'Every brand color comes from one place: the Prem logo gradient.',
    0.55, 0.95, 12, 0.4, sz=13, bold=True, color=INK)
txt(s, 'Source: logo.svg, static.premai.io. Three gradient stops, verified 9 June 2026.',
    0.55, 1.4, 12, 0.32, sz=9.5, color=GRAY)

swatches = [
    ('Periwinkle', '#7F96FF', PERI,  'Gradient stop 0.\nPrem primary accent.'),
    ('Coral',      '#F58E8E', CORAL, 'Gradient stop 0.5.\nFluso primary accent.'),
    ('Sand',       '#F2D398', SAND,  'Gradient stop 1.\nWarm highlight, sparingly.'),
    ('Ink',        '#17171B', INK,   'All text.\nDark surfaces.'),
    ('Paper',      '#FAF9F5', PAPER, 'Default background.\nThe Swiss white space.'),
]
for i, (n, hx, c, u) in enumerate(swatches):
    x = 0.55 + i * 2.52
    sw = box(s, x, 1.95, 2.28, 2.3, fill=c)
    if n == 'Paper': sw.line.color.rgb = PAPER2; sw.line.width = Pt(1)
    txt(s, n, x, 4.38, 2.28, 0.35, sz=12, bold=True, color=INK)
    txt(s, hx, x, 4.74, 2.28, 0.3, sz=9.5, color=GRAY)
    txt(s, u, x, 5.08, 2.28, 0.85, sz=8.5, color=GRAY)

box(s, 0.55, 6.1, 12.23, 1.0, fill=WHITE)
box(s, 0.55, 6.1, 0.06, 1.0, fill=CORAL)
txt(s, 'FLUSO NOTE', 0.85, 6.22, 2.0, 0.3, sz=8, bold=True, color=CORAL)
txt(s, 'Fluso never uses orange. Its accent colors must be lifted from the Fluso logo file exactly as provided. Until those hex values are pulled from the asset, use coral #F58E8E from the family gradient. Do not approximate, tint, or recolor the dot mark.',
    0.85, 6.5, 11.7, 0.55, sz=9.5, color=INK)
notes(s, 'Gradient stops read directly from logo.svg in premAI-io/static.premai.io: stop 0 #7F96FF, stop 0.505785 #F58E8E, stop 1 #F2D398. Neutrals Paper #FAF9F5 and Ink #17171B taken from the existing PremAI deck system. Fluso exact logo hexes were not available to verify; flagged on slide rather than guessed.')


# ══════════════════════════════════════════════════════════
# S11  TYPOGRAPHY
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '07  Typography  |  Pretendard')
txt(s, 'Pretendard', 0.55, 1.0, 8, 0.9, sz=44, bold=True, color=INK)
txt(s, 'The verified brand typeface. Prem already self-hosts Pretendard Regular, SemiBold and Bold at static.premai.io/fonts. One family, both brands, every surface.',
    0.55, 2.0, 7.0, 0.75, sz=11.5, color=GRAY)

type_rows = [
    ('Display / H1', 'Pretendard Bold, 36 to 48pt, tight leading (1.2x)'),
    ('Section head', 'Pretendard SemiBold, 20 to 28pt'),
    ('Body copy',    'Pretendard Regular, 10 to 12pt, leading 1.6x'),
    ('Labels',       'Pretendard SemiBold, 8 to 9pt, all caps, letterspaced'),
    ('Data callout', 'Pretendard Bold at display scale, accent color'),
]
for i, (st, sp) in enumerate(type_rows):
    y = 3.0 + i * 0.62
    txt(s, st, 0.55, y, 2.1, 0.5, sz=10, bold=True, color=INK)
    txt(s, sp, 2.75, y, 4.8, 0.5, sz=9.5, color=GRAY)

box(s, 8.0, 1.0, 4.78, 5.1, fill=WHITE)
txt(s, 'Aa', 8.35, 1.2, 4.0, 1.5, sz=80, bold=True, color=INK)
txt(s, 'Pretendard Bold', 8.35, 2.95, 4.0, 0.35, sz=11, bold=True, color=INK)
txt(s, 'ABCDEFGHIJKLM\nabcdefghijklm\n0123456789', 8.35, 3.35, 4.1, 1.3, sz=14, color=GRAY)
txt(s, 'Files: Pretendard-Regular.woff2, Pretendard-SemiBold.woff2, Pretendard-Bold.woff2',
    8.35, 4.95, 4.1, 0.9, sz=8.5, color=LGRAY)

box(s, 0.55, 6.3, 12.23, 0.8, fill=WHITE)
txt(s, 'RULES   Bold for emphasis, never italics.  Minimum 10pt digital.  Max 70 characters per line.  Sentence case in product and social.  Mono (JetBrains Mono) only for code.',
    0.85, 6.5, 11.7, 0.45, sz=9.5, color=INK)
notes(s, 'Pretendard verified as the brand typeface: the fonts directory of premAI-io/static.premai.io contains Pretendard-Regular.woff2, Pretendard-SemiBold.woff2, Pretendard-Bold.woff2 (checked 9 Jun 2026). If the logo wordmark itself uses a different custom face, confirm against the master logo file; everything else runs on Pretendard.')


# ══════════════════════════════════════════════════════════
# S12  THE SWISS VISUAL WORLD
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '08  The Swiss visual world')
txt(s, 'Prem is built in Lugano. The visuals say so without saying so.',
    0.55, 0.95, 12.2, 0.6, sz=22, bold=True, color=INK)

motifs = [
    ('Alpine ridgelines',
     'Layered mountain silhouettes in gradient tints. Used as footers, section breaks and Reel cover frames. Always calm, never dramatic stock peaks.'),
    ('The dot map',
     'Switzerland drawn in the logo\'s own dot language, gradient applied. Our most ownable asset: the country and the mark in one image.'),
    ('Lake horizon',
     'Generous negative space above a single low horizon line. White space is the loudest part of the layout.'),
    ('Swiss grid',
     'The International Typographic Style is literally Swiss: strict grid, flush-left type, no decoration. We inherit it by birthright.'),
]
for i, (t, b) in enumerate(motifs):
    x = 0.55 + (i % 2) * 6.4
    y = 1.8 + (i // 2) * 1.95
    box(s, x, y, 5.85, 1.75, fill=WHITE)
    box(s, x, y, 0.06, 1.75, fill=grad(i/3))
    txt(s, t, x+0.28, y+0.16, 5.3, 0.4, sz=13, bold=True, color=INK)
    txt(s, b, x+0.28, y+0.6, 5.3, 1.05, sz=9.5, color=GRAY)

box(s, 0.55, 5.8, 12.23, 1.3, fill=INK)
txt(s, 'RULES', 0.85, 5.95, 2.0, 0.3, sz=8, bold=True, color=PERI)
txt(s, 'Photography: real Alps, muted morning light, no saturated stock.  Never use the Swiss flag or cross: protected mark, and a cliche.  Motifs support the message, they never carry it.  Every asset passes the calm test: if it shouts, it ships nowhere.',
    0.85, 6.25, 11.7, 0.75, sz=10, color=PAPER)


# ══════════════════════════════════════════════════════════
# S13  VISUAL APPLICATION (mock templates)
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '09  Visual application  |  Social templates')

# (a) Reel cover 9:16
rx, ry, rw, rh = 0.85, 1.3, 2.6, 4.62
box(s, rx, ry, rw, rh, fill=PAPER, line=PAPER2, lw=1.2)
ridge(s, [(rx, ry+rh-1.1),(rx+0.7, ry+rh-1.7),(rx+1.4, ry+rh-1.2),(rx+2.0, ry+rh-1.8),(rx+rw, ry+rh-1.25)], PERI_T)
# clip ridge visually by overlaying paper below card
box(s, rx-0.02, ry+rh, rw+0.04, 7.5-(ry+rh), fill=PAPER)
txt(s, '23 MIN', rx+0.2, ry+0.5, 2.2, 0.6, sz=28, bold=True, color=INK)
txt(s, 'contract review,\nevery clause sourced', rx+0.2, ry+1.15, 2.2, 0.7, sz=10, color=GRAY)
dot_row(s, rx+0.32, ry+rh-0.32, n=3, r=0.035, gap=0.13)
txt(s, 'REEL COVER  9:16', rx, ry+rh+0.12, rw, 0.3, sz=8, bold=True, color=GRAY)
txt(s, 'Big number, one claim,\nridge footer, dot row', rx, ry+rh+0.42, rw, 0.5, sz=8.5, color=LGRAY)

# (b) Carousel 1:1
cx, cy, cw = 4.55, 1.3, 3.6
box(s, cx, cy, cw, cw, fill=WHITE, line=PAPER2, lw=1.2)
txt(s, 'The 20-minute\nagent test', cx+0.25, cy+0.3, 3.1, 1.0, sz=18, bold=True, color=INK)
txt(s, 'One task. One timer.\nReceipts inside.', cx+0.25, cy+1.45, 3.1, 0.6, sz=10, color=GRAY)
swiss_dot_map(s, cx+1.55, cy+2.1, 1.8, spacing=0.085, r=0.022,
              color_fn=lambda t: lerp((242,211,152),(245,142,142),t), mark_lugano=False)
txt(s, 'CAROUSEL  1:1', cx, cy+cw+0.12, cw, 0.3, sz=8, bold=True, color=GRAY)
txt(s, 'Standalone claim on slide 1, dot-map watermark, summary close (never a CTA slide)',
    cx, cy+cw+0.42, cw, 0.55, sz=8.5, color=LGRAY)

# (c) Quote / data card 16:9
qx, qy, qw, qh = 9.0, 1.3, 3.8, 2.14
box(s, qx, qy, qw, qh, fill=INK)
txt(s, '"Our CISO asked how.\nI sent the Reticle docs."', qx+0.25, qy+0.3, 3.3, 0.85, sz=12, bold=True, color=WHITE)
txt(s, 'Fluso session audit trail', qx+0.25, qy+1.45, 3.3, 0.3, sz=8.5, color=LGRAY)
dot_row(s, qx+0.37, qy+qh-0.28, n=5, r=0.035, gap=0.13)
txt(s, 'QUOTE CARD  16:9', qx, qy+qh+0.12, qw, 0.3, sz=8, bold=True, color=GRAY)
txt(s, 'Ink ground, white type, gradient dot row as signature', qx, qy+qh+0.42, qw, 0.5, sz=8.5, color=LGRAY)

box(s, 9.0, 4.6, 3.8, 2.3, fill=WHITE)
txt(s, 'SYSTEM RULES', 9.25, 4.75, 3.3, 0.3, sz=8, bold=True, color=GRAY)
txt(s, 'Paper or Ink grounds only. One motif per asset. Type does the talking. The dot row is the family signature across every template.',
    9.25, 5.05, 3.35, 1.7, sz=9.5, color=INK)


# ══════════════════════════════════════════════════════════
# S14  LINKEDIN
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '10  Platform playbooks  |  LinkedIn  (Prem primary)')

box(s, 0.55, 1.05, 4.1, 6.05, fill=WHITE)
box(s, 0.55, 1.05, 4.1, 0.06, fill=PERI)
txt(s, 'STRATEGY', 0.8, 1.18, 3.6, 0.3, sz=8.5, bold=True, color=GRAY)
li_strategy = [
    ('Audience','CISOs, CTOs, compliance officers, legal ops, investors'),
    ('Goal','Become the default answer to "which enterprise AI survives a regulator\'s questions?"'),
    ('Cadence','3 to 4 posts per week, maximum. Quality over volume.'),
    ('KPI','Branded search growth, qualified inbound, SQL mentions. Never followers or likes.'),
    ('Engagement','Reply within 2 hours on publish day. Founder and team comments outperform page posts.'),
]
for i, (l, v) in enumerate(li_strategy):
    y = 1.6 + i * 1.08
    txt(s, l.upper(), 0.8, y, 3.5, 0.28, sz=7.5, bold=True, color=PERI)
    txt(s, v, 0.8, y+0.28, 3.6, 0.78, sz=9, color=INK)

box(s, 4.85, 1.05, 7.93, 6.05, fill=WHITE)
box(s, 4.85, 1.05, 7.93, 0.06, fill=PERI)
txt(s, 'FIVE POST ARCHETYPES  (rotate, never repeat two in a row)', 5.1, 1.18, 7.4, 0.3, sz=8.5, bold=True, color=GRAY)
li_arch = [
    ('1  Regulation, decoded',
     'Quote the clause. Plain-English requirement. The gap most deployers miss. One action step.',
     '"Article 26 of the EU AI Act contains one sentence most deployers have not read."'),
    ('2  Architecture teardown',
     'A real problem, a diagram or screenshot, how attestation works in 4 steps, what to ask any vendor.',
     '"Here is what happens in the 300ms after you hit enter on a Prem inference."'),
    ('3  Named evidence',
     'Customer, workload, before-and-after numbers, a quote with name and title. No anonymous case studies.',
     '"A 400-workflow compliance deployment, measured."'),
    ('4  The data take',
     'A common belief, the dataset that contradicts it, what it changes for buyers.',
     '"42% of companies abandoned most AI initiatives in 2025. Model quality is not the reason."'),
    ('5  Operator memo',
     'A decision we faced, the options, what we chose, what it cost us. Builds trust no ad can buy.',
     '"We turned down a seven-figure deal last quarter. The reason matters."'),
]
for i, (t, st, hook) in enumerate(li_arch):
    y = 1.6 + i * 1.08
    txt(s, t, 5.1, y, 2.6, 0.95, sz=10, bold=True, color=INK)
    txt(s, st, 7.55, y, 3.0, 1.0, sz=8, color=GRAY)
    txt(s, hook, 10.5, y, 2.1, 1.0, sz=8, color=PERI)


# ══════════════════════════════════════════════════════════
# S15  X
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '10  Platform playbooks  |  X  (Fluso primary)')

box(s, 0.55, 1.05, 4.1, 6.05, fill=WHITE)
box(s, 0.55, 1.05, 4.1, 0.06, fill=CORAL)
txt(s, 'STRATEGY', 0.8, 1.18, 3.6, 0.3, sz=8.5, bold=True, color=GRAY)
x_strategy = [
    ('Audience','Builders, developers, AI researchers, power users already debating open models and privacy'),
    ('Goal','Be the product builders recommend to each other. Credibility earned, not bought.'),
    ('Voice','A person who uses Fluso daily. Never a brand account voice.'),
    ('Cadence','1 to 2 posts per day max. Replies outgrow posts; spend half the time in other people\'s threads.'),
    ('KPI','Qualified signups and waitlist conversions from profile clicks. Not impressions.'),
]
for i, (l, v) in enumerate(x_strategy):
    y = 1.6 + i * 1.08
    txt(s, l.upper(), 0.8, y, 3.5, 0.28, sz=7.5, bold=True, color=CORAL)
    txt(s, v, 0.8, y+0.28, 3.6, 0.78, sz=9, color=INK)

box(s, 4.85, 1.05, 7.93, 6.05, fill=WHITE)
box(s, 4.85, 1.05, 7.93, 0.06, fill=CORAL)
txt(s, 'FIVE POST ARCHETYPES', 5.1, 1.18, 7.4, 0.3, sz=8.5, bold=True, color=GRAY)
x_arch = [
    ('1  Build log',
     'What shipped, one screenshot, one number. No adjectives.',
     '"Session export shipped. 14 seconds from audit request to PDF."'),
    ('2  Workflow receipt',
     'Task, time, model used, result. The receipt is the post.',
     '"Contract review. 23 minutes. Every clause sourced."'),
    ('3  Honest limit',
     'What Fluso cannot do yet, plus the current workaround. These outperform launch posts.',
     '"Fluso still struggles with scanned tables. Here is what we do instead."'),
    ('4  Model take',
     'Opinionated open-model comparison backed by our own runs. Data attached, no winner hype.',
     '"We ran the same brief through 4 open models. The spread surprised us."'),
    ('5  Community ask',
     'A real open question we have not solved. Respond to every single answer.',
     '"How do you handle citation checking in long research sessions? Genuinely asking."'),
]
for i, (t, st, hook) in enumerate(x_arch):
    y = 1.6 + i * 1.08
    txt(s, t, 5.1, y, 2.6, 0.95, sz=10, bold=True, color=INK)
    txt(s, st, 7.55, y, 3.0, 1.0, sz=8, color=GRAY)
    txt(s, hook, 10.5, y, 2.1, 1.0, sz=8, color=rgb(200, 95, 95))


# ══════════════════════════════════════════════════════════
# S16  INSTAGRAM
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '10  Platform playbooks  |  Instagram  (Fluso)')
txt(s, 'Show the work. The Swiss visual world lives here.', 0.55, 0.95, 12, 0.5, sz=18, bold=True, color=INK)

box(s, 0.55, 1.65, 6.0, 5.45, fill=WHITE)
box(s, 0.55, 1.65, 6.0, 0.06, fill=SAND)
txt(s, 'CONTENT FORMATS', 0.8, 1.78, 5.5, 0.3, sz=8.5, bold=True, color=GRAY)
ig_fmt = [
    ('Reels, 15 to 30s','Screen capture of a real task plus text overlay. No talking heads, no voiceover. Cover frame: big number, ridge footer (template, slide 09).'),
    ('Carousels','Slide 1 carries a standalone claim. Dot-map watermark bottom-right. Close on a summary slide, never a CTA slide.'),
    ('Stories','Behind-the-work only: real screenshots, drafts, whiteboards. No branded templates. Polls only when we genuinely want the answer.'),
    ('Grid rhythm','Alternate Paper-ground and Ink-ground cards so the grid reads as a pattern. Every 9th post: the dot map or a ridge, full bleed.'),
]
for i, (t, b) in enumerate(ig_fmt):
    y = 2.2 + i * 1.22
    txt(s, t, 0.8, y, 5.4, 0.3, sz=10.5, bold=True, color=INK)
    txt(s, b, 0.8, y+0.32, 5.45, 0.85, sz=8.5, color=GRAY)

box(s, 6.8, 1.65, 6.0, 5.45, fill=WHITE)
box(s, 6.8, 1.65, 6.0, 0.06, fill=SAND)
txt(s, 'DISCOVERY AND CAPTIONS', 7.05, 1.78, 5.5, 0.3, sz=8.5, bold=True, color=GRAY)
ig_disc = [
    ('Name field','"Fluso | deep work AI". The name field is search-indexed; keywords live there, not in hashtags.'),
    ('Bio','Line 1: what it does in 8 words. Line 2: one proof point. Line 3: link. No emoji walls.'),
    ('Captions','First line is the hook and must be specific, not clever. 3 to 5 hashtags in the first comment, never in the caption body.'),
    ('Algorithm reality','Hashtags were deweighted in Dec 2024; keyword search and Reels completion rate drive reach. The first 3 seconds decide most of the completion curve.'),
]
for i, (t, b) in enumerate(ig_disc):
    y = 2.2 + i * 1.22
    txt(s, t, 7.05, y, 5.4, 0.3, sz=10.5, bold=True, color=INK)
    txt(s, b, 7.05, y+0.32, 5.45, 0.85, sz=8.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S17  FACEBOOK + CHANNEL MATRIX
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '10  Platform playbooks  |  Facebook + the channel matrix')

box(s, 0.55, 1.05, 12.23, 1.7, fill=WHITE)
box(s, 0.55, 1.05, 12.23, 0.06, fill=LGRAY)
txt(s, 'FACEBOOK', 0.8, 1.18, 3.0, 0.3, sz=8.5, bold=True, color=GRAY)
txt(s, 'Role: warm remarketing and community, not organic reach. Repurpose top LinkedIn posts after 24 to 48 hours with a Facebook-specific first line. The About section is keyword-indexed for search: write it as one dense, accurate paragraph. Retarget site visitors and the email list; LinkedIn handles cold.',
    0.8, 1.5, 11.7, 1.1, sz=10, color=INK)

# Channel matrix
txt(s, 'THE CHANNEL MATRIX', 0.55, 3.05, 6, 0.3, sz=8.5, bold=True, color=GRAY)
box(s, 0.55, 3.4, 12.23, 0.45, fill=INK)
mcols = [('CHANNEL',0.75,1.5),('BRAND',2.4,1.2),('ROLE',3.7,3.3),('CADENCE',7.1,2.0),('KPI',9.2,3.4)]
for (h, xx, ww) in mcols:
    txt(s, h, xx, 3.48, ww, 0.3, sz=8, bold=True, color=WHITE)
mrows = [
    ('LinkedIn','Prem','Authority and inbound in regulated enterprise','3 to 4 per week','Branded search, SQL mentions'),
    ('X','Fluso','Builder credibility, product proof in public','1 to 2 per day','Qualified signups from profile'),
    ('Instagram','Fluso','Process visibility, the Swiss visual world','4 to 5 per week','Profile-to-site clicks, activations'),
    ('Facebook','Both','Warm remarketing and community','2 to 3 per week, repurposed','Re-engagement CTR'),
]
for i, row in enumerate(mrows):
    bg = WHITE if i % 2 == 0 else PAPER2
    box(s, 0.55, 3.85 + i*0.72, 12.23, 0.72, fill=bg)
    for (h, xx, ww), cell in zip(mcols, row):
        txt(s, cell, xx, 3.97 + i*0.72, ww, 0.55, sz=9.5,
            bold=(xx == 0.75), color=INK)
txt(s, 'Website guidance is deliberately absent: no recommendations until the live site has been reviewed page by page.',
    0.55, 6.95, 12.2, 0.35, sz=9, color=GRAY)


# ══════════════════════════════════════════════════════════
# S18  CONTENT PILLARS: PREM
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '11  Content pillars  |  Prem')
txt(s, 'Every Prem piece belongs to one pillar. If it fits none, it does not ship.',
    0.55, 0.95, 12, 0.35, sz=11, color=GRAY)
p_pillars = [
    ('Verifiable AI','40%','Reticle attestation\nTEE architecture explainers\nProof walkthroughs\nOpen source vs policy promises'),
    ('Compliance architecture','30%','EU AI Act interpretation\nGDPR and HIPAA by use case\nAudit readiness frameworks\nWhat a regulator actually asks'),
    ('Engineering depth','20%','Open-model selection\nOn-premise deployment guides\nPortability benchmarks\nReticle contributor updates'),
    ('Customer evidence','10%','Named case studies only\nQuantified outcomes\nQuotes with name and title\nRegulator-facing results'),
]
for i, (t, pct, topics) in enumerate(p_pillars):
    x = 0.55 + i * 3.22
    box(s, x, 1.45, 2.95, 5.55, fill=WHITE)
    box(s, x, 1.45, 2.95, 0.06, fill=grad(i/3))
    txt(s, pct, x+0.25, 1.65, 2.4, 0.6, sz=26, bold=True, color=INK)
    txt(s, t, x+0.25, 2.35, 2.5, 0.75, sz=13, bold=True, color=INK)
    rule(s, x+0.25, 3.2, 2.45, color=PAPER2, weight=0.02)
    txt(s, topics, x+0.25, 3.35, 2.5, 3.4, sz=9.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S19  CONTENT PILLARS: FLUSO
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '11  Content pillars  |  Fluso')
txt(s, 'Utility leads. 60% of output shows the product doing real work.',
    0.55, 0.95, 12, 0.35, sz=11, color=GRAY)
f_pillars = [
    ('Product utility','35%','Use cases with time data\nBefore and after workflows\nModel-to-task matching\nSession audit trail demos'),
    ('Deep work','25%','Cognitive load in AI workflows\nFocus session design\n10 to 30 minute workflows\nWork vs work-about-work'),
    ('Model transparency','20%','Which model, why, when\nHonest accuracy reporting\nLimitation disclosure\nOpen-model reviews'),
    ('Builder community','20%','Reticle open-source updates\nContributor spotlights\nHonest retrospectives\nOpen questions, answered'),
]
for i, (t, pct, topics) in enumerate(f_pillars):
    x = 0.55 + i * 3.22
    box(s, x, 1.45, 2.95, 5.55, fill=WHITE)
    box(s, x, 1.45, 2.95, 0.06, fill=grad(i/3))
    txt(s, pct, x+0.25, 1.65, 2.4, 0.6, sz=26, bold=True, color=INK)
    txt(s, t, x+0.25, 2.35, 2.5, 0.75, sz=13, bold=True, color=INK)
    rule(s, x+0.25, 3.2, 2.45, color=PAPER2, weight=0.02)
    txt(s, topics, x+0.25, 3.35, 2.5, 3.4, sz=9.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S20  AUDIENCE ARCHITECTURE
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '12  Audience architecture  |  Persona to platform')
box(s, 0.55, 1.05, 12.23, 0.46, fill=INK)
acols = [('PERSONA',0.7,2.2),('BRAND',2.95,1.0),('CHANNEL',4.0,1.8),('CONTENT',5.85,3.1),('JOB TO BE DONE',9.0,3.7)]
for (h, xx, ww) in acols:
    txt(s, h, xx, 1.12, ww, 0.32, sz=7.5, bold=True, color=WHITE)
arows = [
    ('Compliance owner (CISO)','Prem','LinkedIn, direct','EU AI Act breakdowns, attestation explainers, audit frameworks','"Help me not get fired when the regulator asks how we use AI."'),
    ('Champion (CTO, VP Eng)','Prem','LinkedIn, GitHub','Architecture posts, portability data, Reticle technical docs','"Give me technical proof I can put in front of the board."'),
    ('Validator (security eng)','Prem + Fluso','GitHub, X','TEE explainers, benchmarks, open-source contribution posts','"Show me the code, not the marketing page."'),
    ('Vertical buyer (legal ops)','Prem','LinkedIn, events','Cost comparisons vs legal AI tools, ROI evidence, case studies','"Beat the incumbents on price, match them on accuracy."'),
    ('Developer / builder','Fluso','X, GitHub, community','API docs, model selection guides, workflow teardowns','"Let me try it in 10 minutes without talking to sales."'),
    ('Prosumer / power user','Fluso','Instagram, X','Deep-work content, 20-minute workflows, productivity data','"Help me do 3 hours of thinking in 45 minutes."'),
]
for i, row in enumerate(arows):
    bg = WHITE if i % 2 == 0 else PAPER2
    box(s, 0.55, 1.51 + i*0.93, 12.23, 0.93, fill=bg)
    for j, ((h, xx, ww), cell) in enumerate(zip(acols, row)):
        txt(s, cell, xx, 1.58 + i*0.93, ww, 0.82, sz=8.5, bold=(j == 0), color=INK)


# ══════════════════════════════════════════════════════════
# S21  VOCABULARY
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '13  Vocabulary  |  Words we own, words we never use')
vcols = [
    ('PREM OWNS', PERI, ['Verifiable','Attestation','Proof, not promise','On-premise','Sovereign','Auditable','"At the hardware layer"','Hardware-signed','"Per inference"','Reticle (always capitalised)','Model portability','"Your jurisdiction governs"']),
    ('FLUSO OWNS', CORAL, ['Deep work','Focused session','"20-minute session"','Compounding memory','Honest accuracy','Session audit trail','Open model','"No black box"','Cognitive load','Workflow, not chat','"Work that needs thinking"','"By Prem"']),
    ('NEITHER USES', FN_RED, ['Revolutionary, game-changing','Effortless, seamless','"AI-powered"','"The future of work"','Unlock, empower','"We\'re excited to announce"','"Privacy-first" as lead claim','Em dashes in copy','"Not X, but Y"','Rhetorical question hooks','Fragment staccato','Uncited figures']),
]
for i, (t, accent, words) in enumerate(vcols):
    x = 0.55 + i * 4.27
    box(s, x, 1.05, 3.95, 6.05, fill=WHITE)
    box(s, x, 1.05, 3.95, 0.06, fill=accent)
    txt(s, t, x+0.25, 1.18, 3.5, 0.32, sz=9, bold=True, color=accent if accent != PERI else rgb(86,104,200))
    for j, w in enumerate(words):
        pre = '×  ' if i == 2 else ''
        txt(s, pre + w, x+0.25, 1.62 + j*0.45, 3.55, 0.4, sz=9.5,
            color=INK if i < 2 else FN_RED)


# ══════════════════════════════════════════════════════════
# S22  POSITIONING TEST
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '13  The positioning test  |  Run before every campaign')
tests = [
    ('Target customer','Who, exactly? Not "enterprises". CISOs and legal ops leads at regulated firms facing EU AI Act deadlines.'),
    ('Market category','Verifiable AI. Not "enterprise AI assistant", not "private LLM", not "secure AI platform".'),
    ('Unique attributes','Hardware-signed proof per inference via the open-source Reticle stack. No other vendor ships this.'),
    ('Value to the buyer','A cryptographic audit trail answers the regulator directly. A policy promise cannot.'),
    ('Competitive alternatives','Policy-based enterprise tools, perimeter security stacks, vertical legal AI suites.'),
    ('Proof of claim','The Reticle repository, TEE documentation, and named deployments with attestation on record.'),
]
for i, (l, b) in enumerate(tests):
    y = 1.1 + i * 0.97
    bg = WHITE if i % 2 == 0 else PAPER2
    box(s, 0.55, y, 12.23, 0.9, fill=bg)
    txt(s, f'{i+1}', 0.78, y+0.18, 0.5, 0.5, sz=16, bold=True, color=LGRAY)
    txt(s, l, 1.45, y+0.12, 2.8, 0.7, sz=10.5, bold=True, color=INK)
    txt(s, b, 4.4, y+0.12, 8.2, 0.74, sz=9.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S23  GOVERNANCE (roles only, no names)
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '13  Governance  |  Sign-off by role')
txt(s, 'Approval means sign-off before publication, not editing rights.', 0.55, 0.95, 12, 0.35, sz=11, color=GRAY)
box(s, 0.55, 1.4, 12.23, 0.44, fill=INK)
gcols = [('CONTENT TYPE',0.75,3.4),('APPROVING ROLE',4.2,2.7),('WHY',7.0,5.8)]
for (h, xx, ww) in gcols:
    txt(s, h, xx, 1.47, ww, 0.32, sz=8, bold=True, color=WHITE)
grows = [
    ('Press and media statements','CEO','Anything attributed to the company or its leadership.'),
    ('Technical claims and benchmarks','CTO','Accuracy figures, architecture claims, certifications. Verify before publishing.'),
    ('Public pricing mentions','Sales lead + CEO','Pricing errors are costly and hard to retract.'),
    ('Campaign strategy and calendars','Marketing lead','Briefs, channel plans, content calendars, influencer briefs.'),
    ('Platform bios and handles','Marketing lead + CEO','Bios change how the brand is found in search.'),
    ('Brand visual changes','CEO + design review','Color, logo or type changes need a consistency audit first.'),
    ('Influencer content','Marketing lead (CEO above 50K reach)','We approve the facts; creators keep their voice.'),
    ('Crisis or regulatory response','CEO + legal counsel','Regulatory inquiries, incidents, significant public criticism.'),
]
for i, (ct, ap, why) in enumerate(grows):
    bg = WHITE if i % 2 == 0 else PAPER2
    box(s, 0.55, 1.84 + i*0.64, 12.23, 0.64, fill=bg)
    txt(s, ct, 0.75, 1.92 + i*0.64, 3.4, 0.52, sz=9, bold=True, color=INK)
    txt(s, ap, 4.2, 1.92 + i*0.64, 2.7, 0.52, sz=9, color=rgb(86,104,200))
    txt(s, why, 7.0, 1.92 + i*0.64, 5.7, 0.52, sz=8.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S24  BACK COVER
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
ridge(s, [(0,6.45),(2.4,5.85),(4.8,6.4),(7.2,5.7),(9.6,6.35),(12.0,5.95),(13.34,6.4)], PERI_T)
ridge(s, [(0,6.95),(3.3,6.45),(6.6,6.9),(9.9,6.35),(13.34,6.85)], SAND_T)
dot_row(s, 0.62, 1.5)
txt(s, 'One rule above all others:', 0.55, 1.85, 11, 0.5, sz=15, color=GRAY)
txt(s, 'Be consistently more useful\nthan you are impressive.', 0.55, 2.4, 11.5, 1.6, sz=34, bold=True, color=INK)
rule(s, 0.57, 4.35, 3.2, color=INK)
txt(s, 'The brand is a promise kept in every post, every document, every conversation. Hold the standard.',
    0.57, 4.55, 9.5, 0.6, sz=11, color=GRAY)
txt(s, 'Prem + Fluso Brand Bible  v2.0   June 2026   Internal use only', 0.57, 5.25, 9, 0.35, sz=9, color=LGRAY)

prs.save('/home/user/Cadence-Architecture-Deck/PremAI-Fluso-Brand-Bible.pptx')
print(f'Saved {len(prs.slides)} slides')
