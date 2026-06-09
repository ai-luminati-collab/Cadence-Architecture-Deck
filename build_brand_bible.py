"""PremAI + Fluso Brand Bible v3 - monochrome Swiss system, portfolio architecture."""
import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def rgb(r, g, b): return RGBColor(r, g, b)

# ── Monochrome palette (live premai.io logo: black dot mark, verified 9 Jun 2026) ──
INK    = rgb(10, 10, 10)      # near-black, logo color
INK2   = rgb(23, 23, 27)      # dark surface
GRAY   = rgb(91, 95, 99)      # body secondary
MGRAY  = rgb(140, 140, 136)   # mid gray
LGRAY  = rgb(176, 176, 170)   # light gray
PAPER  = rgb(250, 250, 248)   # FAFAF8
PAPER2 = rgb(236, 234, 225)   # ECEAE1 panel
WHITE  = rgb(255, 255, 255)
# Ridge tints (grayscale)
R1 = rgb(225, 224, 218)
R2 = rgb(207, 206, 199)
R3 = rgb(189, 188, 181)
# Functional only (never brand)
FN_RED   = rgb(186, 60, 60)
FN_GREEN = rgb(46, 125, 80)
FN_RED_T   = rgb(250, 235, 235)
FN_GREEN_T = rgb(233, 245, 238)

FONT = 'Pretendard'

def lerp(c1, c2, t):
    return rgb(int(c1[0]+(c2[0]-c1[0])*t), int(c1[1]+(c2[1]-c1[1])*t), int(c1[2]+(c2[2]-c1[2])*t))

def ink_ramp(t):
    """Halftone depth ramp for the dot map: ink to mid gray."""
    return lerp((10, 10, 10), (130, 130, 125), t)

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
        align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    f = r.font; f.size = Pt(sz); f.bold = bold; f.color.rgb = color; f.name = FONT
    return tb

def label(slide, text, left, top, w=10, color=GRAY):
    return txt(slide, text.upper(), left, top, w, 0.3, sz=9, bold=True, color=color)

def rule(slide, left, top, w, color=INK, weight=0.012):
    return box(slide, left, top, w, weight, fill=color)

def dot_row(slide, left, top, color=INK):
    """Family signature: five ink dots of varied size, echoing the logo."""
    radii = [0.052, 0.034, 0.046, 0.03, 0.04]
    x = left
    for r in radii:
        dot(slide, x, top, r, color)
        x += 0.17

def ridge(slide, pts, fill):
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
                  color_fn=ink_ramp, mark_lugano=True):
    lons = [p[0] for p in CH_POLY]; lats = [p[1] for p in CH_POLY]
    lon0, lon1 = min(lons), max(lons); lat0, lat1 = min(lats), max(lats)
    ew = (lon1 - lon0) * COS
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
                t = (px / width) * 0.6 + (py / height) * 0.4
                dot(slide, left + px, top + py, r, color_fn(t))
    if mark_lugano:
        lx = left + (8.95 - lon0) * COS * scale
        ly = top + (lat1 - 46.0) * scale
        d = dot(slide, lx, ly, 0.06, WHITE)
        d.line.color.rgb = INK; d.line.width = Pt(1.5)
        txt(slide, 'Lugano HQ', lx + 0.13, ly - 0.12, 1.2, 0.25, sz=8, bold=True, color=INK)
    return height

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]
def ns(): return prs.slides.add_slide(blank)

def header(slide, section):
    label(slide, section, 0.55, 0.42)
    rule(slide, 0.55, 0.78, 12.23, color=INK)

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ══════════════════════════════════════════════════════════
# S1  COVER
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
ridge(s, [(0,6.30),(1.6,5.62),(3.2,6.18),(4.8,5.35),(6.4,6.10),(8.0,5.50),(9.6,6.18),(11.2,5.72),(13.34,6.28)], R1)
ridge(s, [(0,6.72),(2.2,6.05),(4.4,6.62),(6.6,5.95),(8.8,6.52),(11.0,6.12),(13.34,6.62)], R2)
ridge(s, [(0,7.05),(2.6,6.55),(5.2,6.98),(7.8,6.42),(10.4,6.92),(13.34,6.62)], R3)

swiss_dot_map(s, 7.35, 1.05, 5.3)

dot_row(s, 0.62, 1.18)
txt(s, 'THE BRAND BIBLE', 0.55, 1.5, 6.6, 1.0, sz=44, bold=True, color=INK)
txt(s, 'Prem and its products', 0.57, 2.55, 6.5, 0.6, sz=24, color=GRAY)
rule(s, 0.57, 3.35, 3.2, color=INK)
txt(s, 'Private Super Intelligence, built in Lugano.', 0.57, 3.55, 6.2, 0.4, sz=12, color=GRAY)
txt(s, 'Brand architecture  |  Digital focus  |  Voice  |  Monochrome system  |  The Swiss visual world  |  Platform playbooks',
    0.57, 4.05, 6.6, 0.6, sz=9.5, color=LGRAY)
txt(s, 'v3.0   June 2026   Internal use only', 0.57, 4.65, 6, 0.35, sz=9, color=LGRAY)
notes(s, 'Logo verified from the live premai.io site (screenshot, 9 Jun 2026): monochrome black dot-cluster mark, "PREM" wordmark. The gradient logo.svg in static.premai.io is an older asset and is not the current identity. Site hero: "Private Super Intelligence" over dark alpine photography. Tagline on site: "AI is the most powerful technology of our era. Prem makes it private, verifiable, and sovereign."')


# ══════════════════════════════════════════════════════════
# S2  CONTENTS
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, 'Contents')
sections = [
    ('01','Brand architecture','Prem masterbrand and the product portfolio'),
    ('02','Digital focus','Fluso 40%, Prem and the rest 60%'),
    ('03','Positioning: Prem','Private Super Intelligence'),
    ('04','Positioning: Fluso','Work deeper, not longer'),
    ('05','The category','Verifiable AI'),
    ('06','Voice and tone','Four principles, one banned list'),
    ('07','Color','Monochrome, like the mark'),
    ('08','Typography','Pretendard, the brand typeface'),
    ('09','The Swiss visual world','Mountains, the dot map, calm space'),
    ('10','Visual application','Social templates in practice'),
    ('11','Platform playbooks','LinkedIn, X, Instagram, Facebook'),
    ('12','Content and audience','Pillars and persona mapping'),
    ('13','Vocabulary and governance','Words we own, who signs off'),
]
for i, (num, title, sub) in enumerate(sections):
    col = i // 7
    row = i % 7
    x = 0.55 + col * 6.4
    y = 1.15 + row * 0.86
    txt(s, num, x, y, 0.55, 0.35, sz=11, bold=True, color=GRAY)
    txt(s, title, x+0.62, y, 5.0, 0.35, sz=13, bold=True, color=INK)
    txt(s, sub, x+0.62, y+0.38, 5.4, 0.32, sz=9, color=LGRAY)
dot_row(s, 12.25, 7.05)


# ══════════════════════════════════════════════════════════
# S3  BRAND ARCHITECTURE (portfolio)
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '01  Brand architecture')
txt(s, 'One masterbrand. A portfolio of products.', 0.55, 0.95, 11.5, 0.6, sz=26, bold=True, color=INK)

# Masterbrand band
box(s, 0.55, 1.7, 12.23, 1.25, fill=INK2)
txt(s, 'PREM', 0.9, 1.92, 2.5, 0.5, sz=22, bold=True, color=WHITE)
txt(s, 'Private Super Intelligence. The enterprise identity: authoritative, precise, institutional. Carries the trust that every product borrows. LinkedIn, press, partners, investors.',
    2.6, 1.95, 10.0, 0.85, sz=10.5, color=PAPER2)

# Product cards
cards = [
    ('FLUSO', 'Flagship product',
     'Private AI workspace with compounding memory, 50+ connectors and cross-silo workflows. The deep-work product for knowledge professionals. Own voice: direct, builder, honest.',
     'Lockup: "Fluso by Prem". Dot mark used exactly as provided, never recolored.', True),
    ('PREM STUDIO', 'Platform product',
     'Build, fine-tune, evaluate and automate models. Marketed to developers and ML teams through docs, tutorials and engineering content under the Prem voice.',
     'Always written as two words, both capitalised.', False),
    ('RETICLE', 'Open source',
     'The attestation and verification stack behind Verifiable AI. Community-facing: GitHub, contributor content, technical posts. Credibility engine for the whole portfolio.',
     'Always capitalised. Never described as a feature.', False),
]
for i, (name, kind, body, ruleline, flag) in enumerate(cards):
    x = 0.55 + i * 4.27
    box(s, x, 3.2, 3.95, 3.45, fill=WHITE)
    box(s, x, 3.2, 3.95, 0.06, fill=INK)
    txt(s, name, x+0.25, 3.38, 3.0, 0.4, sz=15, bold=True, color=INK)
    txt(s, kind.upper(), x+0.25, 3.82, 3.4, 0.28, sz=7.5, bold=True, color=GRAY)
    if flag:
        box(s, x+2.5, 3.42, 1.25, 0.34, fill=INK)
        txt(s, 'FLAGSHIP', x+2.5, 3.47, 1.25, 0.26, sz=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(s, body, x+0.25, 4.18, 3.45, 1.7, sz=9.5, color=GRAY)
    rule(s, x+0.25, 6.0, 3.45, color=PAPER2, weight=0.02)
    txt(s, ruleline, x+0.25, 6.1, 3.45, 0.5, sz=8.5, color=LGRAY)

txt(s, 'The portfolio is marketed as one family. Fluso leads, the platform earns, Reticle proves. No product gets the whole stage.',
    0.55, 6.95, 12.2, 0.4, sz=10, color=GRAY)
notes(s, 'Product facts sourced: Fluso description (private AI workspace, compounding memory, 50+ connectors, cross-silo workflows) from public premai.io product info via search. Prem Studio from premAI-io/prem-studio-tutorials repo description. Reticle from prior verified research on the attestation stack.')


# ══════════════════════════════════════════════════════════
# S4  DIGITAL FOCUS SPLIT
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '02  Digital focus  |  How attention is allocated')
txt(s, 'Fluso is the flagship, not the whole stage.', 0.55, 0.95, 12, 0.6, sz=26, bold=True, color=INK)

# Split bar
bar_y = 1.95
box(s, 0.55, bar_y, 12.23 * 0.4, 1.0, fill=INK)
box(s, 0.55 + 12.23 * 0.4, bar_y, 12.23 * 0.6, 1.0, fill=PAPER2)
txt(s, '40%  FLUSO', 0.85, bar_y + 0.28, 4.0, 0.45, sz=18, bold=True, color=WHITE)
txt(s, '60%  PREM, PREM STUDIO, RETICLE, RESEARCH', 0.55 + 12.23*0.4 + 0.3, bar_y + 0.28, 7.0, 0.45, sz=18, bold=True, color=INK)

cols = [
    ('What the 40% buys Fluso',
     'Full launch support, daily X presence, Instagram as its visual home, influencer program, signup-driving campaigns. Flagship treatment on every drop.'),
    ('What the 60% covers',
     'Prem enterprise authority on LinkedIn, Prem Studio developer content and tutorials, Reticle open-source and community work, research and industry presence.'),
    ('The discipline',
     'Measure the split monthly across published posts, paid spend and campaign hours. If Fluso drifts past half of total output, rebalance the calendar.'),
]
for i, (t, b) in enumerate(cols):
    x = 0.55 + i * 4.27
    box(s, x, 3.45, 3.95, 3.0, fill=WHITE)
    box(s, x, 3.45, 3.95, 0.06, fill=INK)
    txt(s, t, x+0.25, 3.62, 3.45, 0.65, sz=12.5, bold=True, color=INK)
    txt(s, b, x+0.25, 4.35, 3.45, 2.0, sz=10, color=GRAY)

txt(s, 'The split is about output share, not importance. Fluso gets 100% commitment inside its 40%.',
    0.55, 6.75, 12.2, 0.4, sz=10, color=GRAY)


# ══════════════════════════════════════════════════════════
# S5  POSITIONING: PREM
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '03  Positioning: Prem')
txt(s, 'Private Super Intelligence', 0.55, 1.0, 12, 0.8, sz=34, bold=True, color=INK)
txt(s, '"AI is the most powerful technology of our era. Prem makes it private, verifiable, and sovereign."  (premai.io hero)',
    0.55, 1.85, 11.8, 0.4, sz=11.5, color=GRAY)

pillars = [
    ('Private',
     'Runs on the customer\'s infrastructure. Data never leaves their VPC. No provider-side retention, no training on customer data.'),
    ('Verifiable',
     'Reticle produces a hardware-signed cryptographic attestation at the TEE layer on every inference. Checkable by any party. No competitor ships this.'),
    ('Sovereign',
     'The customer\'s jurisdiction governs. Any open model, portable across infrastructure, auditable by their team, their counsel, their regulator.'),
]
for i, (t, b) in enumerate(pillars):
    x = 0.55 + i * 4.27
    box(s, x, 2.55, 3.95, 3.6, fill=WHITE)
    box(s, x, 2.55, 3.95, 0.06, fill=INK)
    dot(s, x+0.42, 3.05, 0.09, INK)
    txt(s, t, x+0.28, 3.32, 3.4, 0.6, sz=17, bold=True, color=INK)
    txt(s, b, x+0.28, 4.0, 3.4, 2.0, sz=10.5, color=GRAY)

box(s, 0.55, 6.4, 12.23, 0.75, fill=WHITE)
txt(s, 'PROOF BANK   The premai.io homepage shows "Trusted by" logos including Microsoft, NVIDIA, AWS, Index Ventures, Innosuisse, SUPSI, Abu Dhabi Investment Office and Plug and Play. Use them as social proof exactly as the site presents them, nothing more.',
    0.85, 6.58, 11.7, 0.5, sz=9, color=INK)
notes(s, 'Hero copy and trusted-by logos read directly from the live premai.io homepage screenshot, 9 Jun 2026. Positioning language uses the company\'s own words: private, verifiable, sovereign.')


# ══════════════════════════════════════════════════════════
# S6  POSITIONING: FLUSO
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '04  Positioning: Fluso')
txt(s, '"Work deeper, not longer."', 0.55, 1.0, 12, 0.8, sz=32, bold=True, color=INK)
txt(s, 'Deep-work AI for knowledge professionals who think in 20-minute sessions, not prompt chains.',
    0.55, 1.85, 11.5, 0.4, sz=12, color=GRAY)
rule(s, 0.55, 2.5, 12.23, color=PAPER2, weight=0.02)
details = [
    ('FOR','Lawyers, analysts, researchers and engineers doing cognitive-heavy work'),
    ('WHO NEED','AI that fits into deep-work sessions without breaking concentration or creating new overhead'),
    ('UNLIKE','Chat-first tools optimised for quick answers rather than sustained reasoning sessions'),
    ('FLUSO IS','A private AI workspace with compounding memory and 50+ connectors, built for 10 to 30 minute focused workflows, with session audit trail and honest accuracy reporting'),
    ('THE PROOF','59 minutes per day lost to information search. 60% of knowledge-worker time spent on work about work. Fluso attacks both numbers.'),
]
for i, (l, b) in enumerate(details):
    y = 2.75 + i * 0.86
    txt(s, l, 0.55, y, 1.6, 0.5, sz=9, bold=True, color=GRAY)
    txt(s, b, 2.3, y, 10.4, 0.75, sz=11.5, color=INK)
dot_row(s, 12.25, 7.1)


# ══════════════════════════════════════════════════════════
# S7  THE CATEGORY
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '05  The category we are creating')
txt(s, 'Verifiable AI', 0.55, 1.0, 10, 0.85, sz=40, bold=True, color=INK)
txt(s, 'The word is already in the company\'s own hero copy. The category claim makes it ours.',
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
    box(s, x, 2.55, 3.95, 4.35, fill=WHITE if not hi else INK2)
    box(s, x, 2.55, 3.95, 0.06, fill=LGRAY if not hi else INK)
    txt(s, t, x+0.28, 2.72, 3.4, 0.65, sz=13, bold=True, color=GRAY if not hi else WHITE)
    txt(s, b, x+0.28, 3.45, 3.4, 3.1, sz=10, color=GRAY if not hi else PAPER2)
    if hi:
        txt(s, 'PREM OWNS THIS COLUMN', x+0.28, 6.5, 3.4, 0.3, sz=8, bold=True, color=LGRAY)


# ══════════════════════════════════════════════════════════
# S8  VOICE: FOUR PRINCIPLES
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '06  Voice and tone  |  Four principles, every brand surface')
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
    box(s, x, y, 0.06, 2.7, fill=INK)
    txt(s, str(i+1), x+0.3, y+0.22, 0.6, 0.5, sz=22, bold=True, color=LGRAY)
    txt(s, t, x+0.3, y+0.8, 5.2, 0.5, sz=17, bold=True, color=INK)
    txt(s, b, x+0.3, y+1.35, 5.25, 1.25, sz=10.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S9  BANNED LIST
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '06  Voice and tone  |  The banned list')
txt(s, 'Prohibited across all Prem and Fluso content. No exceptions.', 0.55, 0.95, 12, 0.35, sz=11, color=GRAY)
bcols = [
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
for i, (t, items) in enumerate(bcols):
    x = 0.55 + i * 4.27
    box(s, x, 1.4, 3.95, 5.7, fill=WHITE)
    box(s, x, 1.4, 3.95, 0.06, fill=FN_RED)
    txt(s, t, x+0.25, 1.55, 3.5, 0.35, sz=9.5, bold=True, color=FN_RED)
    for j, item in enumerate(items):
        y = 2.05 + j * 0.83
        txt(s, '×', x+0.25, y, 0.3, 0.35, sz=12, bold=True, color=FN_RED)
        txt(s, item, x+0.6, y, 3.15, 0.75, sz=9.5, color=INK)


# ══════════════════════════════════════════════════════════
# S10  VOICE IN ACTION
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '06  Voice and tone  |  In practice')
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
# S11  COLOR (monochrome)
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '07  Color  |  Monochrome, like the mark')
txt(s, 'The Prem logo is black dots on white. The brand follows the logo.',
    0.55, 0.95, 12, 0.4, sz=13, bold=True, color=INK)
txt(s, 'Verified from the live premai.io site. Photography carries the atmosphere; the identity itself stays monochrome.',
    0.55, 1.4, 12, 0.32, sz=9.5, color=GRAY)

swatches = [
    ('Ink',       '#0A0A0A', INK,   'The logo color.\nAll marks, all headlines.'),
    ('Dark',      '#17171B', INK2,  'Dark surfaces,\nquote cards, hero blocks.'),
    ('Gray',      '#5B5F63', GRAY,  'Secondary text,\nlabels, captions.'),
    ('Light gray','#B0B0AA', LGRAY, 'Rules, footnotes,\nhalftone shading.'),
    ('Paper',     '#FAFAF8', PAPER, 'Default ground.\nThe Swiss white space.'),
]
for i, (n, hx, c, u) in enumerate(swatches):
    x = 0.55 + i * 2.52
    sw = box(s, x, 1.95, 2.28, 2.3, fill=c)
    if n == 'Paper': sw.line.color.rgb = PAPER2; sw.line.width = Pt(1)
    txt(s, n, x, 4.38, 2.28, 0.35, sz=12, bold=True, color=INK)
    txt(s, hx, x, 4.74, 2.28, 0.3, sz=9.5, color=GRAY)
    txt(s, u, x, 5.08, 2.28, 0.85, sz=8.5, color=GRAY)

box(s, 0.55, 6.1, 12.23, 1.0, fill=WHITE)
box(s, 0.55, 6.1, 0.06, 1.0, fill=INK)
txt(s, 'TWO RULES', 0.85, 6.22, 2.0, 0.3, sz=8, bold=True, color=GRAY)
txt(s, '1. Color in the feed comes from photography (the Alps, real product screens), never from the identity itself.   2. Fluso accent colors, if any, are lifted from the Fluso logo file exactly as provided. Nothing is approximated, tinted, or recolored. Until that file is checked, Fluso is monochrome too.',
    0.85, 6.5, 11.7, 0.55, sz=9.5, color=INK)
notes(s, 'Current logo verified monochrome from the live premai.io screenshot (9 Jun 2026). The gradient logo.svg (#7F96FF, #F58E8E, #F2D398) found in premAI-io/static.premai.io is an older asset; do not use it unless design confirms it is still active. Neutrals: Paper #FAFAF8 and near-black from the existing PremAI deck system.')


# ══════════════════════════════════════════════════════════
# S12  TYPOGRAPHY
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '08  Typography  |  Pretendard')
txt(s, 'Pretendard', 0.55, 1.0, 8, 0.9, sz=44, bold=True, color=INK)
txt(s, 'The verified brand typeface. Prem self-hosts Pretendard Regular, SemiBold and Bold at static.premai.io/fonts, and the live site is set in a matching grotesque. One family, every brand surface.',
    0.55, 2.0, 7.0, 0.75, sz=11.5, color=GRAY)

type_rows = [
    ('Display / H1', 'Pretendard Bold, 36 to 48pt, tight leading (1.2x)'),
    ('Section head', 'Pretendard SemiBold, 20 to 28pt'),
    ('Body copy',    'Pretendard Regular, 10 to 12pt, leading 1.6x'),
    ('Labels',       'Pretendard SemiBold, 8 to 9pt, all caps, letterspaced'),
    ('Data callout', 'Pretendard Bold at display scale, ink on paper'),
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
notes(s, 'Pretendard files verified in premAI-io/static.premai.io/fonts (9 Jun 2026). If the wordmark itself is a different custom face, confirm against the master logo file; everything else runs on Pretendard.')


# ══════════════════════════════════════════════════════════
# S13  THE SWISS VISUAL WORLD
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '09  The Swiss visual world')
txt(s, 'The website already lives in the Alps. Social extends it.',
    0.55, 0.95, 12.2, 0.6, sz=22, bold=True, color=INK)

motifs = [
    ('Alpine photography',
     'The premai.io hero sets the standard: dark, monumental peaks, mist, muted light. Social photography matches that exact mood. No saturated stock, no postcard blue skies.'),
    ('The dot map',
     'Switzerland drawn in the logo\'s own black-dot language. Our most ownable graphic: the country and the mark in one image.'),
    ('Ridgelines',
     'Grayscale mountain silhouettes as footers, section breaks and Reel cover frames. Quiet geometry, never decoration.'),
    ('Swiss grid',
     'The International Typographic Style is literally Swiss: strict grid, flush-left type, white space doing the work. We inherit it by birthright.'),
]
for i, (t, b) in enumerate(motifs):
    x = 0.55 + (i % 2) * 6.4
    y = 1.8 + (i // 2) * 1.95
    box(s, x, y, 5.85, 1.75, fill=WHITE)
    box(s, x, y, 0.06, 1.75, fill=INK)
    txt(s, t, x+0.28, y+0.16, 5.3, 0.4, sz=13, bold=True, color=INK)
    txt(s, b, x+0.28, y+0.6, 5.3, 1.05, sz=9.5, color=GRAY)

box(s, 0.55, 5.8, 12.23, 1.3, fill=INK2)
txt(s, 'RULES', 0.85, 5.95, 2.0, 0.3, sz=8, bold=True, color=LGRAY)
txt(s, 'Never use the Swiss flag or cross: protected mark, and a cliche.  One motif per asset; they support the message, never carry it.  Color enters only through photography and product screens.  Every asset passes the calm test: if it shouts, it ships nowhere.',
    0.85, 6.25, 11.7, 0.75, sz=10, color=PAPER)


# ══════════════════════════════════════════════════════════
# S14  VISUAL APPLICATION
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '10  Visual application  |  Social templates')

# (a) Reel cover 9:16
rx, ry, rw, rh = 0.85, 1.3, 2.6, 4.62
box(s, rx, ry, rw, rh, fill=PAPER, line=PAPER2, lw=1.2)
ridge(s, [(rx, ry+rh-1.1),(rx+0.7, ry+rh-1.7),(rx+1.4, ry+rh-1.2),(rx+2.0, ry+rh-1.8),(rx+rw, ry+rh-1.25)], R2)
box(s, rx-0.02, ry+rh, rw+0.04, 7.5-(ry+rh), fill=PAPER)
txt(s, '23 MIN', rx+0.2, ry+0.5, 2.2, 0.6, sz=28, bold=True, color=INK)
txt(s, 'contract review,\nevery clause sourced', rx+0.2, ry+1.15, 2.2, 0.7, sz=10, color=GRAY)
dot_row(s, rx+0.32, ry+rh-0.32)
txt(s, 'REEL COVER  9:16', rx, ry+rh+0.12, rw, 0.3, sz=8, bold=True, color=GRAY)
txt(s, 'Big number, one claim,\nridge footer, dot row', rx, ry+rh+0.42, rw, 0.5, sz=8.5, color=LGRAY)

# (b) Carousel 1:1
cx, cy, cw = 4.55, 1.3, 3.6
box(s, cx, cy, cw, cw, fill=WHITE, line=PAPER2, lw=1.2)
txt(s, 'The 20-minute\nagent test', cx+0.25, cy+0.3, 3.1, 1.0, sz=18, bold=True, color=INK)
txt(s, 'One task. One timer.\nReceipts inside.', cx+0.25, cy+1.45, 3.1, 0.6, sz=10, color=GRAY)
swiss_dot_map(s, cx+1.55, cy+2.1, 1.8, spacing=0.085, r=0.022,
              color_fn=lambda t: lerp((176,176,170),(120,120,115),t), mark_lugano=False)
txt(s, 'CAROUSEL  1:1', cx, cy+cw+0.12, cw, 0.3, sz=8, bold=True, color=GRAY)
txt(s, 'Standalone claim on slide 1, dot-map watermark, summary close (never a CTA slide)',
    cx, cy+cw+0.42, cw, 0.55, sz=8.5, color=LGRAY)

# (c) Quote card 16:9
qx, qy, qw, qh = 9.0, 1.3, 3.8, 2.14
box(s, qx, qy, qw, qh, fill=INK2)
txt(s, '"Our CISO asked how.\nI sent the Reticle docs."', qx+0.25, qy+0.3, 3.3, 0.85, sz=12, bold=True, color=WHITE)
txt(s, 'Fluso session audit trail', qx+0.25, qy+1.45, 3.3, 0.3, sz=8.5, color=LGRAY)
dot_row(s, qx+0.37, qy+qh-0.28, color=WHITE)
txt(s, 'QUOTE CARD  16:9', qx, qy+qh+0.12, qw, 0.3, sz=8, bold=True, color=GRAY)
txt(s, 'Dark ground, white type, dot row as signature', qx, qy+qh+0.42, qw, 0.5, sz=8.5, color=LGRAY)

box(s, 9.0, 4.6, 3.8, 2.3, fill=WHITE)
txt(s, 'SYSTEM RULES', 9.25, 4.75, 3.3, 0.3, sz=8, bold=True, color=GRAY)
txt(s, 'Paper or dark grounds only. One motif per asset. Type does the talking. The varied-size dot row is the family signature on every template, straight from the logo.',
    9.25, 5.05, 3.35, 1.7, sz=9.5, color=INK)


# ══════════════════════════════════════════════════════════
# S15  LINKEDIN
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '11  Platform playbooks  |  LinkedIn  (Prem-led, 60% side)')

box(s, 0.55, 1.05, 4.1, 6.05, fill=WHITE)
box(s, 0.55, 1.05, 4.1, 0.06, fill=INK)
txt(s, 'STRATEGY', 0.8, 1.18, 3.6, 0.3, sz=8.5, bold=True, color=GRAY)
li_strategy = [
    ('Audience','CISOs, CTOs, compliance officers, legal ops, investors'),
    ('Goal','Become the default answer to "which enterprise AI survives a regulator\'s questions?"'),
    ('Mix','Prem authority, Prem Studio engineering content and Reticle open-source carry this channel. Fluso appears as proof, not as the program.'),
    ('Cadence','3 to 4 posts per week, maximum. Quality over volume.'),
    ('KPI','Branded search growth, qualified inbound, SQL mentions. Never followers or likes.'),
]
for i, (l, v) in enumerate(li_strategy):
    y = 1.6 + i * 1.08
    txt(s, l.upper(), 0.8, y, 3.5, 0.28, sz=7.5, bold=True, color=GRAY)
    txt(s, v, 0.8, y+0.28, 3.6, 0.78, sz=9, color=INK)

box(s, 4.85, 1.05, 7.93, 6.05, fill=WHITE)
box(s, 4.85, 1.05, 7.93, 0.06, fill=INK)
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
    txt(s, hook, 10.5, y, 2.1, 1.0, sz=8, color=MGRAY)


# ══════════════════════════════════════════════════════════
# S16  X
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '11  Platform playbooks  |  X  (Fluso-led, inside the 40%)')

box(s, 0.55, 1.05, 4.1, 6.05, fill=WHITE)
box(s, 0.55, 1.05, 4.1, 0.06, fill=INK)
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
    txt(s, l.upper(), 0.8, y, 3.5, 0.28, sz=7.5, bold=True, color=GRAY)
    txt(s, v, 0.8, y+0.28, 3.6, 0.78, sz=9, color=INK)

box(s, 4.85, 1.05, 7.93, 6.05, fill=WHITE)
box(s, 4.85, 1.05, 7.93, 0.06, fill=INK)
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
    txt(s, hook, 10.5, y, 2.1, 1.0, sz=8, color=MGRAY)


# ══════════════════════════════════════════════════════════
# S17  INSTAGRAM
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '11  Platform playbooks  |  Instagram  (Fluso-weighted)')
txt(s, 'Show the work. The Swiss visual world lives here.', 0.55, 0.95, 12, 0.5, sz=18, bold=True, color=INK)

box(s, 0.55, 1.65, 6.0, 5.45, fill=WHITE)
box(s, 0.55, 1.65, 6.0, 0.06, fill=INK)
txt(s, 'CONTENT FORMATS', 0.8, 1.78, 5.5, 0.3, sz=8.5, bold=True, color=GRAY)
ig_fmt = [
    ('Reels, 15 to 30s','Screen capture of a real task plus text overlay. No talking heads, no voiceover. Cover frame: big number, ridge footer (template, slide 10).'),
    ('Carousels','Slide 1 carries a standalone claim. Dot-map watermark bottom-right. Close on a summary slide, never a CTA slide.'),
    ('Stories','Behind-the-work only: real screenshots, drafts, whiteboards. No branded templates. Polls only when we genuinely want the answer.'),
    ('Grid rhythm','Alternate paper-ground and dark-ground cards plus alpine photography frames. Every 9th post: the dot map or a ridge, full bleed.'),
]
for i, (t, b) in enumerate(ig_fmt):
    y = 2.2 + i * 1.22
    txt(s, t, 0.8, y, 5.4, 0.3, sz=10.5, bold=True, color=INK)
    txt(s, b, 0.8, y+0.32, 5.45, 0.85, sz=8.5, color=GRAY)

box(s, 6.8, 1.65, 6.0, 5.45, fill=WHITE)
box(s, 6.8, 1.65, 6.0, 0.06, fill=INK)
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
# S18  FACEBOOK + CHANNEL MATRIX
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '11  Platform playbooks  |  Facebook + the channel matrix')

box(s, 0.55, 1.05, 12.23, 1.55, fill=WHITE)
box(s, 0.55, 1.05, 12.23, 0.06, fill=INK)
txt(s, 'FACEBOOK', 0.8, 1.18, 3.0, 0.3, sz=8.5, bold=True, color=GRAY)
txt(s, 'Role: warm remarketing and community, not organic reach. Repurpose top LinkedIn posts after 24 to 48 hours with a Facebook-specific first line. The About section is keyword-indexed for search: write it as one dense, accurate paragraph. Retarget site visitors and the email list; LinkedIn handles cold.',
    0.8, 1.5, 11.7, 1.0, sz=10, color=INK)

txt(s, 'THE CHANNEL MATRIX', 0.55, 2.9, 6, 0.3, sz=8.5, bold=True, color=GRAY)
box(s, 0.55, 3.25, 12.23, 0.45, fill=INK2)
mcols = [('CHANNEL',0.75,1.4),('LED BY',2.2,1.6),('ROLE',3.85,3.2),('CADENCE',7.15,1.95),('KPI',9.2,3.4)]
for (h, xx, ww) in mcols:
    txt(s, h, xx, 3.33, ww, 0.3, sz=8, bold=True, color=WHITE)
mrows = [
    ('LinkedIn','Prem + portfolio','Authority and inbound in regulated enterprise','3 to 4 per week','Branded search, SQL mentions'),
    ('X','Fluso','Builder credibility, product proof in public','1 to 2 per day','Qualified signups from profile'),
    ('Instagram','Fluso-weighted','Process visibility, the Swiss visual world','4 to 5 per week','Profile-to-site clicks, activations'),
    ('Facebook','Both','Warm remarketing and community','2 to 3 per week, repurposed','Re-engagement CTR'),
]
for i, row in enumerate(mrows):
    bg = WHITE if i % 2 == 0 else PAPER2
    box(s, 0.55, 3.7 + i*0.72, 12.23, 0.72, fill=bg)
    for j, ((h, xx, ww), cell) in enumerate(zip(mcols, row)):
        txt(s, cell, xx, 3.82 + i*0.72, ww, 0.55, sz=9.5, bold=(j == 0), color=INK)

txt(s, 'Across all channels combined, output share holds at 40% Fluso, 60% Prem and the rest of the portfolio.',
    0.55, 6.7, 12.2, 0.32, sz=9.5, bold=True, color=INK)
txt(s, 'Website guidance is deliberately absent: no recommendations until the live site has been reviewed page by page.',
    0.55, 7.05, 12.2, 0.3, sz=9, color=GRAY)


# ══════════════════════════════════════════════════════════
# S19  CONTENT PILLARS: PREM + PORTFOLIO
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '12  Content pillars  |  Prem and the portfolio  (the 60%)')
txt(s, 'Every piece belongs to one pillar. If it fits none, it does not ship.',
    0.55, 0.95, 12, 0.35, sz=11, color=GRAY)
p_pillars = [
    ('Verifiable AI','35%','Reticle attestation\nTEE architecture explainers\nProof walkthroughs\nOpen source vs policy promises'),
    ('Compliance architecture','25%','EU AI Act interpretation\nGDPR and HIPAA by use case\nAudit readiness frameworks\nWhat a regulator actually asks'),
    ('Engineering depth','25%','Prem Studio guides and tutorials\nOpen-model selection\nOn-premise deployment\nPortability benchmarks'),
    ('Customer evidence','15%','Named case studies only\nQuantified outcomes\nQuotes with name and title\nRegulator-facing results'),
]
for i, (t, pct, topics) in enumerate(p_pillars):
    x = 0.55 + i * 3.22
    box(s, x, 1.45, 2.95, 5.55, fill=WHITE)
    box(s, x, 1.45, 2.95, 0.06, fill=INK)
    txt(s, pct, x+0.25, 1.65, 2.4, 0.6, sz=26, bold=True, color=INK)
    txt(s, t, x+0.25, 2.35, 2.5, 0.75, sz=13, bold=True, color=INK)
    rule(s, x+0.25, 3.2, 2.45, color=PAPER2, weight=0.02)
    txt(s, topics, x+0.25, 3.35, 2.5, 3.4, sz=9.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S20  CONTENT PILLARS: FLUSO
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '12  Content pillars  |  Fluso  (the 40%)')
txt(s, 'Utility leads. 60% of Fluso output shows the product doing real work.',
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
    box(s, x, 1.45, 2.95, 0.06, fill=INK)
    txt(s, pct, x+0.25, 1.65, 2.4, 0.6, sz=26, bold=True, color=INK)
    txt(s, t, x+0.25, 2.35, 2.5, 0.75, sz=13, bold=True, color=INK)
    rule(s, x+0.25, 3.2, 2.45, color=PAPER2, weight=0.02)
    txt(s, topics, x+0.25, 3.35, 2.5, 3.4, sz=9.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S21  AUDIENCE ARCHITECTURE
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '12  Audience architecture  |  Persona to platform')
box(s, 0.55, 1.05, 12.23, 0.46, fill=INK2)
acols = [('PERSONA',0.7,2.2),('LEAD BRAND',2.95,1.2),('CHANNEL',4.2,1.7),('CONTENT',5.95,3.0),('JOB TO BE DONE',9.0,3.7)]
for (h, xx, ww) in acols:
    txt(s, h, xx, 1.12, ww, 0.32, sz=7.5, bold=True, color=WHITE)
arows = [
    ('Compliance owner (CISO)','Prem','LinkedIn, direct','EU AI Act breakdowns, attestation explainers, audit frameworks','"Help me not get fired when the regulator asks how we use AI."'),
    ('Champion (CTO, VP Eng)','Prem','LinkedIn, GitHub','Architecture posts, portability data, Reticle technical docs','"Give me technical proof I can put in front of the board."'),
    ('Validator (security eng)','Prem + Reticle','GitHub, X','TEE explainers, benchmarks, open-source contribution posts','"Show me the code, not the marketing page."'),
    ('Vertical buyer (legal ops)','Prem','LinkedIn, events','Cost comparisons vs legal AI tools, ROI evidence, case studies','"Beat the incumbents on price, match them on accuracy."'),
    ('Developer / builder','Prem Studio + Fluso','X, GitHub, community','Studio tutorials, model selection guides, workflow teardowns','"Let me try it in 10 minutes without talking to sales."'),
    ('Prosumer / power user','Fluso','Instagram, X','Deep-work content, 20-minute workflows, productivity data','"Help me do 3 hours of thinking in 45 minutes."'),
]
for i, row in enumerate(arows):
    bg = WHITE if i % 2 == 0 else PAPER2
    box(s, 0.55, 1.51 + i*0.93, 12.23, 0.93, fill=bg)
    for j, ((h, xx, ww), cell) in enumerate(zip(acols, row)):
        txt(s, cell, xx, 1.58 + i*0.93, ww, 0.82, sz=8.5, bold=(j == 0), color=INK)


# ══════════════════════════════════════════════════════════
# S22  VOCABULARY
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '13  Vocabulary  |  Words we own, words we never use')
vcols = [
    ('PREM OWNS', ['Private Super Intelligence','Verifiable','Attestation','Proof, not promise','On-premise','Sovereign','Auditable','"At the hardware layer"','Hardware-signed','"Per inference"','Reticle (always capitalised)','"Your jurisdiction governs"']),
    ('FLUSO OWNS', ['Deep work','Focused session','"20-minute session"','Compounding memory','Honest accuracy','Session audit trail','Open model','"No black box"','Cognitive load','Workflow, not chat','"Work that needs thinking"','"By Prem"']),
    ('NEITHER USES', ['Revolutionary, game-changing','Effortless, seamless','"AI-powered"','"The future of work"','Unlock, empower','"We\'re excited to announce"','"Privacy-first" as lead claim','Em dashes in copy','"Not X, but Y"','Rhetorical question hooks','Fragment staccato','Uncited figures']),
]
for i, (t, words) in enumerate(vcols):
    x = 0.55 + i * 4.27
    accent = INK if i < 2 else FN_RED
    box(s, x, 1.05, 3.95, 6.05, fill=WHITE)
    box(s, x, 1.05, 3.95, 0.06, fill=accent)
    txt(s, t, x+0.25, 1.18, 3.5, 0.32, sz=9, bold=True, color=accent)
    for j, w in enumerate(words):
        pre = '×  ' if i == 2 else ''
        txt(s, pre + w, x+0.25, 1.62 + j*0.45, 3.55, 0.4, sz=9.5,
            color=INK if i < 2 else FN_RED)


# ══════════════════════════════════════════════════════════
# S23  POSITIONING TEST
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
# S24  GOVERNANCE (roles only)
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
header(s, '13  Governance  |  Sign-off by role')
txt(s, 'Approval means sign-off before publication, not editing rights.', 0.55, 0.95, 12, 0.35, sz=11, color=GRAY)
box(s, 0.55, 1.4, 12.23, 0.44, fill=INK2)
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
    txt(s, ap, 4.2, 1.92 + i*0.64, 2.7, 0.52, sz=9, color=GRAY)
    txt(s, why, 7.0, 1.92 + i*0.64, 5.7, 0.52, sz=8.5, color=GRAY)


# ══════════════════════════════════════════════════════════
# S25  BACK COVER
# ══════════════════════════════════════════════════════════
s = ns(); add_bg(s, PAPER)
ridge(s, [(0,6.45),(2.4,5.85),(4.8,6.4),(7.2,5.7),(9.6,6.35),(12.0,5.95),(13.34,6.4)], R1)
ridge(s, [(0,6.95),(3.3,6.45),(6.6,6.9),(9.9,6.35),(13.34,6.85)], R2)
dot_row(s, 0.62, 1.5)
txt(s, 'One rule above all others:', 0.55, 1.85, 11, 0.5, sz=15, color=GRAY)
txt(s, 'Be consistently more useful\nthan you are impressive.', 0.55, 2.4, 11.5, 1.6, sz=34, bold=True, color=INK)
rule(s, 0.57, 4.35, 3.2, color=INK)
txt(s, 'The brand is a promise kept in every post, every document, every conversation. Hold the standard.',
    0.57, 4.55, 9.5, 0.6, sz=11, color=GRAY)
txt(s, 'Prem Brand Bible  v3.0   June 2026   Internal use only', 0.57, 5.25, 9, 0.35, sz=9, color=LGRAY)

prs.save('/home/user/Cadence-Architecture-Deck/PremAI-Fluso-Brand-Bible.pptx')
print(f'Saved {len(prs.slides)} slides')
