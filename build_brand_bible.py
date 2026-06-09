from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

def rgb(r, g, b):
    return RGBColor(r, g, b)

# ── Palette ────────────────────────────────────────────────
PREM_NAVY        = rgb(10,  15,  30)
PREM_NAVY2       = rgb(12,  20,  45)
PREM_NAVY3       = rgb(15,  25,  50)
PREM_BLUE        = rgb(37,  99, 235)
PREM_BLUE_LIGHT  = rgb(59, 130, 246)
PREM_SLATE       = rgb(100,116, 139)
PREM_LIGHT       = rgb(248,250, 252)

FLUSO_DARK       = rgb(26,  26,  46)
FLUSO_AMBER      = rgb(245,158,  11)
FLUSO_AMBER_L    = rgb(251,191,  36)
FLUSO_SLATE      = rgb(148,163, 184)
FLUSO_CREAM      = rgb(255,247, 237)

WHITE            = rgb(255,255,255)
OFF_WHITE        = rgb(241,245,249)
DARK_TEXT        = rgb(15,  23,  42)
MID_TEXT         = rgb(71,  85, 105)

RED              = rgb(220, 38,  38)
RED_LIGHT        = rgb(252,165, 165)
RED_BG           = rgb(30,  10,  10)
GREEN            = rgb(34, 197,  94)
INDIGO           = rgb(99, 102, 241)
EMERALD          = rgb(16, 185, 129)

# ── Helpers ────────────────────────────────────────────────
def add_bg(slide, fill):
    s = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Inches(7.5))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()

def box(slide, left, top, w, h, fill=None, line=None, lw=1.0):
    s = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(w), Inches(h))
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    return s

def txt(slide, text, left, top, w, h,
        sz=11, bold=False, color=WHITE,
        align=PP_ALIGN.LEFT, italic=False, name='Calibri'):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run()
    r.text = text; r.font.size = Pt(sz); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = color; r.font.name = name
    return tb

def mtxt(slide, lines, left, top, w, h, name='Calibri'):
    """lines = list of (text, bold, size, color, align?)"""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    for i, line in enumerate(lines):
        if len(line) == 4:
            text, bold, size, color = line
            align = PP_ALIGN.LEFT
        else:
            text, bold, size, color, align = line
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = text; r.font.size = Pt(size)
        r.font.bold = bold; r.font.color.rgb = color
        r.font.name = name
    return tb

def header_bar(slide, section_label, color=PREM_BLUE_LIGHT):
    box(slide, 0, 0, 13.33, 1.1, fill=PREM_NAVY)
    txt(slide, section_label, 0.4, 0.27, 12.5, 0.45, sz=10, bold=True, color=color)

def left_rule(slide, color=PREM_BLUE):
    box(slide, 0, 0, 0.08, 7.5, fill=color)

def section_label(slide, text, color=PREM_BLUE):
    txt(slide, text, 0.5, 0.27, 12, 0.4, sz=10, bold=True, color=color)

# ── Build ──────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

def ns():          # new slide
    return prs.slides.add_slide(blank)


# ══════════════════════════════════════════════════════════
# S1  COVER
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
# decorative bottom-right block
box(s, 9.2, 4.8, 4.1, 2.7, fill=rgb(14, 22, 48))
box(s, 12.8, 0, 0.53, 7.5, fill=FLUSO_AMBER)

txt(s, 'THE BRAND BIBLE',
    0.5, 1.8, 9, 1.1, sz=44, bold=True, color=WHITE)
txt(s, 'PremAI  +  Fluso',
    0.5, 3.0, 9, 0.75, sz=28, color=PREM_BLUE_LIGHT)
txt(s, 'Brand Architecture  |  Voice & Tone  |  Visual Identity  |  Platform Guidelines  |  Content Strategy',
    0.5, 3.85, 11, 0.45, sz=10, color=PREM_SLATE)

box(s, 0.5, 5.1, 3.5, 0.04, fill=PREM_BLUE)
txt(s, 'v1.0  |  June 2026  |  Internal Use Only',
    0.5, 5.25, 6, 0.4, sz=9, color=PREM_SLATE)
txt(s, 'CONFIDENTIAL',
    10.0, 5.25, 3.2, 0.4, sz=9, bold=True, color=PREM_BLUE,
    align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════
# S2  TABLE OF CONTENTS
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_LIGHT)
left_rule(s)
header_bar(s, 'CONTENTS')

sections = [
    ('01', 'Brand Architecture',   'Prem parent / Fluso sub-brand relationship'),
    ('02', 'Brand Positioning',    'What each brand claims and proves'),
    ('03', 'The Category',         '"Verifiable AI" — how we name what we do'),
    ('04', 'Voice & Tone',         '4 principles + the banned list + examples'),
    ('05', 'Visual Identity',      'Colors, typography, logo rules'),
    ('06', 'Platform Playbook',    'LinkedIn, X, Instagram, Facebook, website'),
    ('07', 'Content Pillars',      'The 4 topics each brand owns'),
    ('08', 'Audience Architecture','Persona-to-platform mapping'),
    ('09', 'Brand Vocabulary',     'Words we own / words we never use'),
    ('10', 'Positioning Test',     'The "Obviously Awesome" check'),
    ('11', 'Brand Governance',     'Who approves what — no exceptions'),
]

left_cols = sections[:6]
right_cols = sections[6:]

for i, (num, title, sub) in enumerate(left_cols):
    y = 1.4 + i * 0.88
    txt(s, num,   0.5,  y,       0.5,  0.35, sz=11, bold=True, color=PREM_BLUE)
    txt(s, title, 1.05, y,       5.0,  0.35, sz=11, bold=True, color=DARK_TEXT)
    txt(s, sub,   1.05, y+0.38,  5.0,  0.35, sz=8.5, color=MID_TEXT)

box(s, 6.7, 1.3, 0.02, 5.8, fill=PREM_SLATE)

for i, (num, title, sub) in enumerate(right_cols):
    y = 1.4 + i * 0.88
    txt(s, num,   6.9,  y,       0.5,  0.35, sz=11, bold=True, color=PREM_BLUE)
    txt(s, title, 7.45, y,       5.5,  0.35, sz=11, bold=True, color=DARK_TEXT)
    txt(s, sub,   7.45, y+0.38,  5.5,  0.35, sz=8.5, color=MID_TEXT)


# ══════════════════════════════════════════════════════════
# S3  BRAND ARCHITECTURE
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
section_label(s, '01 — BRAND ARCHITECTURE')

txt(s, 'One family. Two distinct identities.',
    0.5, 0.78, 11, 0.9, sz=28, bold=True, color=WHITE)

# Prem column
box(s, 0.5,  1.95, 5.7, 5.1,  fill=PREM_NAVY2)
box(s, 0.5,  1.95, 5.7, 0.07, fill=PREM_BLUE)
txt(s, 'PREM',         0.8, 2.07, 3.5, 0.55, sz=24, bold=True, color=PREM_BLUE)
txt(s, 'Masterbrand — enterprise identity',
    0.8, 2.67, 5.0, 0.35, sz=9, color=PREM_SLATE)

prem_bullets = [
    ('Audience',   'CISOs, CTOs, compliance officers, enterprise procurement'),
    ('Tone',       'Authoritative, precise, institutional — peer-to-peer CTO voice'),
    ('Channels',   'LinkedIn (primary), website, press, investor and partner docs'),
    ('Purpose',    'Build the category "Verifiable AI." Win regulated enterprise.'),
    ('Logo usage', 'Full Prem wordmark + tagline in all enterprise-facing materials'),
]
for i, (lbl, val) in enumerate(prem_bullets):
    y = 3.12 + i * 0.56
    txt(s, lbl.upper(), 0.8, y, 1.2, 0.3, sz=7.5, bold=True, color=PREM_BLUE)
    txt(s, val,         2.05, y, 3.9, 0.42, sz=9.5, color=WHITE)

# Fluso column
box(s, 6.9,  1.95, 5.9, 5.1,  fill=rgb(22, 18, 6))
box(s, 6.9,  1.95, 5.9, 0.07, fill=FLUSO_AMBER)
txt(s, 'FLUSO',        7.2, 2.07, 3.5, 0.55, sz=24, bold=True, color=FLUSO_AMBER)
txt(s, 'Product sub-brand — user-facing identity',
    7.2, 2.67, 5.4, 0.35, sz=9, color=FLUSO_SLATE)

fluso_bullets = [
    ('Audience',   'Lawyers, analysts, researchers, developers, prosumers'),
    ('Tone',       'Direct, builder-voice, grounded, honest about limitations'),
    ('Channels',   'X (primary), Instagram, product blog, Discord / community'),
    ('Purpose',    'Be the product builders recommend to each other. Earned credibility.'),
    ('Logo usage', '"Fluso by Prem" lockup on product surfaces; standalone on X/IG'),
]
for i, (lbl, val) in enumerate(fluso_bullets):
    y = 3.12 + i * 0.56
    txt(s, lbl.upper(), 7.2, y, 1.2, 0.3, sz=7.5, bold=True, color=FLUSO_AMBER)
    txt(s, val,         8.45, y, 4.1, 0.42, sz=9.5, color=WHITE)

# Middle connector
txt(s, '"Fluso by Prem"', 5.8, 3.3, 1.8, 0.42, sz=10, bold=True, color=WHITE,
    align=PP_ALIGN.CENTER)
txt(s, 'not\n"Prem\'s product"', 5.8, 3.75, 1.8, 0.65, sz=8, color=PREM_SLATE,
    align=PP_ALIGN.CENTER)
txt(s, 'Sub-brand allows\nFluso its own\npersonality while\nborrowing Prem\'s\nenterprise trust.',
    5.72, 4.5, 1.95, 2.0, sz=8, color=PREM_SLATE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════
# S4  POSITIONING — PREM
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_LIGHT)
left_rule(s)
header_bar(s, '02 — BRAND POSITIONING  |  PREM')

# Claim bar
box(s, 0.4, 1.25, 12.5, 1.45, fill=PREM_NAVY)
txt(s, '"Enterprise AI you can prove."',
    0.85, 1.38, 11.5, 0.65, sz=26, bold=True, color=WHITE)
txt(s, 'The only AI platform that ships cryptographic attestation with every inference. Not a policy — a proof.',
    0.85, 2.05, 11.5, 0.45, sz=11, color=PREM_SLATE)

# Three pillars
pillars = [
    ('NOT\na policy promise',  'A hardware-signed\nproof per inference',
     'Reticle produces a cryptographic\nattestation at the TEE layer on every\ninference. No competitor does this.\nMathematically verifiable.', PREM_BLUE),
    ('NOT\na locked platform',  'True model\nportability',
     'Run any open-source model on your\nown infrastructure. Switch models in\nunder one hour. No migration cost.\nNo vendor lock-in.', PREM_BLUE),
    ('NOT\na SaaS subscription', 'On-premise\nsovereignty',
     'Your data never leaves your VPC.\nYour jurisdiction governs — not the\nprovider\'s. Auditable by your team,\nyour legal counsel, your regulator.', PREM_BLUE),
]

for i, (not_l, is_l, detail, accent) in enumerate(pillars):
    x = 0.4 + i * 4.2
    box(s, x, 2.9, 3.9, 4.2, fill=WHITE)
    box(s, x, 2.9, 3.9, 0.06, fill=accent)
    txt(s, not_l, x+0.2, 3.02, 3.5, 0.6, sz=9, color=MID_TEXT)
    txt(s, is_l,  x+0.2, 3.68, 3.5, 0.65, sz=14, bold=True, color=DARK_TEXT)
    box(s, x+0.2, 4.38, 3.4, 0.015, fill=PREM_SLATE)
    txt(s, detail, x+0.2, 4.48, 3.5, 2.5, sz=9.5, color=MID_TEXT)


# ══════════════════════════════════════════════════════════
# S5  POSITIONING — FLUSO
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, rgb(16, 12, 4))
left_rule(s, FLUSO_AMBER)
box(s, 0, 0, 13.33, 1.1, fill=FLUSO_DARK)
txt(s, '02 — BRAND POSITIONING  |  FLUSO',
    0.5, 0.27, 12, 0.45, sz=10, bold=True, color=FLUSO_AMBER)

txt(s, '"Work deeper, not longer."',
    0.5, 1.3, 11, 0.85, sz=34, bold=True, color=WHITE)
txt(s, 'Deep-work AI for knowledge professionals who think in 20-minute sessions, not prompt chains.',
    0.5, 2.2, 11, 0.45, sz=12, color=FLUSO_SLATE)

box(s, 0.5, 2.85, 12.3, 0.015, fill=FLUSO_AMBER)

details = [
    ('FOR',       'Lawyers, analysts, researchers, engineers doing cognitive-heavy work'),
    ('WHO NEED',  'AI that integrates into deep-work sessions without breaking concentration or creating new overhead'),
    ('UNLIKE',    'ChatGPT / Copilot — chat-first tools optimised for quick answers, not sustained reasoning sessions'),
    ('FLUSO',     'Open-model AI workspace built for 10-30 min focused workflows with full session audit trail and honest accuracy reporting'),
    ('THE PROOF', '"59 min/day lost to information search. 60% of knowledge-worker time consumed by work about work." Fluso cuts both.'),
]
for i, (lbl, body) in enumerate(details):
    y = 3.05 + i * 0.8
    txt(s, lbl, 0.5,  y,    1.5, 0.55, sz=8, bold=True, color=FLUSO_AMBER)
    txt(s, body, 2.15, y,   10.6, 0.65, sz=10.5, color=WHITE)


# ══════════════════════════════════════════════════════════
# S6  THE CATEGORY
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
section_label(s, '03 — THE CATEGORY WE\'RE CREATING')

txt(s, 'Verifiable AI',
    0.5, 0.78, 10, 0.85, sz=46, bold=True, color=WHITE)
txt(s, 'Not "Private AI." Not "Secure AI." Not "Compliant AI." A new category Prem owns.',
    0.5, 1.7, 11, 0.45, sz=12, color=PREM_SLATE)

cats = [
    ('Policy-based\n"Private AI"',
     'Contractual promise that\nthe provider handles your\ndata properly.\n\nBreaks when providers\nchange Terms of Service\nor receive a subpoena.\n\nCannot be audited.\nMust be trusted.',
     PREM_SLATE, False),
    ('Perimeter-based\n"Secure AI"',
     'Firewall and access\ncontrols around the AI\nsystem\'s exterior.\n\nDoes not tell you what\nhappened inside the\nmodel during inference.\n\nProves boundary, not\nbehaviour.',
     PREM_SLATE, False),
    ('Verifiable AI',
     'Hardware-signed\ncryptographic proof\nper inference via Reticle\n(TEE layer).\n\nMathematically\nverifiable by any\nparty with access to\nthe open-source code.\n\nNo trust required.',
     PREM_BLUE, True),
]

for i, (title, body, accent, highlight) in enumerate(cats):
    x = 0.5 + i * 4.22
    bg = rgb(18, 32, 72) if highlight else rgb(14, 20, 42)
    box(s, x, 2.35, 3.9, 4.75, fill=bg)
    box(s, x, 2.35, 3.9, 0.07, fill=accent)
    txt(s, title, x+0.22, 2.48, 3.5, 0.7, sz=13, bold=True,
        color=WHITE if highlight else PREM_SLATE)
    txt(s, body,  x+0.22, 3.28, 3.5, 3.5, sz=10, color=WHITE)
    if highlight:
        box(s, x+0.22, 6.65, 3.4, 0.3, fill=rgb(20, 40, 80))
        txt(s, 'PREM\'S CATEGORY  —  OWN IT',
            x+0.22, 6.67, 3.4, 0.28, sz=7.5, bold=True, color=PREM_BLUE)


# ══════════════════════════════════════════════════════════
# S7  VOICE & TONE — 4 PRINCIPLES
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_LIGHT)
left_rule(s)
header_bar(s, '04 — VOICE & TONE  |  FOUR PRINCIPLES  (both brands)')

principles = [
    ('Precise\nover poetic',
     'Name the exact thing. Use numbers when you have them.\n"Reduces audit prep from 3 weeks to 4 days" beats "dramatically accelerates compliance."'),
    ('Useful\nover impressive',
     'Every piece of content should leave the reader able to do something they couldn\'t before.\nIf it reads like a capability tour, rewrite it.'),
    ('Specific\nover sweeping',
     'No "many companies." No "most organisations." Say who, where, how many.\nIf you can\'t source it, don\'t claim it.'),
    ('Honest\nabout limits',
     'State what the product does not do. Admit what the data does not show.\nReaders trust brands that acknowledge constraints more than brands that don\'t.'),
]

for i, (principle, explanation) in enumerate(principles):
    x = 0.4 + i * 3.22
    box(s, x, 1.3, 3.0, 5.8, fill=WHITE)
    box(s, x, 1.3, 3.0, 0.07, fill=PREM_BLUE)
    txt(s, str(i+1), x+0.22, 1.47, 0.45, 0.5, sz=24, bold=True, color=PREM_BLUE)
    txt(s, principle,  x+0.22, 2.07, 2.6,  0.72, sz=15, bold=True, color=DARK_TEXT)
    txt(s, explanation, x+0.22, 2.92, 2.6, 3.9,  sz=10, color=MID_TEXT)


# ══════════════════════════════════════════════════════════
# S8  THE BANNED LIST
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s, RED)
section_label(s, '04 — VOICE & TONE  |  THE BANNED LIST', RED_LIGHT)
txt(s, 'These patterns are prohibited across all PremAI and Fluso content. No exceptions.',
    0.5, 0.73, 12, 0.38, sz=10, color=PREM_SLATE)

banned_cols = [
    ('BANNED WORDS', [
        'Revolutionary / Game-changing / Effortless',
        'Unlock your potential / Empower your team',
        '"AI-powered" (everything is AI-powered)',
        '"The future of work"',
        '"We\'re excited to announce"',
        '"Industry-leading" (unverifiable superlative)',
    ]),
    ('BANNED STRUCTURES', [
        'Em dashes (—) anywhere in copy',
        'Rhetorical hooks: "What if your AI could..."',
        'Fragment staccato: "Faster. Smarter. Better."',
        '"Not X, but Y" construction throughout',
        'Aphoristic closers: "That\'s the future we\'re building."',
        'Ending posts with a question: "What do you think?"',
    ]),
    ('BANNED CLAIMS', [
        '"Swiss jurisdiction" when deployed in customer VPC',
        '"Privacy-first" as a primary positioning claim',
        'Any accuracy figure without a named source',
        'Implying competitors are dishonest — show data instead',
        'Any compliance cert the company does not currently hold',
        '"Pilot momentum" rhythm in product announcements',
    ]),
]

for i, (title, items) in enumerate(banned_cols):
    x = 0.4 + i * 4.28
    box(s, x, 1.28, 4.0, 5.85, fill=RED_BG)
    box(s, x, 1.28, 4.0, 0.06, fill=RED)
    txt(s, title, x+0.2, 1.4, 3.6, 0.38, sz=9.5, bold=True, color=RED_LIGHT)
    for j, item in enumerate(items):
        y = 1.9 + j * 0.82
        txt(s, 'x', x+0.22, y+0.05, 0.28, 0.4, sz=12, bold=True, color=RED)
        txt(s, item, x+0.55, y,  3.2, 0.7, sz=9.5, color=WHITE)


# ══════════════════════════════════════════════════════════
# S9  VOICE IN ACTION  — SIDE BY SIDE
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_LIGHT)
left_rule(s)
header_bar(s, '04 — VOICE & TONE  |  EXAMPLES')

examples = [
    (
        'LinkedIn product intro',
        'Introducing Prem — the revolutionary AI platform that empowers your team to unlock new levels of productivity while keeping data safe.',
        'Prem runs on your infrastructure. Every inference produces a cryptographic proof via Reticle. Your compliance team can audit what ran, when, and on which model.',
    ),
    (
        'X / Fluso post',
        'What if your AI could actually prove it\'s private? That\'s not a dream. That\'s Fluso.',
        'Fluso produced a hardware-signed audit trail for every session in our legal workflow. Our CISO asked how. I sent them the Reticle docs.',
    ),
    (
        'Partnership announcement',
        'We\'re excited to announce our partnership with [Partner]. Together we\'re shaping the future of enterprise AI.',
        '[Partner] deploying Prem across 400 compliance workflows — first regulated-industry deployment at this scale with full attestation. Implementation notes are on our blog.',
    ),
]

for i, (ctx, bad, good) in enumerate(examples):
    y = 1.35 + i * 2.0
    txt(s, ctx.upper(), 0.4, y, 12.5, 0.28, sz=8, bold=True, color=PREM_SLATE)
    box(s, 0.4,  y+0.3, 6.0, 1.5, fill=rgb(255, 241, 242))
    box(s, 0.4,  y+0.3, 6.0, 0.05, fill=RED)
    txt(s, 'DON\'T', 0.6, y+0.38, 1.0, 0.25, sz=7.5, bold=True, color=RED)
    txt(s, bad,  0.6, y+0.65, 5.6, 1.0, sz=9.5, color=DARK_TEXT)

    box(s, 7.0,  y+0.3, 6.0, 1.5, fill=rgb(240, 253, 244))
    box(s, 7.0,  y+0.3, 6.0, 0.05, fill=GREEN)
    txt(s, 'DO',  7.2, y+0.38, 1.0, 0.25, sz=7.5, bold=True, color=GREEN)
    txt(s, good, 7.2, y+0.65, 5.6, 1.0, sz=9.5, color=DARK_TEXT)


# ══════════════════════════════════════════════════════════
# S10  VISUAL IDENTITY — PREM COLORS
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
section_label(s, '05 — VISUAL IDENTITY  |  PREM COLOR SYSTEM')

prem_colors = [
    ('Deep Navy',     '#0A0F1E', PREM_NAVY,       'Primary background.\n60% of any Prem visual.'),
    ('Electric Blue', '#2563EB', PREM_BLUE,       'Primary accent.\nCTAs, highlights, links.'),
    ('Sky Blue',      '#3B82F6', PREM_BLUE_LIGHT, 'Secondary accent.\nSubtle callouts.'),
    ('Slate',         '#64748B', PREM_SLATE,      'Body text on dark.\nSecondary labels.'),
    ('Off-white',     '#F8FAFC', PREM_LIGHT,      'Light-mode backgrounds.\nLong-form content.'),
]

for i, (name, hex_v, color, usage) in enumerate(prem_colors):
    x = 0.4 + i * 2.54
    box(s, x, 1.2, 2.3, 2.9, fill=color)
    txt(s, name,  x,    4.22, 2.3, 0.4, sz=11, bold=True, color=WHITE)
    txt(s, hex_v, x,    4.65, 2.3, 0.3, sz=9, color=PREM_SLATE)
    txt(s, usage, x,    5.02, 2.3, 1.1, sz=9, color=PREM_SLATE)

box(s, 0.4, 6.2, 12.5, 1.0, fill=rgb(14, 22, 50))
txt(s, 'USAGE RULES:   60% Navy  |  25% Off-white  |  15% Blue accents  |  Never: warm tones in Prem assets  |  Never: gradients except within product UI',
    0.7, 6.38, 12.0, 0.6, sz=9, color=PREM_SLATE)


# ══════════════════════════════════════════════════════════
# S11  VISUAL IDENTITY — FLUSO COLORS
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, FLUSO_DARK)
left_rule(s, FLUSO_AMBER)
section_label(s, '05 — VISUAL IDENTITY  |  FLUSO COLOR SYSTEM', FLUSO_AMBER)

fluso_colors = [
    ('Deep Dark',     '#1A1A2E', FLUSO_DARK,   'Primary background.\n65% of Fluso visuals.'),
    ('Focused Amber', '#F59E0B', FLUSO_AMBER,  'Primary accent.\nUse with intention.'),
    ('Bright Amber',  '#FBB924', FLUSO_AMBER_L,'Hover, highlight,\ncallout states.'),
    ('Cool Slate',    '#94A3B8', FLUSO_SLATE,  'Body text labels\non dark backgrounds.'),
    ('Warm Cream',    '#FFF7ED', FLUSO_CREAM,  'Light-mode docs,\nblog backgrounds.'),
]

for i, (name, hex_v, color, usage) in enumerate(fluso_colors):
    x = 0.4 + i * 2.54
    box(s, x, 1.2, 2.3, 2.9, fill=color)
    txt(s, name,  x,    4.22, 2.3, 0.4, sz=11, bold=True, color=WHITE)
    txt(s, hex_v, x,    4.65, 2.3, 0.3, sz=9, color=FLUSO_SLATE)
    txt(s, usage, x,    5.02, 2.3, 1.1, sz=9, color=FLUSO_SLATE)

box(s, 0.4, 6.2, 12.5, 1.0, fill=rgb(30, 24, 8))
txt(s, 'USAGE RULES:   65% Deep Dark  |  20% Cream in light sections  |  15% Amber accents  |  Amber on dark = energy  |  Amber on light = caution. Test both.',
    0.7, 6.38, 12.0, 0.6, sz=9, color=FLUSO_SLATE)


# ══════════════════════════════════════════════════════════
# S12  TYPOGRAPHY
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_LIGHT)
left_rule(s)
header_bar(s, '05 — VISUAL IDENTITY  |  TYPOGRAPHY SYSTEM')

# Left: Prem
txt(s, 'PREM', 0.5, 1.3, 5.7, 0.38, sz=10, bold=True, color=PREM_BLUE)
txt(s, 'Primary: Inter (fallback: Neue Haas Grotesk / Calibri)',
    0.5, 1.72, 5.7, 0.38, sz=12, bold=True, color=DARK_TEXT)

type_rows_prem = [
    ('Display / H1',      'Inter Bold, 36-48pt, tracking -0.5%,\ntight leading (1.2x)'),
    ('H2 / Section head', 'Inter SemiBold, 20-28pt'),
    ('Body copy',         'Inter Regular, 10-12pt, leading 1.6x'),
    ('Caption / Label',   'Inter Medium, 8-9pt, ALL CAPS for labels'),
    ('Data callout',      'Inter Bold, 18-24pt, accent color'),
]
for i, (style, spec) in enumerate(type_rows_prem):
    y = 2.25 + i * 0.68
    txt(s, style, 0.5,  y,     2.0, 0.55, sz=9,   bold=True,  color=DARK_TEXT)
    txt(s, spec,  2.55, y,     3.5, 0.55, sz=8.5, bold=False, color=MID_TEXT)

box(s, 0.5, 5.7, 5.7, 1.5, fill=WHITE)
mtxt(s, [
    ('Typography Rules — Both Brands', True, 9, DARK_TEXT),
    ('Never italicise for emphasis. Bold only.', False, 8.5, MID_TEXT),
    ('Minimum 10pt digital, 8pt print.', False, 8.5, MID_TEXT),
    ('Max 70 characters per body line.', False, 8.5, MID_TEXT),
    ('Body leading minimum 1.5x.', False, 8.5, MID_TEXT),
    ('No decorative or serif fonts in UI or ads.', False, 8.5, MID_TEXT),
], 0.7, 5.78, 5.3, 1.35)

# Divider
box(s, 6.8, 1.25, 0.02, 5.95, fill=PREM_SLATE)

# Right: Fluso
txt(s, 'FLUSO', 7.05, 1.3, 5.7, 0.38, sz=10, bold=True, color=FLUSO_AMBER)
txt(s, 'Same family — Inter. Different treatment.',
    7.05, 1.72, 5.7, 0.38, sz=12, bold=True, color=DARK_TEXT)

type_rows_fluso = [
    ('Letter spacing',   'Slightly looser (+0.5%) than Prem to feel\nmore open and accessible'),
    ('Line height',      '1.7x for body — more breathing room'),
    ('Case convention',  'Sentence case preferred over Title Case in\ncaptions and CTAs'),
    ('Callouts',         'Inter Bold at display scale; amber color;  \nno framing box needed'),
    ('Code / terminal',  'Fira Code or JetBrains Mono for product\nUI and developer content'),
]
for i, (style, spec) in enumerate(type_rows_fluso):
    y = 2.25 + i * 0.68
    txt(s, style, 7.05, y,    2.2, 0.55, sz=9,   bold=True,  color=DARK_TEXT)
    txt(s, spec,  9.3,  y,    3.5, 0.55, sz=8.5, bold=False, color=MID_TEXT)

box(s, 7.05, 5.7, 5.7, 1.5, fill=WHITE)
txt(s, 'Shared typeface = brand family coherence.\nDifferent treatment = distinct product personality.',
    7.25, 5.85, 5.3, 1.1, sz=11, color=DARK_TEXT)


# ══════════════════════════════════════════════════════════
# S13  PLATFORM — LINKEDIN
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
section_label(s, '06 — PLATFORM PLAYBOOK  |  LINKEDIN  (PREM PRIMARY CHANNEL)')

box(s, 0.4, 0.82, 5.9, 6.35, fill=PREM_NAVY2)
txt(s, 'Why LinkedIn', 0.62, 0.95, 5.4, 0.38, sz=12, bold=True, color=WHITE)
li_why = [
    ('Audience',  'CISOs, CTOs, compliance officers, legal ops leads, VCs — all active and reading long-form here.'),
    ('Goal',      'Establish Prem as the default answer to "which enterprise AI can I put in front of a regulator?"'),
    ('Not for',   'Product launches, discount offers, follower counts. These are vanity metrics.'),
    ('KPI',       'Qualified inbound mentions of Prem in a sales context. Branded search volume increase.'),
]
for i, (lbl, val) in enumerate(li_why):
    y = 1.48 + i * 1.32
    txt(s, lbl.upper(), 0.62, y, 1.2, 0.28, sz=7.5, bold=True, color=PREM_BLUE)
    txt(s, val, 0.62, y+0.3, 5.5, 0.88, sz=9.5, color=WHITE)

box(s, 6.8, 0.82, 6.1, 6.35, fill=PREM_NAVY3)
txt(s, 'Execution rules', 7.02, 0.95, 5.6, 0.38, sz=12, bold=True, color=WHITE)
li_rules = [
    'Name field: "Prem AI | Enterprise AI Platform" — keyword-indexed.',
    'First 2 lines of company bio always visible pre-click. Lead with category claim.',
    'Content mix: 80% utility / insight + 20% company news.',
    'Formats: long-form posts 600-900 words > carousels > native video.',
    'Never: memes, "here\'s what I learned" openers, engagement bait.',
    'Frequency: 4x per week maximum. Quality over volume.',
    'Hashtags: 2-3 max. Industry terms only (#EnterpriseAI, #Compliance).',
    'Always cite data with source. Never "studies show."',
    'End with a statement — never with "What do you think?"',
    'Reply within 2h on day of publish. Builds algorithm equity.',
]
for i, rule in enumerate(li_rules):
    y = 1.48 + i * 0.52
    txt(s, f'{i+1}.', 7.02, y, 0.35, 0.38, sz=9, bold=True, color=PREM_BLUE)
    txt(s, rule, 7.4, y, 5.3, 0.45, sz=9, color=WHITE)


# ══════════════════════════════════════════════════════════
# S14  PLATFORM — X
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, rgb(5, 5, 14))
left_rule(s, FLUSO_AMBER)
section_label(s, '06 — PLATFORM PLAYBOOK  |  X  (FLUSO PRIMARY CHANNEL)', FLUSO_AMBER)

box(s, 0.4, 0.82, 5.9, 6.35, fill=rgb(20, 18, 8))
txt(s, 'Why X', 0.62, 0.95, 5.4, 0.38, sz=12, bold=True, color=WHITE)
x_why = [
    ('Audience',  'Builders, developers, early adopters, AI researchers, power users — already talking open-source models and privacy.'),
    ('Goal',      'Make Fluso the product builders recommend to each other. Earned credibility, not bought.'),
    ('Not for',   'Enterprise sales. Compliance deep dives. Long-form PDFs. That\'s LinkedIn.'),
    ('Voice',     'Someone who actually uses Fluso daily — not a brand account. A person.'),
]
for i, (lbl, val) in enumerate(x_why):
    y = 1.48 + i * 1.32
    txt(s, lbl.upper(), 0.62, y, 1.2, 0.28, sz=7.5, bold=True, color=FLUSO_AMBER)
    txt(s, val, 0.62, y+0.3, 5.5, 0.88, sz=9.5, color=WHITE)

box(s, 6.8, 0.82, 6.1, 6.35, fill=rgb(25, 22, 6))
txt(s, 'Execution rules', 7.02, 0.95, 5.6, 0.38, sz=12, bold=True, color=WHITE)
x_rules = [
    'Bio: 160 chars. Category claim first. Product link.',
    'Name field: "Fluso" — clean. Keywords go in bio.',
    'Post type: product observations, honest takes, workflow screenshots, real numbers.',
    'Thread format: lead tweet must stand alone without the thread.',
    'Never: "thread \U0001f9f5" openers. Never: "here\'s what I learned." Start with the insight.',
    'Frequency: 1-2 posts per day maximum. Be memorable, not prolific.',
    'Replies > posts for growth. Engage where builders already are.',
    'Quote tweets: add a take. "This is why we built Fluso" is banned.',
    'Product drops: show the UI. Show the time saved. Not "excited to share."',
    'Pinned tweet: most useful piece. Update monthly.',
]
for i, rule in enumerate(x_rules):
    y = 1.48 + i * 0.52
    txt(s, f'{i+1}.', 7.02, y, 0.35, 0.38, sz=9, bold=True, color=FLUSO_AMBER)
    txt(s, rule, 7.4, y, 5.3, 0.45, sz=9, color=WHITE)


# ══════════════════════════════════════════════════════════
# S15  PLATFORM — INSTAGRAM
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, rgb(18, 14, 24))
left_rule(s, FLUSO_AMBER)
section_label(s, '06 — PLATFORM PLAYBOOK  |  INSTAGRAM  (FLUSO)', FLUSO_AMBER)

txt(s, 'Show work. Don\'t sell product.',
    0.5, 0.78, 10, 0.65, sz=24, bold=True, color=WHITE)
txt(s, 'Instagram is a process-visibility channel. Thinking visible > capabilities claimed.',
    0.5, 1.5, 11, 0.38, sz=11, color=FLUSO_SLATE)

# Bio template
box(s, 0.4, 2.05, 5.9, 2.5, fill=rgb(28, 22, 35))
box(s, 0.4, 2.05, 5.9, 0.06, fill=FLUSO_AMBER)
txt(s, 'BIO TEMPLATE  (Fluso)', 0.62, 2.17, 5.4, 0.35, sz=9, bold=True, color=FLUSO_AMBER)
mtxt(s, [
    ('Name field: "Fluso  deep work AI" (keyword-indexed)', False, 9.5, WHITE),
    ('Line 1: What you do in 8 words or less', False, 9.5, WHITE),
    ('Line 2: Proof or differentiator', False, 9.5, WHITE),
    ('Line 3: Social proof or CTA', False, 9.5, WHITE),
    ('Link: fluso.ai or link-in-bio aggregator', False, 9.5, WHITE),
], 0.62, 2.58, 5.5, 1.75)

# Algorithm note
box(s, 0.4, 4.65, 5.9, 2.5, fill=rgb(28, 22, 35))
box(s, 0.4, 4.65, 5.9, 0.06, fill=FLUSO_AMBER)
txt(s, 'ALGORITHM NOTE  (post Dec 2024)', 0.62, 4.77, 5.4, 0.35, sz=9, bold=True, color=FLUSO_AMBER)
mtxt(s, [
    ('Hashtags deweighted. Keyword discovery now dominant.', False, 9.5, WHITE),
    ('Optimise name field + bio for search terms.', False, 9.5, WHITE),
    ('Reels get 2-3x reach of static posts.', False, 9.5, WHITE),
    ('First 3 seconds determine ~60% of completion rate.', False, 9.5, WHITE),
    ('Save hashtags to first comment, not caption.', False, 9.5, WHITE),
], 0.62, 5.18, 5.5, 1.75)

# Content types
box(s, 6.8, 2.05, 6.1, 5.1, fill=rgb(28, 22, 35))
box(s, 6.8, 2.05, 6.1, 0.06, fill=FLUSO_AMBER)
txt(s, 'CONTENT TYPES', 7.02, 2.17, 5.6, 0.35, sz=9, bold=True, color=FLUSO_AMBER)
content_types = [
    ('Reels  (8 of 15 posts)',
     '15-30s. No voiceover on AI topics — show screen, show result, show time. Text overlay only.'),
    ('Carousels  (7 of 15 posts)',
     'Swipe-to-learn. Slide 1 must work standalone. Never end on a CTA slide.'),
    ('Grid aesthetic',
     'Dark background. Amber accents. No stock photography. Real product UI only.'),
    ('Caption style',
     'First line = hook (specific, not clever). 3-5 hashtags in first comment — never in caption body.'),
    ('Stories',
     'Behind-the-work. Real screenshots. No branded templates. Polls only if genuinely curious.'),
]
for i, (ct, detail) in enumerate(content_types):
    y = 2.65 + i * 0.9
    txt(s, ct,     7.02, y,      5.6, 0.3, sz=9, bold=True, color=WHITE)
    txt(s, detail, 7.02, y+0.3,  5.6, 0.5, sz=8.5, color=FLUSO_SLATE)


# ══════════════════════════════════════════════════════════
# S16  PLATFORM — FACEBOOK + WEBSITE
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_LIGHT)
left_rule(s)
header_bar(s, '06 — PLATFORM PLAYBOOK  |  FACEBOOK + WEBSITE')

# Facebook
box(s, 0.4, 1.3, 5.9, 5.8, fill=WHITE)
box(s, 0.4, 1.3, 5.9, 0.06, fill=PREM_BLUE)
txt(s, 'FACEBOOK  (PREM + FLUSO)', 0.62, 1.42, 5.4, 0.38, sz=10, bold=True, color=PREM_BLUE)
fb = [
    ('Role',         'Community + remarketing. Organic page reach near zero. Do not optimise for organic reach.'),
    ('Page name',    '"Prem AI" — Facebook indexes About section as primary search ranking signal.'),
    ('About section','Full keyword-rich paragraph: enterprise AI, compliance, on-premise, open-source models, Fluso, workflow automation.'),
    ('Content',      'Repurpose top LinkedIn posts (24h delay). Facebook-specific CTA. Groups > page posting for reach.'),
    ('Ads',          'Retarget website visitors + email list. LinkedIn for cold; Facebook for warm re-engagement.'),
    ('Messenger',    'Enable auto-reply. Route sales/support inquiries directly to CRM.'),
]
for i, (lbl, val) in enumerate(fb):
    y = 1.95 + i * 0.84
    txt(s, lbl.upper(), 0.62, y, 1.3, 0.28, sz=7.5, bold=True, color=PREM_BLUE)
    txt(s, val, 0.62, y+0.28, 5.5, 0.48, sz=9, color=DARK_TEXT)

# Website
box(s, 6.9, 1.3, 5.9, 5.8, fill=WHITE)
box(s, 6.9, 1.3, 5.9, 0.06, fill=PREM_BLUE)
txt(s, 'WEBSITE  (premai.io + fluso.ai)', 7.12, 1.42, 5.4, 0.38, sz=10, bold=True, color=PREM_BLUE)
web = [
    ('Hero section', 'Category claim in 8 words or fewer. Proof point below. Real product screenshot — no stock photos.'),
    ('Navigation',   'Lead with use cases: "For Legal," "For Finance," "For Research." Not features, not integrations.'),
    ('Social proof', 'First 3 testimonials must be named individuals with company + title. No anonymous quotes ever.'),
    ('Blog / SEO',   'resources.premai.io on Ghost. Pillar pages: EU AI Act, enterprise AI deployment, model portability.'),
    ('Pricing page', 'Show a number or a range. "Contact for pricing" without any indication of scale loses qualified leads.'),
    ('Footer',       'Legal, Swiss HQ address, only verified certifications, GitHub link for Reticle open-source.'),
]
for i, (lbl, val) in enumerate(web):
    y = 1.95 + i * 0.84
    txt(s, lbl.upper(), 7.12, y, 1.5, 0.28, sz=7.5, bold=True, color=PREM_BLUE)
    txt(s, val, 7.12, y+0.28, 5.5, 0.48, sz=9, color=DARK_TEXT)


# ══════════════════════════════════════════════════════════
# S17  CONTENT PILLARS — PREM
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
section_label(s, '07 — CONTENT PILLARS  |  PREM')
txt(s, 'Every Prem piece belongs to one of four pillars. If it doesn\'t fit, don\'t publish it.',
    0.5, 0.73, 12, 0.38, sz=10, color=PREM_SLATE)

p_pillars = [
    ('01', 'Verifiable AI',        '40%',
     'Reticle attestation\nTEE architecture explainers\nCryptographic proof walkthroughs\nOpenness vs policy promises\nHardware-signed inference',
     PREM_BLUE),
    ('02', 'Compliance Architecture', '30%',
     'EU AI Act interpretation\nGDPR / HIPAA by use case\nAudit readiness frameworks\nRegulated industry deployments\nWhat a regulator actually asks',
     INDIGO),
    ('03', 'Engineering Depth',    '20%',
     'Open-source model selection\nOn-premise deployment guides\nModel portability benchmarks\nPerformance vs cost trade-offs\nReticle contributor updates',
     EMERALD),
    ('04', 'Customer Evidence',    '10%',
     'Named case studies only\nDeployment stories with specifics\nUser quotes: name + company + title\nQuantified time and cost outcomes\nNamed regulator-facing outcomes',
     FLUSO_AMBER),
]
for i, (num, title, pct, topics, accent) in enumerate(p_pillars):
    x = 0.4 + i * 3.22
    box(s, x, 1.28, 3.0, 5.85, fill=PREM_NAVY3)
    box(s, x, 1.28, 3.0, 0.07, fill=accent)
    txt(s, num,   x+0.2, 1.42, 0.5,  0.4,  sz=11, bold=True, color=PREM_SLATE)
    txt(s, title, x+0.2, 1.87, 2.6,  0.65, sz=14, bold=True, color=WHITE)
    txt(s, pct + ' of output', x+0.2, 2.58, 2.6, 0.35, sz=11, bold=True, color=accent)
    txt(s, topics, x+0.2, 3.05, 2.6, 3.85, sz=10, color=PREM_SLATE)


# ══════════════════════════════════════════════════════════
# S18  CONTENT PILLARS — FLUSO
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, FLUSO_DARK)
left_rule(s, FLUSO_AMBER)
section_label(s, '07 — CONTENT PILLARS  |  FLUSO', FLUSO_AMBER)
txt(s, '60% product utility. 20% deep-work philosophy. 20% builder culture. In that order.',
    0.5, 0.73, 12, 0.38, sz=10, color=FLUSO_SLATE)

f_pillars = [
    ('01', 'Deep Work',           '30%',
     'Cognitive load in AI workflows\nFocus session design\n10-30 min workflow examples\n"59 min/day lost" framing\nWork vs meta-work distinction',
     FLUSO_AMBER),
    ('02', 'Product Utility',     '35%',
     'Specific use cases with time data\nBefore / after workflow comparisons\nModel-to-use-case matching\nReal accuracy data with sources\nSession audit trail demonstrations',
     FLUSO_AMBER_L),
    ('03', 'Model Transparency',  '20%',
     'Which model, why, when\nHonest accuracy reporting\nLimitation disclosure\nOpen-source model reviews\n"No black box" content',
     EMERALD),
    ('04', 'Builder Community',   '15%',
     'Reticle open-source updates\nContributor spotlights\nDev environment walkthroughs\nHonest product retrospectives\nOpen questions to community',
     INDIGO),
]
for i, (num, title, pct, topics, accent) in enumerate(f_pillars):
    x = 0.4 + i * 3.22
    box(s, x, 1.28, 3.0, 5.85, fill=rgb(28, 24, 8))
    box(s, x, 1.28, 3.0, 0.07, fill=accent)
    txt(s, num,   x+0.2, 1.42, 0.5,  0.4,  sz=11, bold=True, color=FLUSO_SLATE)
    txt(s, title, x+0.2, 1.87, 2.6,  0.65, sz=14, bold=True, color=WHITE)
    txt(s, pct + ' of output', x+0.2, 2.58, 2.6, 0.35, sz=11, bold=True, color=accent)
    txt(s, topics, x+0.2, 3.05, 2.6, 3.85, sz=10, color=FLUSO_SLATE)


# ══════════════════════════════════════════════════════════
# S19  AUDIENCE ARCHITECTURE
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_LIGHT)
left_rule(s)
header_bar(s, '08 — AUDIENCE ARCHITECTURE  |  Persona-to-Platform Mapping')

# Table header
box(s, 0.4, 1.3, 12.5, 0.46, fill=PREM_NAVY)
cols = [
    ('PERSONA',          0.55, 2.15),
    ('BRAND',            2.75, 1.15),
    ('PRIMARY CHANNEL',  3.95, 1.85),
    ('CONTENT TYPE',     5.85, 3.05),
    ('JTBD (job to be done)', 8.95, 3.95),
]
for (h, xs, ww) in cols:
    txt(s, h, xs, 1.36, ww, 0.34, sz=7.5, bold=True, color=WHITE)

rows = [
    ('Compliance Owner (CISO)',
     'PREM',        'LinkedIn + Direct',
     'Long-form EU AI Act breakdowns, attestation explainers, audit frameworks',
     '"Help me not get fired when the regulator asks how we use AI."'),
    ('Champion (CTO / VP Eng)',
     'PREM',        'LinkedIn + GitHub',
     'Architecture posts, model portability, Reticle technical docs',
     '"Give me technical proof I can put in front of the board."'),
    ('Validator (Security Eng)',
     'PREM + Fluso', 'GitHub + X',
     'TEE explainers, benchmark data, open-source contribution posts',
     '"Show me the code. Not the marketing page."'),
    ('Vertical Buyer (Legal Ops)',
     'PREM',        'LinkedIn + Events',
     'Harvey / Legora cost comparisons, ROI calculator, use-case studies',
     '"Beat Harvey on price, match it on accuracy for litigation."'),
    ('Developer / Builder',
     'FLUSO',       'X + GitHub + Discord',
     'API docs, model selection guides, workflow tool teardowns',
     '"Let me try it in 10 minutes without speaking to sales."'),
    ('Prosumer / Power User',
     'FLUSO',       'Instagram + X',
     'Deep-work content, 20-min workflow posts, productivity data',
     '"Help me do 3 hours of thinking in 45 minutes."'),
]

for i, row in enumerate(rows):
    bg = WHITE if i % 2 == 0 else OFF_WHITE
    box(s, 0.4, 1.76 + i * 0.82, 12.5, 0.8, fill=bg)
    data = [row[0], row[1], row[2], row[3], row[4]]
    for j, ((h, xs, ww), cell) in enumerate(zip(cols, data)):
        c = PREM_BLUE if j == 1 else DARK_TEXT
        b = j == 1
        txt(s, cell, xs, 1.8 + i * 0.82, ww, 0.72, sz=8.5, bold=b, color=c)


# ══════════════════════════════════════════════════════════
# S20  BRAND VOCABULARY
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
section_label(s, '09 — BRAND VOCABULARY  |  Own. Avoid. Never.')
txt(s, 'The vocabulary is the brand. Consistency in language builds category ownership.',
    0.5, 0.72, 12, 0.38, sz=10, color=PREM_SLATE)

# Column 1: Prem owns
box(s, 0.4,  1.28, 3.78, 5.9, fill=PREM_NAVY2)
box(s, 0.4,  1.28, 3.78, 0.06, fill=PREM_BLUE)
txt(s, 'PREM OWNS', 0.62, 1.4, 3.4, 0.35, sz=9, bold=True, color=PREM_BLUE)
prem_owns = [
    'Verifiable', 'Attestation', 'On-premise',
    'Sovereign', 'Proof (not promise)', 'Auditable',
    '"At the hardware layer"', 'Cryptographic proof',
    'Reticle (always capitalise)', 'Hardware-signed',
    '"Per inference"', 'TEE (with plain-English gloss)',
    'Model portability', '"Your jurisdiction governs"',
]
for i, w in enumerate(prem_owns):
    txt(s, w, 0.62, 1.92 + i*0.30, 3.4, 0.27, sz=9.5, color=WHITE)

# Column 2: Fluso owns
box(s, 4.56, 1.28, 3.78, 5.9, fill=rgb(25, 20, 6))
box(s, 4.56, 1.28, 3.78, 0.06, fill=FLUSO_AMBER)
txt(s, 'FLUSO OWNS', 4.78, 1.4, 3.4, 0.35, sz=9, bold=True, color=FLUSO_AMBER)
fluso_owns = [
    'Deep work', 'Flow state', 'Knowledge work',
    '"20-minute session"', 'Honest accuracy',
    '"Actually did X in N minutes"',
    'Session audit trail', 'Open model',
    '"No black box"', 'Cognitive load',
    '"Work that needs thinking"',
    'Workflow (not chat)', '"By Prem"',
    '"Focused session"',
]
for i, w in enumerate(fluso_owns):
    txt(s, w, 4.78, 1.92 + i*0.30, 3.4, 0.27, sz=9.5, color=WHITE)

# Column 3: Neither uses
box(s, 8.72, 1.28, 4.2, 5.9, fill=RED_BG)
box(s, 8.72, 1.28, 4.2, 0.06, fill=RED)
txt(s, 'NEITHER BRAND USES', 8.94, 1.4, 3.8, 0.35, sz=9, bold=True, color=RED_LIGHT)
neither = [
    'Revolutionary / Game-changing',
    'Effortless / Seamless',
    '"AI-powered" (everything is)',
    '"The future of work"',
    'Unlock / Empower',
    '"We\'re excited to announce"',
    '"Privacy-first" as primary claim',
    'Em dashes (—) in any copy',
    '"Not X, but Y" construction',
    'Rhetorical question hooks',
    'Fragment staccato ("Faster. Better.")',
    '"That\'s the future we\'re building."',
    'Uncited accuracy figures',
    '"Swiss jurisdiction" in customer VPC',
]
for i, w in enumerate(neither):
    txt(s, 'x  ' + w, 8.94, 1.92 + i*0.295, 3.8, 0.27, sz=8.5, color=RED_LIGHT)


# ══════════════════════════════════════════════════════════
# S21  POSITIONING TEST
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_LIGHT)
left_rule(s)
header_bar(s, '10 — THE "OBVIOUSLY AWESOME" POSITIONING TEST  (April Dunford framework)')

tests = [
    ('1. Target customer',
     'Who, exactly? Not "enterprises." CISOs at AmLaw 200 law firms managing EU AI Act deployment before the August 2025 grace-period deadline.'),
    ('2. Market category',
     'What game are we playing? "Verifiable AI" — not "enterprise AI assistant," not "private LLM," not "secure AI platform."'),
    ('3. Unique attributes',
     'What do we have that others don\'t? Hardware-signed proof per inference via Reticle (open-source TEE attestation stack). No other vendor ships this.'),
    ('4. Value to the buyer',
     'Why does that matter? Cryptographic audit trail = can answer the regulator directly. A policy promise cannot.'),
    ('5. Competitive alternatives',
     'What would they use instead? Policy-based tools (Anthropic Enterprise), perimeter security (Langchain / Dust), Harvey / Legora for legal vertical.'),
    ('6. Proof of the claim',
     'What verifies it? Reticle open-source repository. TEE architecture documentation. Named customer deployments with attestation on record.'),
]

for i, (label, body) in enumerate(tests):
    y = 1.35 + i * 0.99
    bg = WHITE if i % 2 == 0 else OFF_WHITE
    box(s, 0.4, y, 12.5, 0.92, fill=bg)
    txt(s, label, 0.6,  y+0.1, 2.8, 0.72, sz=10, bold=True, color=DARK_TEXT)
    txt(s, body,  3.5,  y+0.1, 9.0, 0.78, sz=9.5, color=DARK_TEXT)
    # Pass indicator
    box(s, 12.18, y+0.22, 0.54, 0.47, fill=GREEN)
    txt(s, 'Y', 12.22, y+0.27, 0.45, 0.37, sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

txt(s, 'Run this test before every campaign brief, product page draft, and major social push.',
    0.4, 7.15, 12.5, 0.3, sz=8.5, color=PREM_SLATE)


# ══════════════════════════════════════════════════════════
# S22  BRAND GOVERNANCE
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
section_label(s, '11 — BRAND GOVERNANCE  |  Who approves what. No exceptions.')
txt(s, 'Approval does not mean editing. It means sign-off before the content goes live.',
    0.5, 0.72, 12, 0.38, sz=10, color=PREM_SLATE)

# Header row
box(s, 0.4, 1.28, 12.5, 0.44, fill=PREM_NAVY3)
gov_cols = [
    ('CONTENT TYPE',   0.55, 3.1),
    ('APPROVER',       3.72, 2.6),
    ('REASON',         6.38, 6.45),
]
for (h, xs, ww) in gov_cols:
    txt(s, h, xs, 1.34, ww, 0.34, sz=8.5, bold=True, color=WHITE)

gov_rows = [
    ('External press / media statements',
     'CEO (Simone Giacomelli)',
     'Any statement attributed to the company or leadership. No exceptions.'),
    ('Technical claims and benchmarks',
     'CTO (Jaipal Singh)',
     'Accuracy figures, architecture claims, certifications. Fact-check with CTO before publishing.'),
    ('Pricing mentions (public)',
     'Sales (Marco) + CEO',
     'Public pricing errors are costly and hard to retract. Both sign off.'),
    ('Campaign strategy + content calendars',
     'Marketing lead (Yash Ranka)',
     'Campaign briefs, channel plans, content calendars, and influencer briefs.'),
    ('Platform bios and handles',
     'Marketing lead + CEO',
     'Bios change how we are discovered in search. Both approve.'),
    ('Brand visual changes',
     'CEO + external design review',
     'Color, logo, or typography changes require a visual consistency audit before launch.'),
    ('Influencer content',
     'Marketing lead. CEO for reach >50K',
     'Approve key messages and data claims. Influencers own their voice; we approve the facts.'),
    ('Crisis or regulatory response',
     'CEO + legal (if available)',
     'Any response to a regulatory inquiry, data incident, or significant public criticism.'),
]

for i, (ct, approver, reason) in enumerate(gov_rows):
    bg = PREM_NAVY2 if i % 2 == 0 else PREM_NAVY3
    box(s, 0.4, 1.72 + i * 0.68, 12.5, 0.67, fill=bg)
    txt(s, ct,       0.55, 1.77 + i*0.68, 3.1, 0.57, sz=9,   bold=True,  color=WHITE)
    txt(s, approver, 3.72, 1.77 + i*0.68, 2.6, 0.57, sz=9,   bold=False, color=PREM_BLUE_LIGHT)
    txt(s, reason,   6.38, 1.77 + i*0.68, 6.4, 0.57, sz=8.5, bold=False, color=PREM_SLATE)


# ══════════════════════════════════════════════════════════
# S23  BACK COVER
# ══════════════════════════════════════════════════════════
s = ns()
add_bg(s, PREM_NAVY)
left_rule(s)
box(s, 12.8, 0, 0.53, 7.5, fill=FLUSO_AMBER)

txt(s, 'One rule above all others:',
    0.5, 1.85, 11.5, 0.65, sz=17, color=PREM_SLATE)
txt(s, 'Build trust by being consistently\nmore useful than you are impressive.',
    0.5, 2.6, 11.5, 1.65, sz=32, bold=True, color=WHITE)

box(s, 0.5, 4.55, 5.5, 0.04, fill=PREM_BLUE)

txt(s, 'The brand is a promise kept in every post, every document, every conversation.\nHold the standard. The market rewards consistency.',
    0.5, 4.75, 10.5, 0.8, sz=11, color=PREM_SLATE)

txt(s, 'PremAI + Fluso Brand Bible  v1.0  |  June 2026  |  Internal Use Only',
    0.5, 6.82, 9, 0.4, sz=9, color=PREM_SLATE)
txt(s, 'CONFIDENTIAL',
    10.2, 6.82, 3, 0.4, sz=9, bold=True, color=PREM_BLUE, align=PP_ALIGN.RIGHT)


# ── Save ────────────────────────────────────────────────────
out_path = '/home/user/Cadence-Architecture-Deck/PremAI-Fluso-Brand-Bible.pptx'
prs.save(out_path)
print(f"Saved: {out_path}  ({len(prs.slides)} slides)")
