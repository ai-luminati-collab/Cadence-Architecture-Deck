"""Builds the LinkedIn creative-references doc (HTML -> Google Doc import)."""
import html, json, pathlib

CDN = "https://d8j0ntlcm91z4.cloudfront.net/user_3GEQSkDBYPDfhyXYztyME2lTyub/"
MOCK = {
    1: "hf_20261002_081506_f9b36129-7df2-484f-ba9a-3dff617f0429.png",
    2: "hf_20261002_081536_7fcce898-52dc-431e-b79d-1fedaa80d828.png",
    3: "hf_20261002_081506_6929cafd-3e75-493e-8b4e-b858b8a5e8b5.png",
    4: "hf_20261002_081506_d0aa29e7-f416-4fc2-ac6c-7b5406618661.png",
    5: "hf_20261002_081506_9a9bdb0b-9671-4491-956f-2370be6fc6b4.png",
    6: "hf_20261002_081506_34d56740-29df-4747-9f81-a5b683e5f2c5.png",
    7: "hf_20261002_081506_d3e8c5c4-2ebd-426a-8fdd-bef474df5b49.png",
    8: "hf_20261002_081506_10ada9fd-1d05-40df-a167-acd22047da7f.png",
    9: "hf_20261002_081536_5aacdc53-a4ed-4b3f-80f4-f630f5b7eb3a.png",
    10: "hf_20261002_081506_db40ae3d-2cd7-47d3-897c-743d11e617ec.png",
    11: "hf_20261002_081506_62cf34db-47f7-461c-acad-654991e9d3e4.png",
    12: "hf_20261002_081506_599da261-db6b-47a6-916a-0fdc95b51fc9.png",
    13: "hf_20261002_081536_420223b7-862f-4405-9655-fb663b49be8f.png",
    14: "hf_20261002_081536_1fb73282-5c57-460a-9c25-69a86a25efc7.png",
    15: "hf_20261002_081613_4fdf31c0-72c2-40a9-9904-f4438d8aa1d1.png",
    16: "hf_20261002_081613_bcd4a888-3065-4e9e-a342-efa8aa70a664.png",
    17: "hf_20261002_081613_aaed56aa-b08f-4dbe-a99b-a51b5ba651f9.png",
    18: "hf_20261002_081547_b997780a-7c1c-49ca-9fe2-df69e265f026.png",
}

def li_lib(q): return f"https://www.linkedin.com/ad-library/search?accountOwner={q.split('%20')[0]}"
def meta_lib(q): return ("https://www.facebook.com/ads/library/?active_status=all&ad_type=all"
                         f"&country=ALL&q={q}&search_type=keyword_unordered&media_type=all")

CONCEPTS = json.loads(pathlib.Path(__file__).with_name("concepts.json").read_text())

PRODUCTS = [
    ("FLUSO", "Fluso", "The proactive AI workspace",
     "Goal: signups and desktop downloads (fluso.ai). Buyer: lawyers, fund analysts, clinicians and ops leads drowning in Gmail, Slack and Outlook.",
     "Warm, human, editorial. Cream paper, ink black, one burnt-orange accent. Real people at real desks. The product shows up as a calm card, never a robot.",
     "#C2410C"),
    ("ENCLAVE", "Prem Enclave", "Encrypted compute for sovereign AI",
     "Goal: demo requests. Buyer: CISOs, CTOs and infra leads at banks, insurers, government, defense, sovereign clouds and GPU operators.",
     "Monolithic and Swiss-precise. Matte black, steel, ice blue, oversized type. Privacy shown as theatre: auctions, sealed boxes, failed hackers.",
     "#0369A1"),
    ("API", "Prem Enclave API", "Encrypted access to frontier open-weight models",
     "Goal: API signups and keys issued. Buyer: AI engineers, platform teams and CTOs building on LLMs in regulated companies.",
     "Developer-native. Terminal dark, mono type, lime-green accent. Code, numbers and receipts. No stock photos.",
     "#4D7C0F"),
]

e = html.escape

def yt(vid): return f"https://www.youtube.com/watch?v={vid}"
def thumb(vid): return f"https://img.youtube.com/vi/{vid}/hqdefault.jpg"

def card(c, accent):
    is_video = c["format"] == "video"
    mock_h = 350 if c.get("ratio") == "4:5" else 280
    ref = c["ref"]
    rows = []
    rows.append(
        f'<tr><td colspan="2" class="ch">'
        f'<span class="cht">{e(c["code"])}  ·  {e(c["name"])}</span><br>'
        f'<span class="chs">{"OUR FORMAT: VIDEO · " + e(c["length"]) if is_video else "OUR FORMAT: STATIC · single image"} · {e(c["ratio"])} · Mechanic: {e(c["mechanic"])}</span></td></tr>')
    rows.append(
        '<tr>'
        f'<td class="ref">'
        f'<p class="lab">THE REFERENCE · {e(ref["kind"])}</p>'
        f'<p><img src="{thumb(ref["yt"])}" width="280" height="210"></p>'
        f'<p class="b10"><a href="{yt(ref["yt"])}" style="color:{accent}">▶ WATCH THE SOURCE VIDEO ↗</a></p>'
        f'<p class="b10">{e(ref["brand"])}: {e(ref["title"])}</p>'
        + ('' if is_video else '<p class="s9" style="color:#9a3412"><b>Format note:</b> this reference is a video. We borrow its idea and look for a static ad. To see this brand\'s live static ads, use the ad library links below.</p>')
        + f'<p class="s9">{e(ref["why"])}</p>'
        f'<p class="g8">Their live ads: <a href="{li_lib(ref["lib"])}">LinkedIn Ad Library</a> · <a href="{meta_lib(ref["lib"])}">Meta Ad Library</a>'
        + (f' · <a href="{ref["extra"]}">Campaign write-up</a>' if ref.get("extra") else "") + '</p></td>'
        f'<td class="ours">'
        f'<p style="font-size:8pt;color:{accent};font-weight:bold">OUR CONCEPT · {"VIDEO (STILL FROM THE STORYBOARD)" if is_video else "STATIC AD MOCKUP"}</p>'
        f'<p><img src="{CDN + MOCK[c["mock"]]}" width="280" height="{mock_h}"></p>'
        f'<p class="b10">What we steal</p><p class="s9">{e(c["steal"])}</p>'
        '</td></tr>')
    if is_video:
        beats = "".join(f'<li><b>{e(t)}</b>  {e(d)}</li>' for t, d in c["beats"])
        body = (f'<p class="b10">Storyboard ({e(c["length"])}, sound-off, burned-in captions)</p>'
                f'<ul class="s9">{beats}</ul>')
    else:
        body = (f'<p class="s9"><b>Intro text:</b> {e(c["intro"])}</p>'
                f'<p class="s9"><b>On-image headline:</b> {e(c["headline"])}</p>'
                f'<p class="s9"><b>Visual:</b> {e(c["visual"])}</p>')
    body += (f'<p class="s9"><b>Ad headline (below {"video" if is_video else "image"}):</b> {e(c["li_headline"])}   <b>CTA button:</b> {e(c["cta"])}</p>'
             f'<p class="s9"><b>Aim it at:</b> {e(c["audience"])}</p>')
    rows.append(f'<tr><td colspan="2">{body}</td></tr>')
    return f'<table class="t" border="1" cellpadding="6">{"".join(rows)}</table><p></p>'

out = []
out.append('<html><head><meta charset="utf-8"><style>body{font-family:Arial}.lab{font-size:8pt;color:#666666;font-weight:bold}.b10{font-size:10pt;font-weight:bold}.s9{font-size:9pt}.s10{font-size:10pt}.g8{font-size:8pt;color:#666666}.ref{width:50%;vertical-align:top;background:#f4f4f4}.ours{width:50%;vertical-align:top;background:#ffffff}.t{border-collapse:collapse;width:100%;border-color:#dddddd}.hd{color:#ffffff;font-size:9pt;font-weight:bold;background:#111111}.dk{background:#111111}.ch{background:#111111}.cht{color:#ffffff;font-size:15pt;font-weight:bold}.chs{color:#bbbbbb;font-size:9pt}</style></head><body>')
out.append('<p style="font-size:9pt;color:#888888">PREM AI · LINKEDIN CREATIVE · OCTOBER 2026</p>')
out.append('<h1 style="font-size:30pt">LinkedIn Ad References: Fluso, Prem Enclave, Enclave API</h1>')
out.append('<p style="font-size:12pt">18 unique ad concepts (3 statics + 3 short videos per product), each paired with a real reference video from an AI or privacy brand. No two concepts use the same creative mechanic.</p>')
out.append('<p class="s10"><b>How to use this doc.</b> Each card has two halves. <b>Left: the reference.</b> Every reference is a <b>video</b> (a frame from it is shown), and the label above the frame says exactly what it is: a paid ad, a launch video, a keynote, an explainer, or a third-party upload or coverage. The link <b>▶ WATCH THE SOURCE VIDEO ↗</b> sits directly under each frame, because Google Docs import drops links on images. On static concepts, the reference is a video we borrow the idea from; each card links to that brand\'s live ads in the LinkedIn and Meta Ad Libraries for its real static ads. <b>Right: our concept.</b> For a static it\'s a mockup of the ad; for a video it\'s a still from the storyboard. The copy, storyboard and targeting sit underneath.</p>')
out.append('<p style="font-size:9pt;color:#9a3412"><b>Before anything runs:</b> the mockups are directional comps made with an image model. Proofread the text inside every image (models misspell), swap in the real logo, fonts and palette, and get product and legal to sign off on each claim (&lt;40ms, post-quantum, "not even we can see inside", zero data retention). These claims come from Prem\'s own launch material.</p>')

# Patterns section
out.append('<h2>What the best AI brands are doing right now</h2>')
pats = [
    ("Pick a fight with the category", "Anthropic's Super Bowl line \"Ads are coming to AI. But not to Claude.\" boosted Claude daily users by about 11%. Name what everyone else does, then refuse to do it.", "Enclave"),
    ("Warm, human, anti-robot", "Claude \"Keep thinking\" (Mother) and ChatGPT's first brand films use film grain, real people and no glowing brains. The product barely appears.", "Fluso"),
    ("One idea, one image, oversized type", "Apple's \"Privacy. That's iPhone.\" and OpenAI's pointillism Super Bowl spot each say one thing and leave a lot of empty space.", "Enclave and API"),
    ("Privacy as theatre", "Apple's \"Data Auction\" turns an abstract risk into a scene you can't forget. Abstract security beats become visible.", "Enclave"),
    ("The product UI is the hero", "Notion Agents, Superhuman and Cohere North sell the outcome with a calm screen, not a feature list.", "Fluso"),
    ("Speak developer", "Cursor, Vercel and Cerebras use code, terminals and one big number. Credibility comes from specifics.", "API"),
]
out.append('<table class="t" border="1" cellpadding="6"><tr class="dk"><td class="hd"><b>Pattern</b></td><td class="hd"><b>Evidence</b></td><td class="hd"><b>Where we use it</b></td></tr>'
           + "".join(f'<tr><td class="s9"><b>{e(a)}</b></td><td class="s9">{e(b)}</td><td class="s9">{e(c)}</td></tr>' for a, b, c in pats) + '</table>')

# At a glance
out.append('<h2>All 18 at a glance</h2>')
rows = ""
for key, pname, *_ in PRODUCTS:
    for c in CONCEPTS[key]:
        rows += (f'<tr><td class="s9">{e(c["code"])}</td><td class="s9">{e(pname)}</td>'
                 f'<td class="s9">{"Video" if c["format"]=="video" else "Static"}</td>'
                 f'<td class="s9"><b>{e(c["name"])}</b></td><td class="s9">{e(c["mechanic"])}</td>'
                 f'<td class="s9"><a href="{yt(c["ref"]["yt"])}">{e(c["ref"]["brand"])}: {e(c["ref"]["title"])}</a></td>'
                 f'<td class="s9">{e(c["ref"]["kind"])}</td></tr>')
out.append('<table class="t" border="1" cellpadding="5"><tr class="dk">'
           + "".join(f'<td class="hd"><b>{h}</b></td>' for h in ["#", "Product", "Our format", "Concept", "Mechanic", "Reference", "Reference type"]) + f'</tr>{rows}</table>')

for key, pname, tagline, goal, vibe, accent in PRODUCTS:
    out.append('<br style="page-break-before:always">')
    out.append(f'<h1 style="color:{accent}">{e(pname)}</h1>')
    out.append(f'<p style="font-size:13pt"><b>{e(tagline)}</b></p>')
    out.append(f'<p class="s10">{e(goal)}</p>')
    out.append(f'<p class="s10"><b>Look and feel:</b> {e(vibe)}</p>')
    out.append(f'<h2>{e(pname)}: static ads</h2>')
    for c in CONCEPTS[key]:
        if c["format"] == "static": out.append(card(c, accent))
    out.append(f'<h2>{e(pname)}: video ads (5–10 seconds)</h2>')
    for c in CONCEPTS[key]:
        if c["format"] == "video": out.append(card(c, accent))

out.append('<h2>LinkedIn production specs</h2><ul style="font-size:10pt">'
           '<li><b>Single image:</b> 1080×1080 (1:1) or 1080×1350 (4:5, best on mobile). Keep text under about 20% of the frame. Logo small, bottom corner.</li>'
           '<li><b>Video:</b> 1:1 or 4:5, 5–10s, MP4. Show the brand within 2s. Burn in captions because most LinkedIn video plays muted. The end card holds 1.5–2s with the CTA.</li>'
           '<li><b>Copy:</b> intro text under 150 characters so it isn\'t cut off; headline under 70 characters.</li>'
           '<li><b>Testing:</b> run the 3 statics per product as one A/B set first, since they\'re cheaper. Cut the video budget toward whichever mechanic wins (provocation, product UI or number).</li></ul>')
out.append('<p style="font-size:8pt;color:#888888">Mockups generated on Higgsfield (gpt_image_2_5). Reference frames are YouTube thumbnails of the source ads; all rights belong to their owners. They are for internal creative reference only.</p>')
out.append('</body></html>')
doc = "\n".join(out)
for t in ("</tr>", "</p>", "</li>", "<tr>"):
    doc = doc.replace(t, t + "\n") if t != "<tr>" else doc.replace(t, "\n" + t)
pathlib.Path(__file__).with_name("linkedin-creative-references.html").write_text(doc)
print("ok", len("\n".join(out)))
