from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu
from PIL import Image, ImageDraw
import math, os, io

CHARCOAL  = RGBColor(0x11, 0x13, 0x18)
CHARCOAL2 = RGBColor(0x18, 0x1c, 0x24)
CHARCOAL3 = RGBColor(0x1e, 0x23, 0x30)
TEAL      = RGBColor(0x00, 0xc9, 0xb1)
WHITE     = RGBColor(0xf4, 0xf4, 0xf2)
DIM       = RGBColor(0x8b, 0x92, 0xa5)
DIMMER    = RGBColor(0x4a, 0x51, 0x68)

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]

def slide():
    return prs.slides.add_slide(BLANK)

def rect(s, l, t, w, h, fill=None):
    sh = s.shapes.add_shape(1, l, t, w, h)
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    else:
        sh.fill.background()
    sh.line.fill.background()
    return sh

def txb(s, text, l, t, w, h, size, bold=False, color=WHITE,
        align=PP_ALIGN.LEFT, wrap=True, italic=False):
    box = s.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame; tf.word_wrap = wrap
    p   = tf.paragraphs[0]; p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size   = Pt(size)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return box

def hline(s, l, t, w, color=TEAL, thick=2):
    bar = s.shapes.add_shape(1, l, t, w, Pt(thick))
    bar.fill.solid(); bar.fill.fore_color.rgb = color
    bar.line.fill.background()

def eyebrow(s, text, l, t, w):
    txb(s, text.upper(), l, t, w, Inches(0.28), size=9, bold=True, color=TEAL)

def section_header(s, eye, title, sub):
    rect(s, 0, 0, W, H, fill=CHARCOAL)
    rect(s, 0, 0, Inches(0.06), H, fill=TEAL)
    rect(s, Inches(0.06), 0, W - Inches(0.06), Inches(1.15), fill=CHARCOAL2)
    eyebrow(s, eye, Inches(0.3), Inches(0.1), Inches(10))
    txb(s, title, Inches(0.3), Inches(0.32), Inches(12.5), Inches(0.52),
        size=22, bold=True, color=WHITE)
    txb(s, sub, Inches(0.3), Inches(0.84), Inches(12.5), Inches(0.26),
        size=10, color=DIM)
    hline(s, Inches(0.3), Inches(1.18), Inches(1.4))

# ── SLIDE 1: COVER ──────────────────────────────────────────────
s1 = slide()
rect(s1, 0, 0, W, H, fill=CHARCOAL)
rect(s1, 0, 0, Inches(0.06), H, fill=TEAL)
rect(s1, Inches(9.5), 0, W - Inches(9.5), H, fill=CHARCOAL2)

eyebrow(s1, "Account Manager", Inches(0.3), Inches(0.55), Inches(9))

# Photo — crop to square from top, save to buffer at full res
photo_path = "/Users/i763300/Desktop/ClaudeProjects/account-mgr-presentation/myrna.png"
if os.path.exists(photo_path):
    img = Image.open(photo_path).convert("RGB")
    w, h = img.size
    side = min(w, h)
    left = (w - side) // 2
    top = 0  # crop from top to keep face
    img_cropped = img.crop((left, top, left + side, top + side))
    img_cropped = img_cropped.resize((600, 600), Image.LANCZOS)
    buf = io.BytesIO()
    img_cropped.save(buf, format="PNG", dpi=(300, 300))
    buf.seek(0)
    photo_size = Inches(2.6)
    photo_left = Inches(10.2)
    photo_top  = Inches(0.45)
    s1.shapes.add_picture(buf, photo_left, photo_top,
                          width=photo_size, height=photo_size)

txb(s1, "Myrna", Inches(0.3), Inches(1.0), Inches(9), Inches(0.75),
    size=60, bold=True, color=WHITE)
txb(s1, "Gamal Fahim", Inches(0.3), Inches(1.72), Inches(9), Inches(0.75),
    size=60, bold=True, color=TEAL)

txb(s1, "Senior Sales Development Executive  ·  SAP",
    Inches(0.3), Inches(2.6), Inches(9), Inches(0.32), size=13, color=DIM)
txb(s1, "MEA North  ·  Enterprise Clients",
    Inches(0.3), Inches(2.92), Inches(9), Inches(0.28), size=11, color=DIM)

hline(s1, Inches(0.3), Inches(3.32), Inches(3.0))

txb(s1, "I want to own a client relationship completely.\nTo be the person they call, the person they trust,\nand the person who makes them successful.",
    Inches(0.3), Inches(3.45), Inches(8.8), Inches(0.95),
    size=13, color=WHITE, wrap=True)

# Stats
stats = [
    ("10+",  "Years enterprise client experience"),
    ("3",    "Markets: Egypt, South Africa, Morocco"),
    ("200+", "Partners built from ground up at IBM"),
]
sy = Inches(3.5)
for num, label in stats:
    rect(s1, Inches(9.65), sy, Inches(3.45), Inches(1.1), fill=CHARCOAL3)
    hline(s1, Inches(9.65), sy, Inches(3.45))
    txb(s1, num, Inches(9.82), sy + Inches(0.1), Inches(3.0), Inches(0.55),
        size=30, bold=True, color=TEAL)
    txb(s1, label, Inches(9.82), sy + Inches(0.65), Inches(3.0), Inches(0.38),
        size=9, color=WHITE, wrap=True)
    sy += Inches(1.18)

rect(s1, Inches(0.06), Inches(7.1), W - Inches(0.06), Inches(0.4), fill=CHARCOAL2)
txb(s1, "myrna.gamal@aucegypt.edu  ·  linkedin.com/in/myrna-gamalfahim  ·  Cairo, Egypt",
    Inches(0.3), Inches(7.15), Inches(9), Inches(0.28), size=9, color=DIM)
txb(s1, "Prepared for the Account Manager Role",
    Inches(9.65), Inches(7.15), Inches(3.4), Inches(0.28),
    size=9, color=DIM, align=PP_ALIGN.RIGHT)

# ── SLIDE 2: CAREER ─────────────────────────────────────────────
s2 = slide()
rect(s2, 0, 0, W, H, fill=CHARCOAL)
rect(s2, 0, 0, Inches(0.06), H, fill=TEAL)
rect(s2, Inches(0.06), 0, W - Inches(0.06), Inches(1.15), fill=CHARCOAL2)
eyebrow(s2, "Career Journey", Inches(0.3), Inches(0.1), Inches(10))
txb(s2, "Built on owning relationships.", Inches(0.3), Inches(0.32),
    Inches(12.5), Inches(0.5), size=22, bold=True, color=WHITE)
txb(s2, "Every role built on one thing: earning the trust of the person across the table and making them successful.",
    Inches(0.3), Inches(0.82), Inches(12.5), Inches(0.28), size=10, color=DIM)
hline(s2, Inches(0.3), Inches(1.18), Inches(1.4))

# Left skills panel
rect(s2, Inches(0.2), Inches(1.28), Inches(3.4), H - Inches(1.48), fill=CHARCOAL2)
txb(s2, "Core Skills", Inches(0.35), Inches(1.4), Inches(3.1), Inches(0.28),
    size=10, bold=True, color=TEAL)
skills = ["Team Leadership & Mentorship", "Enterprise Account Management",
          "Salesforce & CRM Governance", "Stakeholder Engagement",
          "Pipeline & Revenue Strategy", "AI-Led Selling"]
sy2 = Inches(1.82)
for sk in skills:
    rect(s2, Inches(0.35), sy2, Inches(3.0), Inches(0.38), fill=CHARCOAL3)
    txb(s2, sk, Inches(0.48), sy2 + Inches(0.05), Inches(2.8), Inches(0.3),
        size=10, color=WHITE)
    sy2 += Inches(0.46)

# Timeline roles
roles = [
    ("APR 2025 — PRESENT", "Senior Sales Development Executive", "SAP · Cairo, Egypt",
     "Senior SDE at SAP driving enterprise pipeline across Egypt. Senior stakeholder engagement, structured outreach, and close collaboration with AEs to move opportunities from first meeting to signed deal. Currently serving as team leader for MEA North.",
     "Team Leader · MEA North"),
    ("FEB 2022 — APR 2025", "Data & AI Technical Sales Specialist", "IBM · Cairo, Egypt",
     "Owned the full technical sales cycle for enterprise clients. Built ROI models, delivered executive-level demos and POCs. Became the trusted advisor for IBM business partners across Egypt, South Africa, and Turkey on all things Data & AI.",
     None),
    ("APR 2017 — FEB 2022", "IBM Cloud Developer Advocate Manager", "IBM · Cairo, Egypt",
     "Led a team of Developer Advocates across MEA. Drove 40% increase in active engagement and personally secured 200+ new startup partners. Built C-level relationships that shaped cloud adoption strategies region-wide.",
     "Team Manager · Developer Advocacy"),
]

tx = Inches(3.9); tw = Inches(9.15); ty = Inches(1.28)
for year, role, company, body, badge in roles:
    txb(s2, year, tx, ty, tw, Inches(0.22), size=8, bold=True, color=TEAL)
    txb(s2, role, tx, ty + Inches(0.22), tw, Inches(0.34), size=13, bold=True, color=WHITE)
    txb(s2, company, tx, ty + Inches(0.56), tw, Inches(0.22), size=9, color=TEAL)
    txb(s2, body, tx, ty + Inches(0.78), tw, Inches(0.72), size=9.5, color=DIM, wrap=True)
    if badge:
        rect(s2, tx, ty + Inches(1.52), Inches(2.3), Inches(0.25), fill=TEAL)
        txb(s2, badge, tx + Inches(0.05), ty + Inches(1.52), Inches(2.2), Inches(0.25),
            size=8, bold=True, color=CHARCOAL)
    hline(s2, tx, ty + Inches(1.86), tw, color=CHARCOAL3)
    ty += Inches(1.95)

rect(s2, Inches(0.2), Inches(7.05), W - Inches(0.3), Inches(0.35), fill=CHARCOAL3)
txb(s2, "BSc Computer Science, AUC (GPA 3.8)  ·  MBA, Edinburgh Business School (ongoing)  ·  Summer Abroad: UC Berkeley",
    Inches(0.35), Inches(7.1), W - Inches(0.6), Inches(0.28), size=9, color=DIM)

# ── SLIDE 3: ACHIEVEMENT ────────────────────────────────────────
s3 = slide()
section_header(s3, "Achievement I'm Proud Of",
               "Hitting target twice. The second time before I even had a full year.",
               "SAP · MEA North · 2025")

# Big stat
rect(s3, Inches(0.2), Inches(1.28), Inches(4.6), H - Inches(1.48), fill=CHARCOAL2)
hline(s3, Inches(0.2), Inches(1.28), Inches(4.6))
txb(s3, "140%", Inches(0.4), Inches(1.45), Inches(4.2), Inches(1.1),
    size=72, bold=True, color=TEAL)
txb(s3, "of sales target this year at SAP.\n100% the year before, joining\n4 months into the cycle.",
    Inches(0.4), Inches(2.62), Inches(4.1), Inches(0.85), size=11, color=WHITE, wrap=True)
hline(s3, Inches(0.4), Inches(3.6), Inches(4.0), color=CHARCOAL3)
txb(s3, '"I don\'t wait for conditions to be perfect. I build the conditions."',
    Inches(0.4), Inches(3.72), Inches(4.0), Inches(0.7),
    size=10, italic=True, color=DIM, wrap=True)

# Action blocks
actions = [
    ("100% in year one", "Full target delivered with only 8 months in role, new market and new product portfolio"),
    ("140% this year", "Exceeded target while simultaneously taking on team leadership for MEA North"),
    ("AI-powered prospecting", "Used AI to identify the right accounts and personas, turning cold outreach into precision targeting"),
    ("Lead and deliver", "Hit personal targets while leading the MEA North SDE team, proving I can do both at once"),
]
ax = Inches(5.1); aw = Inches(8.0); ay = Inches(1.33)
for title, body in actions:
    rect(s3, ax, ay, aw, Inches(1.2), fill=CHARCOAL2)
    hline(s3, ax, ay, Inches(0.7))
    txb(s3, "◆  " + title, ax + Inches(0.18), ay + Inches(0.1),
        aw - Inches(0.3), Inches(0.3), size=11, bold=True, color=TEAL)
    txb(s3, body, ax + Inches(0.18), ay + Inches(0.42),
        aw - Inches(0.3), Inches(0.68), size=10, color=WHITE, wrap=True)
    ay += Inches(1.3)

# ── SLIDE 4: VALUE ───────────────────────────────────────────────
s4 = slide()
section_header(s4, "Value to This Role",
               "I don't just manage accounts. I build the system that makes account management work.",
               "Every point below is grounded in something I have actually done.")

values = [
    ("Client Ownership & Cadence",
     "Owned enterprise relationships end-to-end at IBM and SAP: first call to executive review. 140% this year, 100% last year joining 4 months late. I know how to sustain a healthy account rhythm across a complex portfolio."),
    ("Trusted Advisor, Not Just a Vendor",
     "My strongest relationships are built on genuine understanding of the client's business and pressures. At IBM I became the go-to advisor for partners across three markets. Clients come back because they trust me."),
    ("AI-Led Selling",
     "I build AI into how I work as infrastructure. AI tools help me identify the right accounts and personas, prioritise outreach, and build prospecting applications. It is how I compete at 140%."),
    ("Partner Ecosystem",
     "Built and maintained partner relationships throughout my career — from 200+ startup partners in IBM's Global Entrepreneur Program to strategic alliances at SAP. I know how to activate a network to open doors and co-sell."),
    ("Senior Stakeholder Engagement",
     "From CTOs at IBM to C-suite at SAP. Executive-level relationships across Egypt, South Africa, and Morocco. I communicate with clarity, adapt to the room, and earn trust quickly."),
    ("Salesforce & Pipeline Governance",
     "Running pipeline management in Salesforce at SAP. Data hygiene matters. A CRM nobody trusts is just admin overhead. I make it a decision-making tool."),
]

vx0 = Inches(0.2); vy0 = Inches(1.28)
vw = Inches(4.22); vh = Inches(1.95); gap = Inches(0.12)

for idx, (title, body) in enumerate(values):
    row, col = divmod(idx, 3)
    vx = vx0 + col * (vw + gap)
    vy = vy0 + row * (vh + gap)
    rect(s4, vx, vy, vw, vh, fill=CHARCOAL2)
    hline(s4, vx, vy, vw)
    txb(s4, title, vx + Inches(0.18), vy + Inches(0.1),
        vw - Inches(0.25), Inches(0.35), size=11, bold=True, color=WHITE)
    txb(s4, body, vx + Inches(0.18), vy + Inches(0.48),
        vw - Inches(0.25), Inches(1.35), size=9.5, color=DIM, wrap=True)

# ── SLIDE 5: WHY ─────────────────────────────────────────────────
s5 = slide()
rect(s5, 0, 0, W, H, fill=CHARCOAL)
rect(s5, 0, 0, Inches(0.06), H, fill=TEAL)

# Quote block
rect(s5, Inches(0.2), Inches(0.55), W - Inches(0.4), Inches(2.65), fill=CHARCOAL2)
hline(s5, Inches(0.2), Inches(0.55), W - Inches(0.4))
txb(s5, "“", Inches(0.35), Inches(0.5), Inches(0.9), Inches(1.1),
    size=64, bold=True, color=TEAL)
txb(s5, "When I look back at what I'm most proud of, it's never just the number. It's the client who called me first when something went wrong. The partner who trusted me with their roadmap. The relationship that outlasted the deal. That's what I want to build here. At depth, over time.",
    Inches(1.05), Inches(0.72), W - Inches(1.35), Inches(1.8),
    size=14, bold=True, color=WHITE, wrap=True)
txb(s5, "— Myrna Gamal Fahim",
    Inches(1.05), Inches(2.88), W - Inches(1.35), Inches(0.28),
    size=10, color=DIM)

# Two reason blocks
reasons = [
    ("I want to go deeper, not just wider.",
     "I've spent 10 years building pipeline and winning new business. Now I want to go deeper: own a portfolio of clients over time, understand their business fully, and become the person they rely on. That is a different and harder skill. I'm ready for it."),
    ("The best account managers are obsessed with client success.",
     "I don't just want to retain accounts. I want clients to grow, to refer others, to say this firm understands them. That obsession with genuine client success (not just contract renewal) is what I bring to this role."),
]
rx = Inches(0.2); ry = Inches(3.35); rw = (W - Inches(0.55)) / 2
for i, (head, body) in enumerate(reasons):
    rx2 = rx + i * (rw + Inches(0.12))
    rect(s5, rx2, ry, rw, Inches(1.6), fill=CHARCOAL2)
    hline(s5, rx2, ry, rw)
    txb(s5, head, rx2 + Inches(0.18), ry + Inches(0.1),
        rw - Inches(0.3), Inches(0.35), size=11, bold=True, color=WHITE, wrap=True)
    txb(s5, body, rx2 + Inches(0.18), ry + Inches(0.5),
        rw - Inches(0.3), Inches(1.0), size=9.5, color=DIM, wrap=True)

# 30/60/90 strip
commits = [
    ("Day 1",    "Listen & map", "Understand the client landscape, team dynamics, and existing rhythms"),
    ("30 Days",  "Audit & diagnose", "Salesforce data quality, reporting cadence, stakeholder coverage"),
    ("90 Days",  "Embed & deliver", "Structured account cadence in place, first client value report delivered"),
    ("6 Months", "Measure & grow", "Visible uplift in feedback volumes, pipeline accuracy, on-time reporting"),
]
cy = Inches(5.12); cw = (W - Inches(0.55)) / 4
for i, (period, strong, body) in enumerate(commits):
    cx = Inches(0.2) + i * (cw + Inches(0.05))
    rect(s5, cx, cy, cw, Inches(1.7), fill=CHARCOAL2)
    hline(s5, cx, cy, cw)
    txb(s5, period, cx + Inches(0.15), cy + Inches(0.1),
        cw - Inches(0.25), Inches(0.25), size=9, bold=True, color=TEAL)
    txb(s5, strong, cx + Inches(0.15), cy + Inches(0.38),
        cw - Inches(0.25), Inches(0.28), size=11, bold=True, color=WHITE)
    txb(s5, body, cx + Inches(0.15), cy + Inches(0.7),
        cw - Inches(0.25), Inches(0.85), size=9.5, color=DIM, wrap=True)

rect(s5, Inches(0.06), Inches(7.1), W - Inches(0.06), Inches(0.4), fill=CHARCOAL2)
txb(s5, "Thank you. I welcome the opportunity to discuss how my background can contribute as your next Account Manager.",
    Inches(0.3), Inches(7.15), W - Inches(0.6), Inches(0.28), size=9, color=DIM)

OUT = "/Users/i763300/Desktop/ClaudeProjects/account-mgr-presentation/AccountManager_IC_Presentation.pptx"
prs.save(OUT)
print("Saved:", OUT)
