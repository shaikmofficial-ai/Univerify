#!/usr/bin/env python3
"""
UniVerify V2.0 - Final Project Presentation Generator
Generates a professional, IIT-style technical presentation (.pptx)
Theme matches the live app: dark GitHub-style palette with blue/green accents.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# THEME / PALETTE  (inspired by the UniVerify UI)
# ----------------------------------------------------------------------------
BG_DARK      = RGBColor(0x0D, 0x11, 0x17)   # deep background
BG_PANEL     = RGBColor(0x16, 0x1B, 0x22)   # card / panel
BG_PANEL_2   = RGBColor(0x1C, 0x22, 0x2B)   # alt panel
ACCENT_BLUE  = RGBColor(0x58, 0xA6, 0xFF)   # primary blue
ACCENT_GREEN = RGBColor(0x3F, 0xB9, 0x50)   # success green
ACCENT_GREEN2= RGBColor(0x23, 0x86, 0x36)
TEXT_WHITE   = RGBColor(0xF0, 0xF6, 0xFC)
TEXT_GREY    = RGBColor(0x8B, 0x94, 0x9E)
BORDER_BLUE  = RGBColor(0x30, 0x4A, 0x6E)
LINE_SUBTLE  = RGBColor(0x30, 0x36, 0x3D)
PURPLE       = RGBColor(0xBC, 0x8C, 0xFF)
ORANGE       = RGBColor(0xF7, 0x8A, 0x3B)

FONT_TITLE = "Segoe UI Semibold"
FONT_HEAD  = "Segoe UI"
FONT_BODY  = "Segoe UI"
FONT_MONO  = "Consolas"

# 16:9 widescreen
SW = Inches(13.333)
SH = Inches(7.5)

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------------
def add_slide():
    return prs.slides.add_slide(BLANK)


def bg(slide, color=BG_DARK):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = color


def rect(slide, x, y, w, h, color, line_color=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    if line_color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line_color
        sp.line.width = line_w or Pt(1)
    sp.shadow.inherit = False
    return sp


def no_autofit(tf):
    # prevent text auto-resizing weirdness
    try:
        tf.word_wrap = True
    except Exception:
        pass


def txt(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """lines: list of dicts {text, size, color, bold, font, space_after, level, bullet}"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ln.get("align", align)
        if ln.get("space_after") is not None:
            p.space_after = Pt(ln["space_after"])
        if ln.get("space_before") is not None:
            p.space_before = Pt(ln["space_before"])
        if ln.get("line_spacing"):
            p.line_spacing = ln["line_spacing"]
        run = p.add_run()
        run.text = ln["text"]
        f = run.font
        f.size = Pt(ln.get("size", 18))
        f.bold = ln.get("bold", False)
        f.italic = ln.get("italic", False)
        f.name = ln.get("font", FONT_BODY)
        f.color.rgb = ln.get("color", TEXT_WHITE)
    return tb


def header(slide, title, kicker="UNIVERIFY V2.0", num=None):
    """Standard content-slide header with accent bar + kicker + title."""
    # top accent bar
    rect(slide, 0, 0, SW, Inches(0.12), ACCENT_BLUE)
    # side accent block
    rect(slide, Inches(0.55), Inches(0.55), Inches(0.13), Inches(0.95), ACCENT_BLUE)
    txt(slide, Inches(0.85), Inches(0.5), Inches(10), Inches(0.35),
        [{"text": kicker, "size": 12, "color": ACCENT_BLUE, "bold": True, "font": FONT_HEAD}])
    txt(slide, Inches(0.83), Inches(0.78), Inches(11.6), Inches(0.8),
        [{"text": title, "size": 30, "color": TEXT_WHITE, "bold": True, "font": FONT_TITLE}])
    # footer
    rect(slide, 0, SH - Inches(0.06), SW, Inches(0.06), LINE_SUBTLE)
    txt(slide, Inches(0.55), SH - Inches(0.42), Inches(8), Inches(0.3),
        [{"text": "Decentralized Academic Credential Verification  •  Blockchain Mini Project",
          "size": 9, "color": TEXT_GREY, "font": FONT_BODY}])
    if num is not None:
        txt(slide, SW - Inches(1.3), SH - Inches(0.42), Inches(0.8), Inches(0.3),
            [{"text": f"{num:02d}", "size": 11, "color": ACCENT_BLUE, "bold": True, "align": PP_ALIGN.RIGHT}],
            align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, fill=BG_PANEL, border=LINE_SUBTLE, bw=Pt(1), round=True):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if round else MSO_SHAPE.RECTANGLE
    sp = rect(slide, x, y, w, h, fill, border, bw, shape=shape)
    if round:
        try:
            sp.adjustments[0] = 0.06
        except Exception:
            pass
    return sp


def bullet_block(slide, x, y, w, h, items, size=16, gap=10, color=TEXT_WHITE,
                 marker="▸", marker_color=ACCENT_BLUE):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        p.line_spacing = 1.05
        # marker run
        r1 = p.add_run()
        r1.text = marker + "  "
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = marker_color
        r1.font.name = FONT_BODY
        # bold lead support: "Lead::rest"
        if "::" in it:
            lead, rest = it.split("::", 1)
            rb = p.add_run()
            rb.text = lead + "  "
            rb.font.size = Pt(size)
            rb.font.bold = True
            rb.font.color.rgb = TEXT_WHITE
            rb.font.name = FONT_BODY
            rr = p.add_run()
            rr.text = rest.strip()
            rr.font.size = Pt(size)
            rr.font.color.rgb = color
            rr.font.name = FONT_BODY
        else:
            r2 = p.add_run()
            r2.text = it
            r2.font.size = Pt(size)
            r2.font.color.rgb = color
            r2.font.name = FONT_BODY
    return tb


def code_block(slide, x, y, w, h, code_lines, title="solidity"):
    card(slide, x, y, w, h, fill=RGBColor(0x0A, 0x0E, 0x14), border=BORDER_BLUE, bw=Pt(1.25))
    # title bar dots
    for i, c in enumerate([RGBColor(0xF8,0x51,0x49), RGBColor(0xF7,0x8A,0x3B), ACCENT_GREEN]):
        rect(slide, x + Inches(0.25) + Inches(0.28)*i, y + Inches(0.22), Inches(0.14), Inches(0.14),
             c, shape=MSO_SHAPE.OVAL)
    txt(slide, x + Inches(1.3), y + Inches(0.13), w - Inches(1.5), Inches(0.3),
        [{"text": title, "size": 11, "color": TEXT_GREY, "font": FONT_MONO, "align": PP_ALIGN.RIGHT}],
        align=PP_ALIGN.RIGHT)
    tb = slide.shapes.add_textbox(x + Inches(0.3), y + Inches(0.55), w - Inches(0.6), h - Inches(0.7))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (line, col) in enumerate(code_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(2)
        p.line_spacing = 1.0
        r = p.add_run()
        r.text = line if line else " "
        r.font.size = Pt(11.5)
        r.font.name = FONT_MONO
        r.font.color.rgb = col
    return tb


# ============================================================================
# SLIDE 1 — TITLE
# ============================================================================
s = add_slide()
bg(s)
# decorative chain blocks across the top-right
import random
random.seed(7)
for i in range(7):
    bx = Inches(7.2 + i*0.78)
    rect(s, bx, Inches(0.45), Inches(0.5), Inches(0.5),
         BG_PANEL_2, BORDER_BLUE, Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    if i < 6:
        rect(s, bx + Inches(0.5), Inches(0.67), Inches(0.28), Inches(0.06), ACCENT_BLUE)
# big faint watermark hex
rect(s, Inches(9.6), Inches(4.6), Inches(3.3), Inches(3.3),
     BG_PANEL, shape=MSO_SHAPE.HEXAGON)
# left accent vertical
rect(s, 0, 0, Inches(0.18), SH, ACCENT_BLUE)

txt(s, Inches(0.9), Inches(1.7), Inches(11), Inches(0.4),
    [{"text": "BLOCKCHAIN  •  WEB3  •  MINI PROJECT", "size": 14, "color": ACCENT_BLUE,
      "bold": True, "font": FONT_HEAD}])
txt(s, Inches(0.85), Inches(2.15), Inches(11.6), Inches(1.6),
    [{"text": "UniVerify V2.0", "size": 70, "color": TEXT_WHITE, "bold": True, "font": FONT_TITLE}])
txt(s, Inches(0.9), Inches(3.5), Inches(11), Inches(0.7),
    [{"text": "Decentralized Academic Credential Verification Portal", "size": 24,
      "color": ACCENT_GREEN, "bold": True, "font": FONT_HEAD}])
txt(s, Inches(0.9), Inches(4.25), Inches(10.5), Inches(0.6),
    [{"text": "Eliminating certificate forgery using immutable Ethereum smart contracts.",
      "size": 15, "color": TEXT_GREY, "italic": True}])

# info chips
chips = [("Solidity", ACCENT_BLUE), ("Ethers.js", ACCENT_GREEN),
         ("Sepolia Testnet", PURPLE), ("MetaMask", ORANGE)]
cx = Inches(0.9)
for label, col in chips:
    w = Inches(0.25 + 0.13*len(label))
    c = card(s, cx, Inches(5.05), w, Inches(0.5), fill=BG_PANEL, border=col, bw=Pt(1.25))
    txt(s, cx, Inches(5.12), w, Inches(0.36),
        [{"text": label, "size": 12, "color": col, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += w + Inches(0.2)

rect(s, Inches(0.9), Inches(5.95), Inches(4.2), Pt(2), ACCENT_BLUE)
txt(s, Inches(0.9), Inches(6.1), Inches(11.5), Inches(1.2),
    [{"text": "Presented by:", "size": 13, "color": ACCENT_BLUE, "bold": True, "space_after": 2},
     {"text": "Shaik Abdulla M  (231061101493)  \u2022  CSE", "size": 14, "color": TEXT_WHITE, "bold": True},
     {"text": "Narendra Reddy  (231061101489)  \u2022  CSE", "size": 14, "color": TEXT_WHITE, "bold": True, "space_after": 3},
     {"text": "Department of Computer Science & Engineering  |  Final Project Review",
      "size": 11.5, "color": TEXT_GREY}])

# ============================================================================
# SLIDE 2 — AGENDA / OUTLINE
# ============================================================================
s = add_slide(); bg(s)
header(s, "Presentation Outline", "ROADMAP", 2)
agenda = [
    ("01", "Abstract", "Project at a glance"),
    ("02", "Problem Statement", "Why forgery is a real threat"),
    ("03", "Existing System", "Limitations of current methods"),
    ("04", "Proposed Solution", "Objectives & approach"),
    ("05", "System Architecture", "End-to-end design"),
    ("06", "Technology Stack", "Tools & frameworks"),
    ("07", "Smart Contract", "On-chain logic"),
    ("08", "Working / Flow", "Issuance & verification"),
    ("09", "Implementation", "Frontend + Web3 bridge"),
    ("10", "Results", "Live demonstration"),
    ("11", "Advantages & Scope", "Benefits & future work"),
    ("12", "Conclusion", "Key takeaways"),
]
colw = Inches(3.85)
gx = Inches(0.7); gy = Inches(1.95)
for i, (n, t, sub) in enumerate(agenda):
    col = i % 3
    row = i // 3
    x = gx + col * (colw + Inches(0.18))
    y = gy + row * Inches(1.18)
    card(s, x, y, colw, Inches(1.0), fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, x, y, Inches(0.09), Inches(1.0), ACCENT_BLUE)
    txt(s, x + Inches(0.25), y + Inches(0.13), Inches(1.0), Inches(0.7),
        [{"text": n, "size": 26, "color": ACCENT_BLUE, "bold": True, "font": FONT_TITLE}])
    txt(s, x + Inches(1.15), y + Inches(0.16), colw - Inches(1.3), Inches(0.8),
        [{"text": t, "size": 15, "color": TEXT_WHITE, "bold": True},
         {"text": sub, "size": 10.5, "color": TEXT_GREY, "space_before": 1}])

# ============================================================================
# SLIDE 3 — ABSTRACT
# ============================================================================
s = add_slide(); bg(s)
header(s, "Abstract", "PROJECT AT A GLANCE", 3)
card(s, Inches(0.7), Inches(1.95), Inches(7.4), Inches(4.7), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.25))
txt(s, Inches(1.05), Inches(2.25), Inches(6.8), Inches(4.2),
    [
     {"text": "UniVerify V2.0 is a Web3 application engineered to eliminate academic credential forgery.",
      "size": 17, "color": TEXT_WHITE, "bold": True, "line_spacing": 1.15, "space_after": 10},
     {"text": "Educational institutions anchor degree metadata to the Ethereum Sepolia Testnet using cryptographic hashes. Once a transaction is mined, the record becomes permanent and tamper-proof.",
      "size": 14, "color": TEXT_GREY, "line_spacing": 1.2, "space_after": 10},
     {"text": "Verification is instant: a recruiter or institution enters a student's Register Number, and the system queries the blockchain in O(1) time via Solidity mappings, returning an authentic, cryptographically-backed record \u2014 with no central authority required.",
      "size": 14, "color": TEXT_GREY, "line_spacing": 1.2, "space_after": 10},
     {"text": "The result is a trustless, transparent and globally verifiable system for academic certificates.",
      "size": 14, "color": ACCENT_GREEN, "italic": True, "line_spacing": 1.2},
    ])
# right stats column
stats = [("100%", "Tamper-proof records", ACCENT_GREEN),
         ("O(1)", "Search complexity", ACCENT_BLUE),
         ("0", "Central authorities needed", PURPLE),
         ("24/7", "Global verification", ORANGE)]
sy = Inches(1.95)
for val, lab, col in stats:
    card(s, Inches(8.35), sy, Inches(4.25), Inches(1.0), fill=BG_PANEL_2, border=col, bw=Pt(1.25))
    txt(s, Inches(8.6), sy + Inches(0.12), Inches(1.7), Inches(0.8),
        [{"text": val, "size": 28, "color": col, "bold": True, "font": FONT_TITLE}],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(10.2), sy + Inches(0.12), Inches(2.3), Inches(0.8),
        [{"text": lab, "size": 13, "color": TEXT_WHITE, "bold": True}],
        anchor=MSO_ANCHOR.MIDDLE)
    sy += Inches(1.18)

# ============================================================================
# SLIDE 4 — PROBLEM STATEMENT
# ============================================================================
s = add_slide(); bg(s)
header(s, "Problem Statement", "THE CHALLENGE", 4)
txt(s, Inches(0.7), Inches(1.85), Inches(11.9), Inches(0.5),
    [{"text": "Fake degrees and certificate fraud cost organizations time, money and trust worldwide.",
      "size": 16, "color": TEXT_GREY, "italic": True}])
probs = [
    ("Rampant Forgery", "Certificates are easily photoshopped or printed. Manual seals & signatures offer weak protection.", RGBColor(0xF8,0x51,0x49)),
    ("Slow Verification", "Employers must email or call universities and wait days/weeks for manual confirmation.", ORANGE),
    ("Single Point of Failure", "If a university's database is lost, hacked or shut down, records become unverifiable.", PURPLE),
    ("No Global Standard", "Cross-border verification is inconsistent, costly and prone to manipulation.", ACCENT_BLUE),
]
gx = Inches(0.7); gy = Inches(2.55); cw = Inches(5.85); ch = Inches(1.85)
for i, (t, d, col) in enumerate(probs):
    x = gx + (i % 2) * (cw + Inches(0.25))
    y = gy + (i // 2) * (ch + Inches(0.25))
    card(s, x, y, cw, ch, fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, x, y, cw, Inches(0.09), col)
    rect(s, x + Inches(0.3), y + Inches(0.38), Inches(0.55), Inches(0.55), col, shape=MSO_SHAPE.OVAL)
    txt(s, x + Inches(0.3), y + Inches(0.38), Inches(0.55), Inches(0.55),
        [{"text": "!", "size": 24, "color": BG_DARK, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(1.05), y + Inches(0.3), cw - Inches(1.3), Inches(0.5),
        [{"text": t, "size": 18, "color": TEXT_WHITE, "bold": True}])
    txt(s, x + Inches(1.05), y + Inches(0.78), cw - Inches(1.3), Inches(1.0),
        [{"text": d, "size": 13, "color": TEXT_GREY, "line_spacing": 1.15}])

# ============================================================================
# SLIDE 5 — EXISTING SYSTEM vs PROPOSED
# ============================================================================
s = add_slide(); bg(s)
header(s, "Existing System vs. Proposed System", "THE GAP", 5)
# Existing
card(s, Inches(0.7), Inches(1.95), Inches(5.85), Inches(4.7), fill=BG_PANEL, border=RGBColor(0xF8,0x51,0x49), bw=Pt(1.25))
txt(s, Inches(0.95), Inches(2.15), Inches(5.4), Inches(0.5),
    [{"text": "❌  Traditional / Existing System", "size": 18, "color": RGBColor(0xF8,0x51,0x49), "bold": True}])
bullet_block(s, Inches(1.0), Inches(2.85), Inches(5.3), Inches(3.6),
    ["Paper certificates with physical seals",
     "Centralized university databases",
     "Manual, email-based verification",
     "Easily forged or duplicated",
     "Vulnerable to data loss & hacking",
     "No transparency for third parties",
     "Time-consuming (days to weeks)"],
    size=14, gap=12, color=TEXT_GREY, marker="✗", marker_color=RGBColor(0xF8,0x51,0x49))
# Proposed
card(s, Inches(6.8), Inches(1.95), Inches(5.85), Inches(4.7), fill=BG_PANEL, border=ACCENT_GREEN, bw=Pt(1.25))
txt(s, Inches(7.05), Inches(2.15), Inches(5.4), Inches(0.5),
    [{"text": "✅  UniVerify (Proposed)", "size": 18, "color": ACCENT_GREEN, "bold": True}])
bullet_block(s, Inches(7.1), Inches(2.85), Inches(5.3), Inches(3.6),
    ["Digital records on Ethereum blockchain",
     "Decentralized & distributed ledger",
     "Instant, self-service verification",
     "Cryptographically tamper-proof",
     "Immutable \u2014 cannot be altered/deleted",
     "Fully transparent & publicly auditable",
     "Real-time (seconds) verification"],
    size=14, gap=12, color=TEXT_WHITE, marker="✓", marker_color=ACCENT_GREEN)

# ============================================================================
# SLIDE 6 — OBJECTIVES
# ============================================================================
s = add_slide(); bg(s)
header(s, "Objectives of the Project", "GOALS", 6)
objs = [
    ("Immutability", "Anchor credential metadata on-chain so records are permanent and cannot be altered once mined.", ACCENT_BLUE),
    ("Instant Verification", "Provide O(1) lookup of certificates using Solidity mappings keyed by Register Number.", ACCENT_GREEN),
    ("Trustless Authority", "Restrict issuance to the university wallet via on-chain access control \u2014 no middlemen.", PURPLE),
    ("Accessible UI", "Deliver a clean Web3 portal where anyone can verify a degree with just a register number.", ORANGE),
    ("Security", "Use MetaMask signing for administrative actions and cryptographic hashes for file integrity.", RGBColor(0xF8,0x51,0x49)),
    ("Cost-Efficient Demo", "Deploy on Ethereum Sepolia Testnet for a zero-cost, production-like environment.", ACCENT_BLUE),
]
gx = Inches(0.7); gy = Inches(1.95); cw = Inches(3.85); ch = Inches(2.18)
for i, (t, d, col) in enumerate(objs):
    x = gx + (i % 3) * (cw + Inches(0.18))
    y = gy + (i // 3) * (ch + Inches(0.2))
    card(s, x, y, cw, ch, fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, x + Inches(0.28), y + Inches(0.28), Inches(0.7), Inches(0.7), BG_PANEL_2, col, Pt(1.5),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, x + Inches(0.28), y + Inches(0.28), Inches(0.7), Inches(0.7),
        [{"text": str(i+1), "size": 26, "color": col, "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_TITLE}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(0.28), y + Inches(1.1), cw - Inches(0.5), Inches(0.4),
        [{"text": t, "size": 16, "color": TEXT_WHITE, "bold": True}])
    txt(s, x + Inches(0.28), y + Inches(1.5), cw - Inches(0.5), Inches(0.6),
        [{"text": d, "size": 11.5, "color": TEXT_GREY, "line_spacing": 1.1}])

# ============================================================================
# SLIDE 7 — SYSTEM ARCHITECTURE
# ============================================================================
s = add_slide(); bg(s)
header(s, "System Architecture", "HOW IT FITS TOGETHER", 7)

def arch_node(slide, x, y, w, h, title, sub, col):
    card(slide, x, y, w, h, fill=BG_PANEL, border=col, bw=Pt(1.5))
    txt(slide, x, y + Inches(0.18), w, Inches(0.4),
        [{"text": title, "size": 15, "color": col, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER)
    txt(slide, x + Inches(0.1), y + Inches(0.62), w - Inches(0.2), h - Inches(0.7),
        [{"text": sub, "size": 11, "color": TEXT_GREY, "align": PP_ALIGN.CENTER, "line_spacing": 1.1}],
        align=PP_ALIGN.CENTER)

def arrow(slide, x, y, w, color=ACCENT_BLUE):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, Inches(0.4))
    a.fill.solid(); a.fill.fore_color.rgb = color
    a.line.fill.background(); a.shadow.inherit = False

ny = Inches(2.5); nh = Inches(1.7); nw = Inches(2.45)
arch_node(s, Inches(0.6), ny, nw, nh, "User / Recruiter", "Enters Register Number in the web portal", ACCENT_BLUE)
arrow(s, Inches(3.12), ny + Inches(0.65), Inches(0.55))
arch_node(s, Inches(3.75), ny, nw, nh, "Frontend (DApp)", "HTML / CSS / JS\nUI & input handling", ACCENT_GREEN)
arrow(s, Inches(6.27), ny + Inches(0.65), Inches(0.55))
arch_node(s, Inches(6.9), ny, nw, nh, "Ethers.js + MetaMask", "Web3 bridge signs & sends calls", PURPLE)
arrow(s, Inches(9.42), ny + Inches(0.65), Inches(0.55))
arch_node(s, Inches(10.05), ny, Inches(2.65), nh, "Smart Contract", "AcademicCertificates on Sepolia", ORANGE)

# blockchain ledger strip
txt(s, Inches(0.6), Inches(4.6), Inches(11), Inches(0.35),
    [{"text": "Ethereum Sepolia Testnet  —  Immutable Distributed Ledger", "size": 13, "color": ACCENT_BLUE, "bold": True}])
bx = Inches(0.6)
for i in range(8):
    rect(s, bx, Inches(5.0), Inches(1.35), Inches(0.9), BG_PANEL_2, BORDER_BLUE, Pt(1.25),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, bx, Inches(5.12), Inches(1.35), Inches(0.3),
        [{"text": f"Block #{i+1}", "size": 9.5, "color": ACCENT_GREEN, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER)
    txt(s, bx, Inches(5.42), Inches(1.35), Inches(0.4),
        [{"text": "0x" + "%04x" % (i*4231 % 65535), "size": 9, "color": TEXT_GREY, "align": PP_ALIGN.CENTER, "font": FONT_MONO}],
        align=PP_ALIGN.CENTER)
    if i < 7:
        rect(s, bx + Inches(1.35), Inches(5.4), Inches(0.18), Inches(0.08), ACCENT_BLUE)
    bx += Inches(1.53)
txt(s, Inches(0.6), Inches(6.05), Inches(11.9), Inches(0.4),
    [{"text": "Each block is cryptographically linked to the previous one — altering any record breaks the entire chain.",
      "size": 11.5, "color": TEXT_GREY, "italic": True}])

# ============================================================================
# SLIDE 8 — TECHNOLOGY STACK
# ============================================================================
s = add_slide(); bg(s)
header(s, "Technology Stack", "TOOLS & FRAMEWORKS", 8)
tech = [
    ("Solidity", "Smart Contract Language", "On-chain logic for issuing & retrieving certificates", ACCENT_BLUE),
    ("Ethereum (Sepolia)", "Blockchain Network", "Public test network hosting the immutable ledger", PURPLE),
    ("Ethers.js v5.7.2", "Web3 Library", "Connects the frontend to the blockchain", ACCENT_GREEN),
    ("MetaMask", "Crypto Wallet", "Signs administrative transactions securely", ORANGE),
    ("HTML / CSS", "Frontend UI", "Responsive, glass-morphism verification portal", ACCENT_BLUE),
    ("JavaScript", "Client Logic", "Handles user input and async contract calls", ACCENT_GREEN),
    ("Remix IDE", "Development Env", "Writing, compiling & deploying contracts", PURPLE),
    ("IPFS Hash", "File Reference", "Links the on-chain record to the actual document", ORANGE),
]
gx = Inches(0.7); gy = Inches(1.95); cw = Inches(2.92); ch = Inches(2.18)
for i, (name, role, desc, col) in enumerate(tech):
    x = gx + (i % 4) * (cw + Inches(0.16))
    y = gy + (i // 4) * (ch + Inches(0.22))
    card(s, x, y, cw, ch, fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, x, y, cw, Inches(0.08), col)
    txt(s, x + Inches(0.22), y + Inches(0.28), cw - Inches(0.4), Inches(0.5),
        [{"text": name, "size": 15.5, "color": TEXT_WHITE, "bold": True}])
    txt(s, x + Inches(0.22), y + Inches(0.78), cw - Inches(0.4), Inches(0.35),
        [{"text": role.upper(), "size": 10, "color": col, "bold": True}])
    txt(s, x + Inches(0.22), y + Inches(1.2), cw - Inches(0.4), Inches(0.9),
        [{"text": desc, "size": 11.5, "color": TEXT_GREY, "line_spacing": 1.12}])

# ============================================================================
# SLIDE 9 — SMART CONTRACT
# ============================================================================
s = add_slide(); bg(s)
header(s, "Smart Contract Design", "ON-CHAIN LOGIC", 9)
code = [
    ("// SPDX-License-Identifier: MIT", TEXT_GREY),
    ("pragma solidity ^0.8.0;", PURPLE),
    ("", TEXT_WHITE),
    ("contract AcademicCertificates {", ACCENT_BLUE),
    ("    struct Certificate {", ACCENT_GREEN),
    ("        string studentName;", TEXT_WHITE),
    ("        string courseName;", TEXT_WHITE),
    ("        string ipfsHash;   // link to PDF", TEXT_GREY),
    ("    }", ACCENT_GREEN),
    ("    mapping(uint256 => Certificate) public certificates;", TEXT_WHITE),
    ("    address public university;  // contract owner", TEXT_WHITE),
    ("", TEXT_WHITE),
    ("    function issueDegree(...) public {", ORANGE),
    ("        require(msg.sender == university,", RGBColor(0xF8,0x51,0x49)),
    ('          "Only the University can issue!");', RGBColor(0xF8,0x51,0x49)),
    ("        certificates[nextId] = Certificate(...);", TEXT_WHITE),
    ("    }", ORANGE),
    ("}", ACCENT_BLUE),
]
code_block(s, Inches(0.7), Inches(1.95), Inches(7.0), Inches(4.75), code, title="AcademicCert.sol")
# right explanation
notes = [
    "Struct::groups student name, course & IPFS file hash into one record.",
    "Mapping::stores certificates keyed by ID for O(1) instant retrieval.",
    "Owner Lock::`university` address is set in the constructor at deploy time.",
    "Access Control::`require()` ensures only the university wallet can issue degrees.",
    "Immutability::once a transaction is mined, the record is permanent on-chain.",
]
card(s, Inches(7.9), Inches(1.95), Inches(4.7), Inches(4.75), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.25))
txt(s, Inches(8.15), Inches(2.15), Inches(4.2), Inches(0.4),
    [{"text": "Key Design Points", "size": 16, "color": ACCENT_BLUE, "bold": True}])
bullet_block(s, Inches(8.15), Inches(2.7), Inches(4.25), Inches(3.9), notes,
             size=12.5, gap=13, color=TEXT_GREY)

# ============================================================================
# SLIDE 10 — WORKING / FLOW
# ============================================================================
s = add_slide(); bg(s)
header(s, "Working Methodology", "ISSUANCE & VERIFICATION", 10)
# Issuance flow
txt(s, Inches(0.7), Inches(1.85), Inches(11), Inches(0.4),
    [{"text": "A.  Certificate Issuance  (Admin only)", "size": 16, "color": ACCENT_GREEN, "bold": True}])
issue_steps = ["University connects MetaMask wallet", "Enters Reg No, Name, Course & file hash",
               "Signs the transaction", "Record mined & stored on-chain"]
sx = Inches(0.7)
for i, st in enumerate(issue_steps):
    card(s, sx, Inches(2.35), Inches(2.7), Inches(1.25), fill=BG_PANEL, border=ACCENT_GREEN, bw=Pt(1.25))
    txt(s, sx + Inches(0.2), Inches(2.5), Inches(0.6), Inches(0.5),
        [{"text": str(i+1), "size": 22, "color": ACCENT_GREEN, "bold": True, "font": FONT_TITLE}])
    txt(s, sx + Inches(0.2), Inches(2.95), Inches(2.35), Inches(0.6),
        [{"text": st, "size": 11.5, "color": TEXT_WHITE, "line_spacing": 1.1}])
    if i < 3:
        arrow(s, sx + Inches(2.72), Inches(2.83), Inches(0.28), ACCENT_GREEN)
    sx += Inches(3.0)

# Verification flow
txt(s, Inches(0.7), Inches(3.95), Inches(11), Inches(0.4),
    [{"text": "B.  Certificate Verification  (Anyone)", "size": 16, "color": ACCENT_BLUE, "bold": True}])
ver_steps = ["User enters Register Number", "Ethers.js calls getCertificate()",
             "Contract returns stored data", "Authentic record displayed instantly"]
sx = Inches(0.7)
for i, st in enumerate(ver_steps):
    card(s, sx, Inches(4.45), Inches(2.7), Inches(1.25), fill=BG_PANEL, border=ACCENT_BLUE, bw=Pt(1.25))
    txt(s, sx + Inches(0.2), Inches(4.6), Inches(0.6), Inches(0.5),
        [{"text": str(i+1), "size": 22, "color": ACCENT_BLUE, "bold": True, "font": FONT_TITLE}])
    txt(s, sx + Inches(0.2), Inches(5.05), Inches(2.35), Inches(0.6),
        [{"text": st, "size": 11.5, "color": TEXT_WHITE, "line_spacing": 1.1}])
    if i < 3:
        arrow(s, sx + Inches(2.72), Inches(4.93), Inches(0.28), ACCENT_BLUE)
    sx += Inches(3.0)

txt(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(0.5),
    [{"text": "If no record exists for the Register Number, the contract call reverts and a clear \u201cVerification Failed\u201d message is shown.",
      "size": 12, "color": TEXT_GREY, "italic": True}])

# ============================================================================
# SLIDE 11 — IMPLEMENTATION (Frontend + Web3)
# ============================================================================
s = add_slide(); bg(s)
header(s, "Implementation — Web3 Bridge", "FRONTEND INTEGRATION", 11)
js_code = [
    ("// Connect to MetaMask provider", TEXT_GREY),
    ("const provider =", ACCENT_BLUE),
    ("  new ethers.providers.Web3Provider(", TEXT_WHITE),
    ("     window.ethereum);", TEXT_WHITE),
    ("", TEXT_WHITE),
    ("// Build contract instance", TEXT_GREY),
    ("const contract = new ethers.Contract(", ACCENT_GREEN),
    ("   contractAddress, abi, provider);", TEXT_WHITE),
    ("", TEXT_WHITE),
    ("// Query the blockchain (O(1))", TEXT_GREY),
    ("const data =", ORANGE),
    ("  await contract.getCertificate(regNo);", TEXT_WHITE),
    ("", TEXT_WHITE),
    ("// data[0]=Name  data[1]=Course  data[2]=Hash", PURPLE),
]
code_block(s, Inches(0.7), Inches(1.95), Inches(6.7), Inches(4.75), js_code, title="script.js")
card(s, Inches(7.6), Inches(1.95), Inches(5.0), Inches(4.75), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.25))
txt(s, Inches(7.85), Inches(2.15), Inches(4.5), Inches(0.4),
    [{"text": "Implementation Highlights", "size": 16, "color": ACCENT_BLUE, "bold": True}])
bullet_block(s, Inches(7.85), Inches(2.7), Inches(4.5), Inches(3.9),
    ["Contract Address::hard-coded deployed V2 address on Sepolia.",
     "ABI::defines issueCertificate & getCertificate interfaces.",
     "Async/Await::non-blocking calls keep the UI responsive.",
     "Error Handling::try/catch detects reverted (non-existent) records.",
     "Dynamic UI::results rendered as success / error states with color cues."],
    size=12.5, gap=12, color=TEXT_GREY)

# ============================================================================
# SLIDE 12 — RESULTS / DEMO
# ============================================================================
s = add_slide(); bg(s)
header(s, "Results & Demonstration", "IT WORKS", 12)
# mock the app UI
card(s, Inches(0.8), Inches(2.1), Inches(5.4), Inches(4.3), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.5))
txt(s, Inches(0.8), Inches(2.45), Inches(5.4), Inches(0.5),
    [{"text": "UniVerify V2.0", "size": 24, "color": ACCENT_BLUE, "bold": True, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER)
txt(s, Inches(0.8), Inches(3.0), Inches(5.4), Inches(0.4),
    [{"text": "Decentralized Certificate Verification", "size": 12, "color": TEXT_GREY, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER)
card(s, Inches(1.2), Inches(3.55), Inches(4.6), Inches(0.55), fill=RGBColor(0x16,0x1B,0x22), border=LINE_SUBTLE)
txt(s, Inches(1.4), Inches(3.62), Inches(4.3), Inches(0.4),
    [{"text": "2021CSE101", "size": 13, "color": TEXT_WHITE, "font": FONT_MONO}], anchor=MSO_ANCHOR.MIDDLE)
card(s, Inches(1.2), Inches(4.25), Inches(4.6), Inches(0.55), fill=ACCENT_GREEN2, border=None)
txt(s, Inches(1.2), Inches(4.32), Inches(4.6), Inches(0.4),
    [{"text": "Verify on Blockchain", "size": 13, "color": TEXT_WHITE, "bold": True, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# result box
card(s, Inches(1.2), Inches(5.0), Inches(4.6), Inches(1.15), fill=RGBColor(0x10,0x26,0x16), border=ACCENT_GREEN, bw=Pt(1.25))
txt(s, Inches(1.4), Inches(5.12), Inches(4.3), Inches(1.0),
    [{"text": "✅ AUTHENTIC RECORD FOUND", "size": 12.5, "color": ACCENT_GREEN, "bold": True, "space_after": 4},
     {"text": "Student: Aarav Sharma", "size": 11, "color": TEXT_WHITE},
     {"text": "Course: B.Tech CSE", "size": 11, "color": TEXT_WHITE},
     {"text": "Hash: Qm7x...a91f", "size": 11, "color": TEXT_GREY, "font": FONT_MONO}])
# outcomes
card(s, Inches(6.6), Inches(2.1), Inches(6.0), Inches(4.3), fill=BG_PANEL, border=LINE_SUBTLE)
txt(s, Inches(6.9), Inches(2.35), Inches(5.5), Inches(0.4),
    [{"text": "Achieved Outcomes", "size": 17, "color": ACCENT_BLUE, "bold": True}])
bullet_block(s, Inches(6.9), Inches(2.95), Inches(5.5), Inches(3.4),
    ["Successfully deployed contract on Sepolia Testnet.",
     "Issued sample certificates from the university wallet.",
     "Verified records instantly by Register Number.",
     "Confirmed unauthorized wallets are blocked from issuing.",
     "Validated tamper-resistance \u2014 records cannot be edited.",
     "Clean, responsive portal with success / error feedback."],
    size=14, gap=14, color=TEXT_WHITE, marker="✓", marker_color=ACCENT_GREEN)

# ============================================================================
# SLIDE 13 — ADVANTAGES & APPLICATIONS
# ============================================================================
s = add_slide(); bg(s)
header(s, "Advantages & Applications", "WHY IT MATTERS", 13)
txt(s, Inches(0.7), Inches(1.9), Inches(6), Inches(0.4),
    [{"text": "Key Advantages", "size": 17, "color": ACCENT_GREEN, "bold": True}])
bullet_block(s, Inches(0.7), Inches(2.45), Inches(5.6), Inches(4.0),
    ["Tamper-proof & permanent records",
     "Instant, self-service verification",
     "No central authority or middlemen",
     "Globally accessible & transparent",
     "Low cost on test networks",
     "Eliminates certificate forgery"],
    size=15, gap=15, color=TEXT_WHITE, marker="✓", marker_color=ACCENT_GREEN)
txt(s, Inches(6.9), Inches(1.9), Inches(6), Inches(0.4),
    [{"text": "Real-World Applications", "size": 17, "color": ACCENT_BLUE, "bold": True}])
apps = [
    ("Universities & Colleges", "Issue verifiable digital degrees"),
    ("Recruiters & HR", "Instantly validate candidate credentials"),
    ("Online Course Platforms", "Tamper-proof completion certificates"),
    ("Professional Bodies", "License & membership verification"),
]
ay = Inches(2.45)
for t, d in apps:
    card(s, Inches(6.9), ay, Inches(5.7), Inches(0.85), fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, Inches(6.9), ay, Inches(0.08), Inches(0.85), ACCENT_BLUE)
    txt(s, Inches(7.15), ay + Inches(0.12), Inches(5.4), Inches(0.35),
        [{"text": t, "size": 14, "color": TEXT_WHITE, "bold": True}])
    txt(s, Inches(7.15), ay + Inches(0.46), Inches(5.4), Inches(0.3),
        [{"text": d, "size": 11.5, "color": TEXT_GREY}])
    ay += Inches(1.0)

# ============================================================================
# SLIDE 14 — LIMITATIONS & FUTURE SCOPE
# ============================================================================
s = add_slide(); bg(s)
header(s, "Limitations & Future Scope", "THE ROAD AHEAD", 14)
card(s, Inches(0.7), Inches(1.95), Inches(5.85), Inches(4.7), fill=BG_PANEL, border=ORANGE, bw=Pt(1.25))
txt(s, Inches(0.95), Inches(2.15), Inches(5.4), Inches(0.4),
    [{"text": "Current Limitations", "size": 17, "color": ORANGE, "bold": True}])
bullet_block(s, Inches(1.0), Inches(2.75), Inches(5.3), Inches(3.8),
    ["Deployed on a testnet (not mainnet) for now.",
     "Document stored off-chain via IPFS hash only.",
     "Single university-admin (no multi-institution roles).",
     "Requires MetaMask for issuing certificates.",
     "Gas costs apply on a real production network."],
    size=14, gap=14, color=TEXT_GREY, marker="•", marker_color=ORANGE)
card(s, Inches(6.8), Inches(1.95), Inches(5.85), Inches(4.7), fill=BG_PANEL, border=ACCENT_GREEN, bw=Pt(1.25))
txt(s, Inches(7.05), Inches(2.15), Inches(5.4), Inches(0.4),
    [{"text": "Future Enhancements", "size": 17, "color": ACCENT_GREEN, "bold": True}])
bullet_block(s, Inches(7.1), Inches(2.75), Inches(5.3), Inches(3.8),
    ["Multi-institution role-based access control.",
     "Full IPFS / decentralized document storage.",
     "QR-code based instant mobile verification.",
     "NFT-based degrees (soulbound tokens).",
     "Deployment on Ethereum / Polygon mainnet.",
     "Bulk issuance & analytics dashboard."],
    size=14, gap=13, color=TEXT_WHITE, marker="▸", marker_color=ACCENT_GREEN)

# ============================================================================
# SLIDE 15 — CONCLUSION
# ============================================================================
s = add_slide(); bg(s)
header(s, "Conclusion", "WRAPPING UP", 15)
card(s, Inches(0.9), Inches(2.1), Inches(11.5), Inches(2.7), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.25))
txt(s, Inches(1.3), Inches(2.45), Inches(10.7), Inches(2.1),
    [
     {"text": "UniVerify V2.0 demonstrates how blockchain technology can solve a real-world problem \u2014 academic credential forgery \u2014 by making records immutable, transparent and instantly verifiable.",
      "size": 18, "color": TEXT_WHITE, "bold": True, "line_spacing": 1.25, "space_after": 12},
     {"text": "By combining Solidity smart contracts, the Ethereum Sepolia testnet and an Ethers.js-powered Web3 frontend, the system removes the need for slow, centralized verification while guaranteeing data integrity through cryptography.",
      "size": 14, "color": TEXT_GREY, "line_spacing": 1.3},
    ])
# 3 takeaways
takes = [("Immutable", "Records can never be forged or altered", ACCENT_GREEN),
         ("Instant", "O(1) verification in seconds", ACCENT_BLUE),
         ("Trustless", "No central authority required", PURPLE)]
tx = Inches(0.9)
for t, d, col in takes:
    card(s, tx, Inches(5.05), Inches(3.7), Inches(1.5), fill=BG_PANEL_2, border=col, bw=Pt(1.25))
    txt(s, tx + Inches(0.3), Inches(5.25), Inches(3.2), Inches(0.5),
        [{"text": t, "size": 20, "color": col, "bold": True, "font": FONT_TITLE}])
    txt(s, tx + Inches(0.3), Inches(5.8), Inches(3.2), Inches(0.6),
        [{"text": d, "size": 12.5, "color": TEXT_GREY, "line_spacing": 1.1}])
    tx += Inches(3.92)

# ============================================================================
# SLIDE 16 — REFERENCES + THANK YOU
# ============================================================================
s = add_slide(); bg(s)
header(s, "References & Acknowledgement", "RESOURCES", 16)
refs = [
    "Solidity Documentation — docs.soliditylang.org",
    "Ethers.js v5 Documentation — docs.ethers.org/v5",
    "Ethereum Sepolia Testnet — ethereum.org/developers",
    "MetaMask Developer Docs — docs.metamask.io",
    "Remix IDE — remix.ethereum.org",
    "Project Repository — github.com/shaikmofficial-ai/Univerify",
]
card(s, Inches(0.7), Inches(1.95), Inches(11.9), Inches(3.0), fill=BG_PANEL, border=LINE_SUBTLE)
bullet_block(s, Inches(1.05), Inches(2.25), Inches(11.2), Inches(2.5), refs,
             size=14, gap=11, color=TEXT_GREY, marker="▸", marker_color=ACCENT_BLUE)
# thank you banner
card(s, Inches(0.7), Inches(5.15), Inches(11.9), Inches(1.5), fill=BG_PANEL_2, border=ACCENT_BLUE, bw=Pt(1.5))
txt(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(0.7),
    [{"text": "Thank You!", "size": 34, "color": ACCENT_BLUE, "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_TITLE}],
    align=PP_ALIGN.CENTER)
txt(s, Inches(0.7), Inches(6.1), Inches(11.9), Inches(0.4),
    [{"text": "Questions & Discussion  \u2022  UniVerify V2.0  \u2014  Decentralized Academic Credential Verification",
      "size": 13, "color": TEXT_GREY, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER)

# ----------------------------------------------------------------------------
prs.save("/projects/sandbox/Univerify/UniVerify_Final_Presentation.pptx")
print("Saved presentation with", len(prs.slides._sldIdLst), "slides.")
