#!/usr/bin/env python3
"""
UniVerify V2.0 - Final Project Presentation Generator (Academic Format)
22 slides: Title, Abstract, Introduction, Literature Survey, Existing System,
Existing Architecture, Problem Statement, Proposed System, Proposed Architecture,
Module Diagram, 7x UML Diagrams, Algorithm, Output, Comparative Graph,
Future Enhancement & Conclusion, References.
Footer: project + presenter + register no (left), slide NN/total (right).
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- THEME
BG_DARK      = RGBColor(0x0D, 0x11, 0x17)
BG_PANEL     = RGBColor(0x16, 0x1B, 0x22)
BG_PANEL_2   = RGBColor(0x1C, 0x22, 0x2B)
ACCENT_BLUE  = RGBColor(0x58, 0xA6, 0xFF)
ACCENT_GREEN = RGBColor(0x3F, 0xB9, 0x50)
TEXT_WHITE   = RGBColor(0xF0, 0xF6, 0xFC)
TEXT_GREY    = RGBColor(0x8B, 0x94, 0x9E)
BORDER_BLUE  = RGBColor(0x30, 0x4A, 0x6E)
LINE_SUBTLE  = RGBColor(0x30, 0x36, 0x3D)
PURPLE       = RGBColor(0xBC, 0x8C, 0xFF)
ORANGE       = RGBColor(0xF7, 0x8A, 0x3B)
RED          = RGBColor(0xF8, 0x51, 0x49)

FONT_TITLE = "Segoe UI Semibold"
FONT_HEAD  = "Segoe UI"
FONT_BODY  = "Segoe UI"
FONT_MONO  = "Consolas"

# Identity / footer
PROJECT   = "UniVerify V2.0"
PRESENTER = "Shaik Abdulla M"
REGNO     = "231061101493"
GUIDE     = "Dr. [Guide Name]"          # update when provided
TEAM = [("Shaik Abdulla M", "231061101493", "CSE"),
        ("Narendra Reddy",  "231061101489", "CSE")]
TOTAL = 22

SW = Inches(13.333)
SH = Inches(7.5)

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
BLANK = prs.slide_layouts[6]


def E(v):
    return int(v)


def add_slide():
    return prs.slides.add_slide(BLANK)


def bg(slide, color=BG_DARK):
    f = slide.background.fill
    f.solid(); f.fore_color.rgb = color


def rect(slide, x, y, w, h, color, line_color=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, E(x), E(y), E(w), E(h))
    if color is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line_color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line_color; sp.line.width = line_w or Pt(1)
    sp.shadow.inherit = False
    return sp


def txt(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ln.get("align", align)
        if ln.get("space_after") is not None: p.space_after = Pt(ln["space_after"])
        if ln.get("space_before") is not None: p.space_before = Pt(ln["space_before"])
        if ln.get("line_spacing"): p.line_spacing = ln["line_spacing"]
        run = p.add_run()
        run.text = ln["text"]
        f = run.font
        f.size = Pt(ln.get("size", 18)); f.bold = ln.get("bold", False)
        f.italic = ln.get("italic", False); f.name = ln.get("font", FONT_BODY)
        f.color.rgb = ln.get("color", TEXT_WHITE)
    return tb


def card(slide, x, y, w, h, fill=BG_PANEL, border=LINE_SUBTLE, bw=Pt(1), round=True):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if round else MSO_SHAPE.RECTANGLE
    sp = rect(slide, x, y, w, h, fill, border, bw, shape=shape)
    if round:
        try: sp.adjustments[0] = 0.05
        except Exception: pass
    return sp


def header(slide, title, kicker, num):
    rect(slide, 0, 0, SW, Inches(0.12), ACCENT_BLUE)
    rect(slide, Inches(0.55), Inches(0.5), Inches(0.13), Inches(0.92), ACCENT_BLUE)
    txt(slide, Inches(0.85), Inches(0.47), Inches(10), Inches(0.33),
        [{"text": kicker.upper(), "size": 12, "color": ACCENT_BLUE, "bold": True, "font": FONT_HEAD}])
    txt(slide, Inches(0.83), Inches(0.74), Inches(11.8), Inches(0.75),
        [{"text": title, "size": 28, "color": TEXT_WHITE, "bold": True, "font": FONT_TITLE}])
    footer(slide, num)


def footer(slide, num):
    rect(slide, 0, SH - Inches(0.05), SW, Inches(0.05), LINE_SUBTLE)
    # left: project | name - regno
    tb = slide.shapes.add_textbox(Inches(0.55), SH - Inches(0.42), Inches(10), Inches(0.3))
    p = tb.text_frame.paragraphs[0]
    p.text = ""
    r1 = p.add_run(); r1.text = PROJECT + "  "
    r1.font.size = Pt(9.5); r1.font.bold = True; r1.font.color.rgb = ACCENT_BLUE; r1.font.name = FONT_BODY
    r2 = p.add_run(); r2.text = f"|  {PRESENTER} - {REGNO}"
    r2.font.size = Pt(9.5); r2.font.color.rgb = TEXT_GREY; r2.font.name = FONT_BODY
    # right: NN / TOTAL
    txt(slide, SW - Inches(1.7), SH - Inches(0.42), Inches(1.15), Inches(0.3),
        [{"text": f"{num:02d} / {TOTAL}", "size": 11, "color": ACCENT_BLUE, "bold": True,
          "align": PP_ALIGN.RIGHT}], align=PP_ALIGN.RIGHT)


def bullets(slide, x, y, w, h, items, size=15, gap=10, color=TEXT_WHITE,
            marker="\u25B8", marker_color=ACCENT_BLUE):
    tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
    tf = tb.text_frame; tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = 1.05
        r1 = p.add_run(); r1.text = marker + "  "
        r1.font.size = Pt(size); r1.font.bold = True; r1.font.color.rgb = marker_color; r1.font.name = FONT_BODY
        if "::" in it:
            lead, rest = it.split("::", 1)
            rb = p.add_run(); rb.text = lead + "  "
            rb.font.size = Pt(size); rb.font.bold = True; rb.font.color.rgb = TEXT_WHITE; rb.font.name = FONT_BODY
            rr = p.add_run(); rr.text = rest.strip()
            rr.font.size = Pt(size); rr.font.color.rgb = color; rr.font.name = FONT_BODY
        else:
            r2 = p.add_run(); r2.text = it
            r2.font.size = Pt(size); r2.font.color.rgb = color; r2.font.name = FONT_BODY
    return tb


# ---------------------------------------------------------------- DIAGRAM HELPERS
def connector(slide, x1, y1, x2, y2, color=ACCENT_BLUE, w=1.5, dashed=False, arrow=True, both=False):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2))
    cn.line.color.rgb = color; cn.line.width = Pt(w)
    try: cn.shadow.inherit = False
    except Exception: pass
    ln = cn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn('a:prstDash'), {'val': 'dash'}))
    if both:
        ln.append(ln.makeelement(qn('a:headEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
    if arrow:
        ln.append(ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
    return cn


def label(slide, x, y, w, txt_, size=11, color=TEXT_WHITE, bold=False, align=PP_ALIGN.CENTER,
          font=FONT_BODY, italic=False):
    return txt(slide, x, y, w, Inches(0.3),
               [{"text": txt_, "size": size, "color": color, "bold": bold, "align": align,
                 "font": font, "italic": italic}], align=align)


def actor(slide, cx, top, name, color=ACCENT_BLUE):
    hd = Inches(0.32)
    rect(slide, cx - hd / 2, top, hd, hd, BG_PANEL_2, color, Pt(2), shape=MSO_SHAPE.OVAL)
    body_t = top + hd
    body_b = body_t + Inches(0.5)
    connector(slide, cx, body_t, cx, body_b, color, 2, arrow=False)
    arm_y = body_t + Inches(0.14)
    connector(slide, cx - Inches(0.3), arm_y, cx + Inches(0.3), arm_y, color, 2, arrow=False)
    connector(slide, cx, body_b, cx - Inches(0.24), body_b + Inches(0.34), color, 2, arrow=False)
    connector(slide, cx, body_b, cx + Inches(0.24), body_b + Inches(0.34), color, 2, arrow=False)
    label(slide, cx - Inches(1.1), body_b + Inches(0.4), Inches(2.2), name, 12, TEXT_WHITE, True)


def usecase(slide, x, y, w, h, name, color=ACCENT_GREEN):
    rect(slide, x, y, w, h, BG_PANEL, color, Pt(1.25), shape=MSO_SHAPE.OVAL)
    txt(slide, x + Inches(0.08), y, w - Inches(0.16), h,
        [{"text": name, "size": 10.5, "color": TEXT_WHITE, "align": PP_ALIGN.CENTER, "line_spacing": 1.0}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def uml_class(slide, x, y, w, name, attrs, methods, color=ACCENT_BLUE):
    th = 0.4
    ah = 0.25 * max(1, len(attrs)) + 0.14
    mh = 0.25 * max(1, len(methods)) + 0.14
    rect(slide, x, y, w, Inches(th), color)
    txt(slide, x, y + Inches(0.05), w, Inches(0.32),
        [{"text": name, "size": 13, "color": BG_DARK, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER)
    ay = y + Inches(th)
    rect(slide, x, ay, w, Inches(ah), BG_PANEL, color, Pt(1.25))
    txt(slide, x + Inches(0.14), ay + Inches(0.07), w - Inches(0.24), Inches(ah),
        [{"text": a, "size": 10.5, "color": TEXT_WHITE, "font": FONT_MONO, "space_after": 3} for a in attrs])
    my = ay + Inches(ah)
    rect(slide, x, my, w, Inches(mh), BG_PANEL, color, Pt(1.25))
    txt(slide, x + Inches(0.14), my + Inches(0.07), w - Inches(0.24), Inches(mh),
        [{"text": m, "size": 10.5, "color": ACCENT_GREEN, "font": FONT_MONO, "space_after": 3} for m in methods])
    return Inches(th) + Inches(ah) + Inches(mh)


def node_box(slide, x, y, w, h, title, sub="", color=ACCENT_BLUE, fill=BG_PANEL):
    card(slide, x, y, w, h, fill=fill, border=color, bw=Pt(1.5))
    if sub:
        txt(slide, x, y + Inches(0.13), w, Inches(0.4),
            [{"text": title, "size": 13, "color": color, "bold": True, "align": PP_ALIGN.CENTER}],
            align=PP_ALIGN.CENTER)
        txt(slide, x + Inches(0.08), y + Inches(0.5), w - Inches(0.16), h - Inches(0.55),
            [{"text": sub, "size": 10.5, "color": TEXT_GREY, "align": PP_ALIGN.CENTER, "line_spacing": 1.05}],
            align=PP_ALIGN.CENTER)
    else:
        txt(slide, x + Inches(0.06), y, w - Inches(0.12), h,
            [{"text": title, "size": 12.5, "color": TEXT_WHITE, "bold": True, "align": PP_ALIGN.CENTER,
              "line_spacing": 1.05}], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def code_block(slide, x, y, w, h, code_lines, title="solidity"):
    card(slide, x, y, w, h, fill=RGBColor(0x0A, 0x0E, 0x14), border=BORDER_BLUE, bw=Pt(1.25))
    for i, c in enumerate([RED, ORANGE, ACCENT_GREEN]):
        rect(slide, x + Inches(0.25) + Inches(0.28) * i, y + Inches(0.2), Inches(0.13), Inches(0.13),
             c, shape=MSO_SHAPE.OVAL)
    txt(slide, x + Inches(1.3), y + Inches(0.12), w - Inches(1.5), Inches(0.3),
        [{"text": title, "size": 11, "color": TEXT_GREY, "font": FONT_MONO, "align": PP_ALIGN.RIGHT}],
        align=PP_ALIGN.RIGHT)
    tb = slide.shapes.add_textbox(E(x + Inches(0.3)), E(y + Inches(0.5)), E(w - Inches(0.6)), E(h - Inches(0.65)))
    tf = tb.text_frame; tf.word_wrap = True
    for i, (line, col) in enumerate(code_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(2); p.line_spacing = 1.0
        r = p.add_run(); r.text = line if line else " "
        r.font.size = Pt(11); r.font.name = FONT_MONO; r.font.color.rgb = col


# ================================================================ SLIDE 1: TITLE
s = add_slide(); bg(s)
rect(s, 0, 0, Inches(0.18), SH, ACCENT_BLUE)
for i in range(7):
    bx = Inches(7.2 + i * 0.78)
    rect(s, bx, Inches(0.5), Inches(0.5), Inches(0.5), BG_PANEL_2, BORDER_BLUE, Pt(1),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    if i < 6:
        rect(s, bx + Inches(0.5), Inches(0.72), Inches(0.28), Inches(0.06), ACCENT_BLUE)
txt(s, Inches(0.9), Inches(1.45), Inches(11), Inches(0.4),
    [{"text": "BLOCKCHAIN  \u2022  WEB3  \u2022  FINAL YEAR PROJECT", "size": 14, "color": ACCENT_BLUE,
      "bold": True, "font": FONT_HEAD}])
txt(s, Inches(0.85), Inches(1.9), Inches(11.8), Inches(1.4),
    [{"text": "UniVerify V2.0", "size": 64, "color": TEXT_WHITE, "bold": True, "font": FONT_TITLE}])
txt(s, Inches(0.9), Inches(3.15), Inches(11.4), Inches(0.6),
    [{"text": "Decentralized Academic Credential Verification Portal", "size": 22,
      "color": ACCENT_GREEN, "bold": True, "font": FONT_HEAD}])
txt(s, Inches(0.9), Inches(3.82), Inches(10.8), Inches(0.5),
    [{"text": "Eliminating certificate forgery using immutable Ethereum smart contracts.",
      "size": 14, "color": TEXT_GREY, "italic": True}])
# Work done by
card(s, Inches(0.9), Inches(4.55), Inches(6.6), Inches(2.05), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.25))
txt(s, Inches(1.15), Inches(4.72), Inches(6), Inches(0.35),
    [{"text": "WORK DONE BY", "size": 12, "color": ACCENT_BLUE, "bold": True}])
ty = Inches(5.12)
for nm, rg, br in TEAM:
    txt(s, Inches(1.15), ty, Inches(6.2), Inches(0.4),
        [{"text": f"{nm}", "size": 15, "color": TEXT_WHITE, "bold": True}])
    txt(s, Inches(4.6), ty + Inches(0.02), Inches(2.8), Inches(0.4),
        [{"text": f"{rg}  \u2022  {br}", "size": 13, "color": TEXT_GREY, "align": PP_ALIGN.RIGHT}],
        align=PP_ALIGN.RIGHT)
    ty += Inches(0.52)
# Guide
card(s, Inches(7.75), Inches(4.55), Inches(4.65), Inches(2.05), fill=BG_PANEL, border=ACCENT_GREEN, bw=Pt(1.25))
txt(s, Inches(8.0), Inches(4.72), Inches(4), Inches(0.35),
    [{"text": "PROJECT GUIDE", "size": 12, "color": ACCENT_GREEN, "bold": True}])
txt(s, Inches(8.0), Inches(5.18), Inches(4.2), Inches(0.5),
    [{"text": GUIDE, "size": 17, "color": TEXT_WHITE, "bold": True}])
txt(s, Inches(8.0), Inches(5.7), Inches(4.2), Inches(0.8),
    [{"text": "Department of Computer Science & Engineering", "size": 12, "color": TEXT_GREY,
      "line_spacing": 1.1}])
footer(s, 1)

# ================================================================ SLIDE 2: ABSTRACT
s = add_slide(); bg(s)
header(s, "Abstract", "Project at a Glance", 2)
card(s, Inches(0.7), Inches(1.75), Inches(7.5), Inches(4.9), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.25))
txt(s, Inches(1.05), Inches(2.05), Inches(6.9), Inches(4.4),
    [{"text": "UniVerify V2.0 is a Web3 application engineered to eliminate academic credential forgery.",
      "size": 17, "color": TEXT_WHITE, "bold": True, "line_spacing": 1.15, "space_after": 10},
     {"text": "Educational institutions anchor degree metadata to the Ethereum Sepolia Testnet using cryptographic hashes. Once a transaction is mined, the record becomes permanent and tamper-proof.",
      "size": 13.5, "color": TEXT_GREY, "line_spacing": 1.2, "space_after": 9},
     {"text": "Verification is instant: a recruiter enters a student's Register Number and the system queries the blockchain in O(1) time via Solidity mappings, returning a cryptographically-backed authentic record \u2014 with no central authority required.",
      "size": 13.5, "color": TEXT_GREY, "line_spacing": 1.2, "space_after": 9},
     {"text": "The result is a trustless, transparent and globally verifiable system for academic certificates.",
      "size": 13.5, "color": ACCENT_GREEN, "italic": True, "line_spacing": 1.2}])
stats = [("100%", "Tamper-proof records", ACCENT_GREEN), ("O(1)", "Lookup complexity", ACCENT_BLUE),
         ("0", "Central authorities", PURPLE), ("24/7", "Global verification", ORANGE)]
sy = Inches(1.75)
for v, lab, col in stats:
    card(s, Inches(8.4), sy, Inches(4.2), Inches(1.07), fill=BG_PANEL_2, border=col, bw=Pt(1.25))
    txt(s, Inches(8.65), sy + Inches(0.12), Inches(1.7), Inches(0.85),
        [{"text": v, "size": 27, "color": col, "bold": True, "font": FONT_TITLE}], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(10.2), sy + Inches(0.12), Inches(2.3), Inches(0.85),
        [{"text": lab, "size": 13, "color": TEXT_WHITE, "bold": True}], anchor=MSO_ANCHOR.MIDDLE)
    sy += Inches(1.22)

# ================================================================ SLIDE 3: INTRODUCTION
s = add_slide(); bg(s)
header(s, "Introduction", "Background & Motivation", 3)
txt(s, Inches(0.7), Inches(1.75), Inches(11.9), Inches(1.3),
    [{"text": "Academic certificates are the primary proof of a person's qualifications, yet they remain one of the easiest documents to forge. As education and hiring go global and digital, the need for a fast, trustworthy and universally verifiable credential system has never been greater.",
      "size": 15, "color": TEXT_GREY, "line_spacing": 1.3}])
intro = [
    ("Blockchain Foundation", "A blockchain is a distributed, append-only ledger where records are cryptographically chained, making stored data immutable and transparent.", ACCENT_BLUE),
    ("Smart Contracts", "Self-executing programs on Ethereum that enforce rules automatically \u2014 here, only the university wallet may issue a certificate.", PURPLE),
    ("Why UniVerify", "It moves credential records from vulnerable centralized databases onto a public, tamper-proof blockchain accessible to anyone.", ACCENT_GREEN),
    ("Scope of Work", "A working DApp with a Solidity contract on Sepolia Testnet and a Web3 frontend for issuing and verifying degrees.", ORANGE),
]
gx = Inches(0.7); gy = Inches(3.25); cw = Inches(5.85); ch = Inches(1.6)
for i, (t, d, col) in enumerate(intro):
    x = gx + (i % 2) * (cw + Inches(0.25)); y = gy + (i // 2) * (ch + Inches(0.22))
    card(s, x, y, cw, ch, fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, x, y, Inches(0.09), ch, col)
    txt(s, x + Inches(0.3), y + Inches(0.18), cw - Inches(0.5), Inches(0.4),
        [{"text": t, "size": 16, "color": TEXT_WHITE, "bold": True}])
    txt(s, x + Inches(0.3), y + Inches(0.62), cw - Inches(0.55), Inches(0.9),
        [{"text": d, "size": 12.5, "color": TEXT_GREY, "line_spacing": 1.15}])

# ================================================================ SLIDE 4: LITERATURE SURVEY
s = add_slide(); bg(s)
header(s, "Literature Survey", "Related Work", 4)
rows = [
    ("Author / Year", "Approach / Focus", "Limitation Addressed by UniVerify"),
    ("Gresch et al., 2018", "Blockchain-based diploma verification on Ethereum", "Complex setup; UniVerify offers a lightweight register-no lookup"),
    ("Cheng et al., 2018", "Hash of certificate stored on-chain (Blockcerts style)", "No simple public UI; UniVerify adds a self-service web portal"),
    ("Arenas & Fernandez, 2018", "University credential issuance via smart contracts", "Centralized issuer keys; UniVerify uses on-chain owner lock"),
    ("Saleh et al., 2020", "IPFS + blockchain for document integrity", "UniVerify links on-chain record to IPFS hash for the file"),
    ("Traditional Systems", "Paper certificates & centralized DB verification", "Forgeable & slow; UniVerify is immutable & instant"),
]
tbl_x, tbl_y = Inches(0.7), Inches(1.85)
tbl_w, tbl_h = Inches(11.9), Inches(4.6)
gtbl = s.shapes.add_table(len(rows), 3, E(tbl_x), E(tbl_y), E(tbl_w), E(tbl_h)).table
gtbl.columns[0].width = E(Inches(2.7)); gtbl.columns[1].width = E(Inches(4.2)); gtbl.columns[2].width = E(Inches(5.0))
for r in range(len(rows)):
    for c in range(3):
        cell = gtbl.cell(r, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = ACCENT_BLUE if r == 0 else (BG_PANEL if r % 2 else BG_PANEL_2)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.12); cell.margin_right = Inches(0.1)
        cell.margin_top = Inches(0.04); cell.margin_bottom = Inches(0.04)
        para = cell.text_frame.paragraphs[0]
        run = para.add_run(); run.text = rows[r][c]
        run.font.size = Pt(12 if r == 0 else 11)
        run.font.bold = (r == 0 or c == 0)
        run.font.color.rgb = BG_DARK if r == 0 else (TEXT_WHITE if c == 0 else TEXT_GREY)
        run.font.name = FONT_BODY

# ================================================================ SLIDE 5: EXISTING SYSTEM
s = add_slide(); bg(s)
header(s, "Existing System", "How Verification Works Today", 5)
txt(s, Inches(0.7), Inches(1.75), Inches(11.9), Inches(0.5),
    [{"text": "Most institutions still rely on paper certificates and centralized databases, verified manually on request.",
      "size": 15, "color": TEXT_GREY, "italic": True}])
ex = [
    ("Paper Certificates", "Physical documents with seals, stamps and signatures issued by the institution."),
    ("Centralized Database", "Records kept in a single university server / management system."),
    ("Manual Verification", "Employers email or call the university and wait for confirmation."),
    ("Document Scans", "PDFs / photocopies shared over email as 'proof' of a degree."),
]
gx = Inches(0.7); gy = Inches(2.5); cw = Inches(5.85); ch = Inches(1.85)
for i, (t, d) in enumerate(ex):
    x = gx + (i % 2) * (cw + Inches(0.25)); y = gy + (i // 2) * (ch + Inches(0.25))
    card(s, x, y, cw, ch, fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, x, y, cw, Inches(0.09), ORANGE)
    txt(s, x + Inches(0.3), y + Inches(0.28), cw - Inches(0.5), Inches(0.45),
        [{"text": t, "size": 17, "color": TEXT_WHITE, "bold": True}])
    txt(s, x + Inches(0.3), y + Inches(0.8), cw - Inches(0.55), Inches(0.9),
        [{"text": d, "size": 13, "color": TEXT_GREY, "line_spacing": 1.15}])

# ================================================================ SLIDE 6: EXISTING ARCHITECTURE
s = add_slide(); bg(s)
header(s, "Architecture of Existing System", "Centralized Model", 6)
ny = Inches(2.65); nh = Inches(1.5); nw = Inches(2.55)
node_box(s, Inches(0.7), ny, nw, nh, "Student /\nApplicant", "Submits / shares paper certificate", ORANGE)
connector(s, Inches(3.3), ny + Inches(0.6), Inches(3.85), ny + Inches(0.6), ORANGE)
node_box(s, Inches(3.95), ny, nw, nh, "University\nAdmin Office", "Issues & records certificate", ACCENT_BLUE)
connector(s, Inches(6.55), ny + Inches(0.6), Inches(7.1), ny + Inches(0.6), ACCENT_BLUE)
node_box(s, Inches(7.2), ny, nw, nh, "Centralized\nDatabase", "Single server stores records", PURPLE)
# verifier path
connector(s, Inches(9.8), ny + Inches(0.45), Inches(10.6), ny + Inches(0.45), RED, both=True)
node_box(s, Inches(10.65), ny, Inches(2.0), nh, "Employer /\nVerifier", "Emails to confirm", RED)
# DB to verifier dashed (manual)
txt(s, Inches(0.7), Inches(4.6), Inches(11.9), Inches(0.4),
    [{"text": "Verification is manual and slow \u2014 the verifier must contact the university, which queries its single database.",
      "size": 13, "color": TEXT_GREY, "italic": True}])
# weak points
for i, w in enumerate(["Single point of failure", "No public transparency", "Forgeable documents", "Slow (days/weeks)"]):
    bx = Inches(0.7 + i * 3.0)
    card(s, bx, Inches(5.2), Inches(2.8), Inches(0.95), fill=BG_PANEL, border=RED, bw=Pt(1.25))
    txt(s, bx + Inches(0.15), Inches(5.2), Inches(2.5), Inches(0.95),
        [{"text": "\u26A0  " + w, "size": 12.5, "color": TEXT_WHITE, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ SLIDE 7: PROBLEM STATEMENT
s = add_slide(); bg(s)
header(s, "Drawbacks & Problem Statement", "The Challenge", 7)
card(s, Inches(0.7), Inches(1.8), Inches(11.9), Inches(1.15), fill=BG_PANEL_2, border=RED, bw=Pt(1.25))
txt(s, Inches(1.0), Inches(1.95), Inches(11.3), Inches(0.9),
    [{"text": "Problem Statement", "size": 13, "color": RED, "bold": True, "space_after": 3},
     {"text": "Existing credential verification depends on centralized, manual and forgeable processes \u2014 there is no fast, tamper-proof and globally trusted way to confirm whether an academic certificate is authentic.",
      "size": 14, "color": TEXT_WHITE, "line_spacing": 1.18}])
probs = [
    ("Rampant Forgery", "Certificates are easily edited, copied or fabricated.", RED),
    ("Slow Verification", "Manual checks take days or weeks per request.", ORANGE),
    ("Single Point of Failure", "A lost or hacked database makes records unverifiable.", PURPLE),
    ("No Global Standard", "Cross-border verification is inconsistent and costly.", ACCENT_BLUE),
]
gx = Inches(0.7); gy = Inches(3.2); cw = Inches(5.85); ch = Inches(1.6)
for i, (t, d, col) in enumerate(probs):
    x = gx + (i % 2) * (cw + Inches(0.25)); y = gy + (i // 2) * (ch + Inches(0.2))
    card(s, x, y, cw, ch, fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, x + Inches(0.28), y + Inches(0.32), Inches(0.55), Inches(0.55), col, shape=MSO_SHAPE.OVAL)
    txt(s, x + Inches(0.28), y + Inches(0.32), Inches(0.55), Inches(0.55),
        [{"text": "!", "size": 24, "color": BG_DARK, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(1.05), y + Inches(0.26), cw - Inches(1.3), Inches(0.5),
        [{"text": t, "size": 17, "color": TEXT_WHITE, "bold": True}])
    txt(s, x + Inches(1.05), y + Inches(0.72), cw - Inches(1.3), Inches(0.8),
        [{"text": d, "size": 12.5, "color": TEXT_GREY, "line_spacing": 1.15}])

# ================================================================ SLIDE 8: PROPOSED SYSTEM
s = add_slide(); bg(s)
header(s, "Proposed System", "The UniVerify Solution", 8)
txt(s, Inches(0.7), Inches(1.75), Inches(11.9), Inches(0.55),
    [{"text": "A decentralized portal that anchors credentials on the Ethereum blockchain for instant, trustless verification.",
      "size": 15, "color": ACCENT_GREEN, "italic": True}])
feats = [
    ("Immutable Records", "Credential metadata is stored on-chain and can never be altered or deleted.", ACCENT_BLUE),
    ("Instant Verification", "Anyone can verify a degree in seconds using only the Register Number.", ACCENT_GREEN),
    ("Owner-Locked Issuance", "Only the university wallet can issue certificates via on-chain access control.", PURPLE),
    ("Cryptographic Integrity", "An IPFS hash links the on-chain record to the actual certificate file.", ORANGE),
    ("No Middlemen", "Trustless verification removes the need to contact the university.", ACCENT_BLUE),
    ("Web3 Portal", "A clean MetaMask-connected DApp for issuing and verifying degrees.", ACCENT_GREEN),
]
gx = Inches(0.7); gy = Inches(2.45); cw = Inches(3.85); ch = Inches(1.95)
for i, (t, d, col) in enumerate(feats):
    x = gx + (i % 3) * (cw + Inches(0.18)); y = gy + (i // 3) * (ch + Inches(0.2))
    card(s, x, y, cw, ch, fill=BG_PANEL, border=LINE_SUBTLE)
    rect(s, x + Inches(0.26), y + Inches(0.24), Inches(0.65), Inches(0.65), BG_PANEL_2, col, Pt(1.5),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, x + Inches(0.26), y + Inches(0.24), Inches(0.65), Inches(0.65),
        [{"text": str(i + 1), "size": 24, "color": col, "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_TITLE}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(1.05), y + Inches(0.3), cw - Inches(1.2), Inches(0.55),
        [{"text": t, "size": 14.5, "color": TEXT_WHITE, "bold": True, "line_spacing": 1.0}])
    txt(s, x + Inches(0.26), y + Inches(1.05), cw - Inches(0.5), Inches(0.85),
        [{"text": d, "size": 11.5, "color": TEXT_GREY, "line_spacing": 1.12}])

# ================================================================ SLIDE 9: PROPOSED ARCHITECTURE
s = add_slide(); bg(s)
header(s, "Architecture of Proposed System", "End-to-End Decentralized Flow", 9)
ny = Inches(2.3); nh = Inches(1.6); nw = Inches(2.45)
node_box(s, Inches(0.6), ny, nw, nh, "User /\nRecruiter", "Enters Register No in web portal", ACCENT_BLUE)
connector(s, Inches(3.1), ny + Inches(0.55), Inches(3.7), ny + Inches(0.55))
node_box(s, Inches(3.75), ny, nw, nh, "Frontend DApp", "HTML / CSS / JS UI", ACCENT_GREEN)
connector(s, Inches(6.25), ny + Inches(0.55), Inches(6.85), ny + Inches(0.55), ACCENT_GREEN)
node_box(s, Inches(6.9), ny, nw, nh, "Ethers.js +\nMetaMask", "Web3 bridge signs & sends calls", PURPLE)
connector(s, Inches(9.4), ny + Inches(0.55), Inches(10.0), ny + Inches(0.55), PURPLE)
node_box(s, Inches(10.05), ny, Inches(2.65), nh, "Smart\nContract", "AcademicCertificates on Sepolia", ORANGE)
txt(s, Inches(0.6), Inches(4.4), Inches(11), Inches(0.35),
    [{"text": "Ethereum Sepolia Testnet  \u2014  Immutable Distributed Ledger", "size": 13, "color": ACCENT_BLUE, "bold": True}])
bx = Inches(0.6)
for i in range(8):
    rect(s, bx, Inches(4.8), Inches(1.35), Inches(0.9), BG_PANEL_2, BORDER_BLUE, Pt(1.25),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, bx, Inches(4.92), Inches(1.35), Inches(0.3),
        [{"text": f"Block #{i + 1}", "size": 9.5, "color": ACCENT_GREEN, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER)
    txt(s, bx, Inches(5.22), Inches(1.35), Inches(0.4),
        [{"text": "0x" + "%04x" % (i * 4231 % 65535), "size": 9, "color": TEXT_GREY, "align": PP_ALIGN.CENTER, "font": FONT_MONO}],
        align=PP_ALIGN.CENTER)
    if i < 7:
        rect(s, bx + Inches(1.35), Inches(5.2), Inches(0.18), Inches(0.08), ACCENT_BLUE)
    bx += Inches(1.53)
txt(s, Inches(0.6), Inches(5.9), Inches(11.9), Inches(0.4),
    [{"text": "Each block is cryptographically linked to the previous one \u2014 altering any record breaks the entire chain.",
      "size": 12, "color": TEXT_GREY, "italic": True}])

# ================================================================ SLIDE 10: MODULE DIAGRAM
s = add_slide(); bg(s)
header(s, "Module Diagram", "System Modules", 10)
cx, cy = Inches(6.66), Inches(4.15)
rect(s, cx - Inches(1.4), cy - Inches(0.6), Inches(2.8), Inches(1.2), BG_PANEL_2, ACCENT_BLUE, Pt(2),
     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
txt(s, cx - Inches(1.4), cy - Inches(0.6), Inches(2.8), Inches(1.2),
    [{"text": "UniVerify\nCore System", "size": 16, "color": ACCENT_BLUE, "bold": True, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
mods = [
    ("1. Wallet Connection", "Connects MetaMask & detects account", ACCENT_GREEN, Inches(0.7), Inches(1.95)),
    ("2. Certificate Issuance", "Admin writes record to blockchain", PURPLE, Inches(9.0), Inches(1.95)),
    ("3. Verification", "Reads record by Register Number", ORANGE, Inches(0.7), Inches(5.35)),
    ("4. Blockchain Interaction", "Ethers.js calls to smart contract", ACCENT_BLUE, Inches(9.0), Inches(5.35)),
]
for t, d, col, mx, my in mods:
    card(s, mx, my, Inches(3.65), Inches(1.25), fill=BG_PANEL, border=col, bw=Pt(1.5))
    txt(s, mx + Inches(0.18), my + Inches(0.16), Inches(3.3), Inches(0.4),
        [{"text": t, "size": 14, "color": col, "bold": True}])
    txt(s, mx + Inches(0.18), my + Inches(0.58), Inches(3.3), Inches(0.6),
        [{"text": d, "size": 11.5, "color": TEXT_GREY, "line_spacing": 1.1}])
    # connect to center
    mx_c = mx + Inches(1.82); my_c = my + Inches(0.6)
    connector(s, mx_c, my_c, cx, cy, col, 1.5, arrow=False)
# 5th module (result display) below center
card(s, cx - Inches(1.82), Inches(5.95), Inches(3.65), Inches(1.0), fill=BG_PANEL, border=RED, bw=Pt(1.5))
txt(s, cx - Inches(1.64), Inches(6.08), Inches(3.3), Inches(0.4),
    [{"text": "5. Result / UI Display", "size": 14, "color": RED, "bold": True}])
txt(s, cx - Inches(1.64), Inches(6.46), Inches(3.3), Inches(0.4),
    [{"text": "Renders success / error states to user", "size": 11, "color": TEXT_GREY}])
connector(s, cx, Inches(5.95), cx, cy + Inches(0.6), RED, 1.5, arrow=False)

# ================================================================ SLIDE 11: UML - USE CASE
s = add_slide(); bg(s)
header(s, "UML \u2014 Use Case Diagram", "Diagram 1 of 7", 11)
# system boundary
rect(s, Inches(3.7), Inches(1.85), Inches(6.0), Inches(4.7), None, ACCENT_BLUE, Pt(1.75),
     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
txt(s, Inches(3.7), Inches(1.95), Inches(6.0), Inches(0.35),
    [{"text": "UniVerify System", "size": 13, "color": ACCENT_BLUE, "bold": True, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER)
# actors
actor(s, Inches(1.6), Inches(2.7), "University Admin", PURPLE)
actor(s, Inches(11.5), Inches(2.7), "Verifier /\nRecruiter", ACCENT_GREEN)
# use cases (left col = admin, right col = verifier, middle shared)
uc = [
    (Inches(4.0), Inches(2.45), "Connect\nWallet", ACCENT_BLUE),
    (Inches(4.0), Inches(3.5), "Issue\nCertificate", PURPLE),
    (Inches(4.0), Inches(4.55), "Store Record\nOn-Chain", PURPLE),
    (Inches(6.75), Inches(2.45), "Enter\nRegister No", ACCENT_GREEN),
    (Inches(6.75), Inches(3.5), "Verify\nCertificate", ACCENT_GREEN),
    (Inches(6.75), Inches(4.55), "View\nResult", ACCENT_BLUE),
    (Inches(5.35), Inches(5.55), "Query\nBlockchain", ORANGE),
]
uw, uh = Inches(2.0), Inches(0.85)
for x, y, name, col in uc:
    usecase(s, x, y, uw, uh, name, col)
# admin connections
connector(s, Inches(2.0), Inches(3.1), Inches(4.0), Inches(2.85), PURPLE, 1.25, arrow=False)
connector(s, Inches(2.0), Inches(3.3), Inches(4.0), Inches(3.9), PURPLE, 1.25, arrow=False)
connector(s, Inches(2.0), Inches(3.5), Inches(4.0), Inches(4.95), PURPLE, 1.25, arrow=False)
# verifier connections
connector(s, Inches(11.1), Inches(3.1), Inches(8.75), Inches(2.85), ACCENT_GREEN, 1.25, arrow=False)
connector(s, Inches(11.1), Inches(3.3), Inches(8.75), Inches(3.9), ACCENT_GREEN, 1.25, arrow=False)
connector(s, Inches(11.1), Inches(3.5), Inches(8.75), Inches(4.95), ACCENT_GREEN, 1.25, arrow=False)
# include arrows to query
connector(s, Inches(5.0), Inches(4.95), Inches(5.9), Inches(5.55), ORANGE, 1.25, dashed=True)
connector(s, Inches(7.5), Inches(4.4), Inches(6.6), Inches(5.55), ORANGE, 1.25, dashed=True)
label(s, Inches(8.0), Inches(5.05), Inches(2.2), "\u00abinclude\u00bb", 9, ORANGE, italic=True, align=PP_ALIGN.LEFT)

# ================================================================ SLIDE 12: UML - CLASS
s = add_slide(); bg(s)
header(s, "UML \u2014 Class Diagram", "Diagram 2 of 7", 12)
uml_class(s, Inches(0.7), Inches(2.0), Inches(3.6), "Certificate",
          ["- studentName: string", "- courseName: string", "- ipfsHash: string"], ["(struct type)"], PURPLE)
uml_class(s, Inches(5.1), Inches(1.95), Inches(4.0), "AcademicCertificates",
          ["- university: address", "- nextId: uint256", "- certificates: mapping"],
          ["+ issueDegree()", "+ getCertificate()"], ACCENT_BLUE)
uml_class(s, Inches(9.8), Inches(2.0), Inches(2.9), "Web3Frontend",
          ["- provider", "- contract", "- abi"], ["+ verify()", "+ connect()"], ACCENT_GREEN)
# relationships
connector(s, Inches(4.3), Inches(2.55), Inches(5.1), Inches(2.55), TEXT_GREY, 1.5, both=True)
label(s, Inches(4.0), Inches(2.15), Inches(1.5), "1   *", 11, TEXT_GREY)
label(s, Inches(4.05), Inches(2.62), Inches(1.4), "contains", 9.5, TEXT_GREY, italic=True)
connector(s, Inches(9.1), Inches(2.55), Inches(9.8), Inches(2.55), TEXT_GREY, 1.5)
label(s, Inches(8.95), Inches(2.62), Inches(1.5), "calls", 9.5, TEXT_GREY, italic=True)
txt(s, Inches(0.7), Inches(5.6), Inches(11.9), Inches(0.9),
    [{"text": "The AcademicCertificates contract owns a mapping of Certificate structs keyed by Register Number. "
              "The Web3Frontend class invokes the contract's public methods through Ethers.js.",
      "size": 13, "color": TEXT_GREY, "line_spacing": 1.2}])

# ================================================================ SLIDE 13: UML - SEQUENCE
s = add_slide(); bg(s)
header(s, "UML \u2014 Sequence Diagram", "Diagram 3 of 7", 13)
lifelines = [("User", Inches(1.6), ACCENT_BLUE), ("Frontend", Inches(4.0), ACCENT_GREEN),
             ("Ethers.js", Inches(6.66), PURPLE), ("MetaMask", Inches(9.0), ORANGE),
             ("SmartContract", Inches(11.5), RED)]
top = Inches(1.95); bot = Inches(6.5)
for nm, lx, col in lifelines:
    card(s, lx - Inches(0.85), top, Inches(1.7), Inches(0.5), fill=BG_PANEL, border=col, bw=Pt(1.5))
    txt(s, lx - Inches(0.85), top, Inches(1.7), Inches(0.5),
        [{"text": nm, "size": 12, "color": col, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    connector(s, lx, top + Inches(0.5), lx, bot, col, 1.25, dashed=True, arrow=False)
msgs = [
    (Inches(1.6), Inches(4.0), "1: enter Register No", Inches(2.65), ACCENT_BLUE, False),
    (Inches(4.0), Inches(6.66), "2: getCertificate(regNo)", Inches(3.2), ACCENT_GREEN, False),
    (Inches(6.66), Inches(9.0), "3: request signature", Inches(3.75), PURPLE, False),
    (Inches(9.0), Inches(11.5), "4: send call to chain", Inches(4.3), ORANGE, False),
    (Inches(11.5), Inches(6.66), "5: return record data", Inches(4.85), RED, True),
    (Inches(6.66), Inches(4.0), "6: parsed result", Inches(5.4), PURPLE, True),
    (Inches(4.0), Inches(1.6), "7: display certificate", Inches(5.95), ACCENT_GREEN, True),
]
for x1, x2, m, y, col, ret in msgs:
    connector(s, x1, y, x2, y, col, 1.5, dashed=ret)
    midx = min(x1, x2)
    label(s, midx + Inches(0.05), y - Inches(0.3), abs(int(x2 - x1)), m, 10.5, TEXT_WHITE,
          align=PP_ALIGN.CENTER)

# ================================================================ SLIDE 14: UML - ACTIVITY
s = add_slide(); bg(s)
header(s, "UML \u2014 Activity Diagram", "Diagram 4 of 7", 14)
midx = Inches(4.2)
rect(s, midx - Inches(0.16), Inches(1.85), Inches(0.32), Inches(0.32), ACCENT_GREEN, shape=MSO_SHAPE.OVAL)
def act(y, t, col=ACCENT_BLUE, w=Inches(3.2)):
    card(s, midx - w / 2, y, w, Inches(0.6), fill=BG_PANEL, border=col, bw=Pt(1.5))
    txt(s, midx - w / 2, y, w, Inches(0.6),
        [{"text": t, "size": 12.5, "color": TEXT_WHITE, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
act(Inches(2.35), "Open UniVerify portal", ACCENT_BLUE)
act(Inches(3.15), "Enter Register Number", ACCENT_BLUE)
act(Inches(3.95), "Call getCertificate()", PURPLE)
# decision diamond
dy = Inches(4.85)
rect(s, midx - Inches(0.95), dy, Inches(1.9), Inches(0.95), BG_PANEL, ORANGE, Pt(1.5), shape=MSO_SHAPE.DIAMOND)
txt(s, midx - Inches(0.9), dy, Inches(1.8), Inches(0.95),
    [{"text": "Record\nexists?", "size": 11.5, "color": ORANGE, "bold": True, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
connector(s, midx, Inches(2.17), midx, Inches(2.35), ACCENT_GREEN)
connector(s, midx, Inches(2.95), midx, Inches(3.15), ACCENT_BLUE)
connector(s, midx, Inches(3.75), midx, Inches(3.95), ACCENT_BLUE)
connector(s, midx, Inches(4.55), midx, dy, PURPLE)
# yes branch (right) -> display
yes_x = Inches(8.4)
card(s, yes_x - Inches(1.6), Inches(4.9), Inches(3.2), Inches(0.6), fill=BG_PANEL, border=ACCENT_GREEN, bw=Pt(1.5))
txt(s, yes_x - Inches(1.6), Inches(4.9), Inches(3.2), Inches(0.6),
    [{"text": "Display authentic record", "size": 12.5, "color": ACCENT_GREEN, "bold": True, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
connector(s, midx + Inches(0.95), dy + Inches(0.48), yes_x - Inches(1.6), Inches(5.2), ACCENT_GREEN)
label(s, midx + Inches(1.0), dy + Inches(0.05), Inches(1.2), "Yes", 11, ACCENT_GREEN, True, align=PP_ALIGN.LEFT)
# no branch (left)
no_x = Inches(1.7)
card(s, no_x - Inches(0.1), Inches(4.9), Inches(3.0), Inches(0.6), fill=BG_PANEL, border=RED, bw=Pt(1.5))
txt(s, no_x - Inches(0.1), Inches(4.9), Inches(3.0), Inches(0.6),
    [{"text": "Show 'Verification Failed'", "size": 12, "color": RED, "bold": True, "align": PP_ALIGN.CENTER}],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
connector(s, midx - Inches(0.95), dy + Inches(0.48), no_x + Inches(2.9), Inches(5.2), RED)
label(s, midx - Inches(2.3), dy + Inches(0.05), Inches(1.2), "No", 11, RED, True, align=PP_ALIGN.RIGHT)
# end node
ey = Inches(6.3)
rect(s, midx - Inches(0.2), ey, Inches(0.4), Inches(0.4), None, TEXT_WHITE, Pt(2), shape=MSO_SHAPE.OVAL)
rect(s, midx - Inches(0.11), ey + Inches(0.09), Inches(0.22), Inches(0.22), TEXT_WHITE, shape=MSO_SHAPE.OVAL)
connector(s, yes_x, Inches(5.5), midx + Inches(0.05), ey, ACCENT_GREEN)
connector(s, no_x + Inches(1.4), Inches(5.5), midx - Inches(0.05), ey, RED)

# ================================================================ SLIDE 15: UML - STATE CHART
s = add_slide(); bg(s)
header(s, "UML \u2014 State Chart Diagram", "Diagram 5 of 7", 15)
rect(s, Inches(0.9), Inches(3.6), Inches(0.34), Inches(0.34), TEXT_WHITE, shape=MSO_SHAPE.OVAL)
states = [
    (Inches(1.7), Inches(3.45), "Idle", ACCENT_BLUE),
    (Inches(3.85), Inches(3.45), "Wallet\nConnected", PURPLE),
    (Inches(6.0), Inches(3.45), "Querying\nBlockchain", ORANGE),
    (Inches(8.15), Inches(2.5), "Result\nDisplayed", ACCENT_GREEN),
    (Inches(8.15), Inches(4.4), "Error\nState", RED),
]
sw, sh_ = Inches(1.85), Inches(0.9)
pos = {}
for x, y, nm, col in states:
    card(s, x, y, sw, sh_, fill=BG_PANEL, border=col, bw=Pt(1.5))
    txt(s, x, y, sw, sh_, [{"text": nm, "size": 12.5, "color": col, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    pos[nm] = (x, y)
connector(s, Inches(1.24), Inches(3.77), Inches(1.7), Inches(3.85), TEXT_WHITE)
connector(s, Inches(3.55), Inches(3.9), Inches(3.85), Inches(3.9), ACCENT_BLUE)
label(s, Inches(3.4), Inches(3.5), Inches(1.0), "connect", 9.5, TEXT_GREY, italic=True)
connector(s, Inches(5.7), Inches(3.9), Inches(6.0), Inches(3.9), PURPLE)
label(s, Inches(5.5), Inches(3.5), Inches(1.0), "verify", 9.5, TEXT_GREY, italic=True)
connector(s, Inches(7.85), Inches(3.7), Inches(8.15), Inches(2.95), ORANGE)
label(s, Inches(7.55), Inches(3.0), Inches(1.0), "found", 9.5, ACCENT_GREEN, italic=True)
connector(s, Inches(7.85), Inches(4.1), Inches(8.15), Inches(4.85), ORANGE)
label(s, Inches(7.55), Inches(4.6), Inches(1.0), "not found", 9.5, RED, italic=True)
# final
rect(s, Inches(11.4), Inches(3.7), Inches(0.4), Inches(0.4), None, TEXT_WHITE, Pt(2), shape=MSO_SHAPE.OVAL)
rect(s, Inches(11.49), Inches(3.79), Inches(0.22), Inches(0.22), TEXT_WHITE, shape=MSO_SHAPE.OVAL)
connector(s, Inches(10.0), Inches(2.95), Inches(11.45), Inches(3.75), ACCENT_GREEN)
connector(s, Inches(10.0), Inches(4.85), Inches(11.45), Inches(4.05), RED)

# ================================================================ SLIDE 16: UML - COMPONENT
s = add_slide(); bg(s)
header(s, "UML \u2014 Component Diagram", "Diagram 6 of 7", 16)
comps = [
    (Inches(0.8), Inches(2.4), "UI Component", "index.html / CSS", ACCENT_BLUE),
    (Inches(4.0), Inches(2.4), "Web3 Bridge", "Ethers.js (script.js)", ACCENT_GREEN),
    (Inches(7.2), Inches(2.4), "Wallet", "MetaMask provider", ORANGE),
    (Inches(10.2), Inches(2.4), "Smart Contract", "AcademicCertificates", PURPLE),
    (Inches(4.0), Inches(4.7), "Ethereum Node", "Sepolia Testnet", ACCENT_BLUE),
    (Inches(7.2), Inches(4.7), "IPFS Store", "Certificate file hash", ACCENT_GREEN),
]
cw2, ch2 = Inches(2.5), Inches(1.3)
center = {}
for x, y, t, sub, col in comps:
    card(s, x, y, cw2, ch2, fill=BG_PANEL, border=col, bw=Pt(1.5))
    # component tabs
    rect(s, x - Inches(0.06), y + Inches(0.18), Inches(0.28), Inches(0.18), BG_PANEL_2, col, Pt(1.25))
    rect(s, x - Inches(0.06), y + Inches(0.5), Inches(0.28), Inches(0.18), BG_PANEL_2, col, Pt(1.25))
    txt(s, x + Inches(0.3), y + Inches(0.16), cw2 - Inches(0.4), Inches(0.4),
        [{"text": "\u00abcomponent\u00bb", "size": 9, "color": col, "italic": True}])
    txt(s, x + Inches(0.3), y + Inches(0.48), cw2 - Inches(0.4), Inches(0.4),
        [{"text": t, "size": 13, "color": TEXT_WHITE, "bold": True}])
    txt(s, x + Inches(0.3), y + Inches(0.86), cw2 - Inches(0.4), Inches(0.4),
        [{"text": sub, "size": 10, "color": TEXT_GREY}])
    center[t] = (x + cw2 / 2, y + ch2 / 2)
connector(s, Inches(3.3), Inches(3.05), Inches(4.0), Inches(3.05))
connector(s, Inches(6.5), Inches(3.05), Inches(7.2), Inches(3.05), ACCENT_GREEN)
connector(s, Inches(9.7), Inches(3.05), Inches(10.2), Inches(3.05), ORANGE)
connector(s, Inches(8.45), Inches(3.7), Inches(11.45), Inches(5.0), PURPLE, dashed=True)
connector(s, Inches(5.25), Inches(3.7), Inches(5.25), Inches(4.7), ACCENT_GREEN)
connector(s, Inches(8.45), Inches(3.7), Inches(8.45), Inches(4.7), ACCENT_GREEN)
txt(s, Inches(0.8), Inches(6.3), Inches(11.9), Inches(0.5),
    [{"text": "Components communicate through provided/required interfaces \u2014 the Web3 Bridge connects the UI to the wallet, contract, blockchain node and IPFS store.",
      "size": 12.5, "color": TEXT_GREY, "italic": True, "line_spacing": 1.15}])

# ================================================================ SLIDE 17: UML - DEPLOYMENT
s = add_slide(); bg(s)
header(s, "UML \u2014 Deployment Diagram", "Diagram 7 of 7", 17)
def dnode(x, y, w, h, title, items, col):
    rect(s, x, y, w, h, BG_PANEL, col, Pt(1.75), shape=MSO_SHAPE.CUBE)
    txt(s, x + Inches(0.35), y + Inches(0.2), w - Inches(0.5), Inches(0.4),
        [{"text": "\u00abdevice\u00bb", "size": 9.5, "color": col, "italic": True},
         {"text": title, "size": 14, "color": TEXT_WHITE, "bold": True, "space_before": 1}])
    yy = y + Inches(0.95)
    for it, c2 in items:
        card(s, x + Inches(0.35), yy, w - Inches(0.75), Inches(0.55), fill=BG_PANEL_2, border=c2, bw=Pt(1.25))
        txt(s, x + Inches(0.35), yy, w - Inches(0.75), Inches(0.55),
            [{"text": it, "size": 11, "color": TEXT_WHITE, "bold": True, "align": PP_ALIGN.CENTER}],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        yy += Inches(0.7)
dnode(Inches(0.8), Inches(2.1), Inches(3.7), Inches(3.6), "Client Browser",
      [("UniVerify Web UI", ACCENT_BLUE), ("MetaMask Extension", ORANGE), ("Ethers.js Library", ACCENT_GREEN)], ACCENT_BLUE)
dnode(Inches(5.3), Inches(2.1), Inches(3.7), Inches(3.6), "Ethereum Sepolia Node",
      [("AcademicCertificates", PURPLE), ("Distributed Ledger", ACCENT_BLUE)], PURPLE)
dnode(Inches(9.8), Inches(2.1), Inches(2.85), Inches(3.6), "IPFS Network",
      [("Certificate Files", ACCENT_GREEN)], ACCENT_GREEN)
connector(s, Inches(4.5), Inches(3.6), Inches(5.5), Inches(3.6), TEXT_WHITE, 1.75, both=True)
label(s, Inches(4.3), Inches(3.2), Inches(1.4), "JSON-RPC / HTTPS", 9.5, TEXT_GREY, italic=True)
connector(s, Inches(9.0), Inches(3.6), Inches(10.0), Inches(3.6), TEXT_WHITE, 1.75, both=True)
label(s, Inches(8.85), Inches(3.2), Inches(1.3), "hash link", 9.5, TEXT_GREY, italic=True)

# ================================================================ SLIDE 18: ALGORITHM
s = add_slide(); bg(s)
header(s, "Algorithm Used in Proposed Solution", "Core Logic", 18)
# Issuance algorithm
card(s, Inches(0.7), Inches(1.85), Inches(5.85), Inches(4.7), fill=BG_PANEL, border=PURPLE, bw=Pt(1.25))
txt(s, Inches(0.95), Inches(2.05), Inches(5.4), Inches(0.4),
    [{"text": "A.  Certificate Issuance (Admin)", "size": 15, "color": PURPLE, "bold": True}])
bullets(s, Inches(1.0), Inches(2.6), Inches(5.4), Inches(3.9),
        ["Start::admin connects MetaMask wallet.",
         "Input::Register No, Name, Course, IPFS hash.",
         "Check::require(msg.sender == university).",
         "If not owner::revert transaction, stop.",
         "Else::certificates[regNo] = Certificate(...).",
         "Commit::transaction mined & stored on-chain.",
         "End::record is permanent and immutable."],
        size=12.5, gap=11, color=TEXT_GREY, marker="\u2192", marker_color=PURPLE)
# Verification algorithm
card(s, Inches(6.8), Inches(1.85), Inches(5.85), Inches(4.7), fill=BG_PANEL, border=ACCENT_GREEN, bw=Pt(1.25))
txt(s, Inches(7.05), Inches(2.05), Inches(5.4), Inches(0.4),
    [{"text": "B.  Certificate Verification (Anyone)", "size": 15, "color": ACCENT_GREEN, "bold": True}])
bullets(s, Inches(7.1), Inches(2.6), Inches(5.4), Inches(3.9),
        ["Start::user enters Register Number.",
         "Call::contract.getCertificate(regNo) via Ethers.js.",
         "Lookup::mapping returns record in O(1) time.",
         "If record empty::call reverts \u2192 'Verification Failed'.",
         "Else::receive (name, course, hash).",
         "Render::display authentic record to user.",
         "End::verification complete in seconds."],
        size=12.5, gap=11, color=TEXT_WHITE, marker="\u2192", marker_color=ACCENT_GREEN)

# ================================================================ SLIDE 19: OUTPUT SCREENSHOTS
s = add_slide(); bg(s)
header(s, "Output Screenshots", "Live Demonstration", 19)
def mockui(x, title, regno, ok, sname, scourse, col):
    card(s, x, Inches(1.95), Inches(3.8), Inches(4.55), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.5))
    txt(s, x, Inches(2.2), Inches(3.8), Inches(0.45),
        [{"text": "UniVerify V2.0", "size": 18, "color": ACCENT_BLUE, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER)
    txt(s, x, Inches(2.65), Inches(3.8), Inches(0.35),
        [{"text": title, "size": 11, "color": TEXT_GREY, "align": PP_ALIGN.CENTER}], align=PP_ALIGN.CENTER)
    card(s, x + Inches(0.35), Inches(3.15), Inches(3.1), Inches(0.5), fill=BG_PANEL_2, border=LINE_SUBTLE)
    txt(s, x + Inches(0.5), Inches(3.22), Inches(2.9), Inches(0.36),
        [{"text": regno, "size": 12, "color": TEXT_WHITE, "font": FONT_MONO}], anchor=MSO_ANCHOR.MIDDLE)
    card(s, x + Inches(0.35), Inches(3.75), Inches(3.1), Inches(0.5), fill=RGBColor(0x23, 0x86, 0x36), border=None)
    txt(s, x + Inches(0.35), Inches(3.82), Inches(3.1), Inches(0.36),
        [{"text": "Verify on Blockchain", "size": 12, "color": TEXT_WHITE, "bold": True, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rb = RGBColor(0x10, 0x26, 0x16) if ok else RGBColor(0x2b, 0x12, 0x12)
    card(s, x + Inches(0.35), Inches(4.4), Inches(3.1), Inches(1.85), fill=rb, border=col, bw=Pt(1.25))
    if ok:
        txt(s, x + Inches(0.5), Inches(4.55), Inches(2.85), Inches(1.7),
            [{"text": "\u2705 AUTHENTIC RECORD", "size": 12, "color": col, "bold": True, "space_after": 5},
             {"text": f"Student: {sname}", "size": 11, "color": TEXT_WHITE, "space_after": 2},
             {"text": f"Course: {scourse}", "size": 11, "color": TEXT_WHITE, "space_after": 2},
             {"text": "Hash: Qm7x...a91f", "size": 10.5, "color": TEXT_GREY, "font": FONT_MONO}])
    else:
        txt(s, x + Inches(0.5), Inches(4.55), Inches(2.85), Inches(1.7),
            [{"text": "\u274C VERIFICATION FAILED", "size": 12, "color": col, "bold": True, "space_after": 5},
             {"text": "No record found for this", "size": 11, "color": TEXT_WHITE, "space_after": 2},
             {"text": "Register Number on-chain.", "size": 11, "color": TEXT_WHITE}])
mockui(Inches(0.85), "Successful Verification", "231061101493", True, "Aarav Sharma", "B.Tech CSE", ACCENT_GREEN)
mockui(Inches(4.77), "Invalid / Forged Record", "000000000000", False, "", "", RED)
# notes panel
card(s, Inches(8.7), Inches(1.95), Inches(3.9), Inches(4.55), fill=BG_PANEL, border=LINE_SUBTLE)
txt(s, Inches(8.95), Inches(2.2), Inches(3.5), Inches(0.4),
    [{"text": "Demonstrated Output", "size": 15, "color": ACCENT_BLUE, "bold": True}])
bullets(s, Inches(8.95), Inches(2.75), Inches(3.45), Inches(3.6),
        ["Contract deployed on Sepolia via Remix IDE.",
         "Certificate issued from university wallet.",
         "Valid Reg No returns authentic record.",
         "Invalid Reg No shows clear failure message.",
         "Records confirmed tamper-proof on-chain."],
        size=12, gap=12, color=TEXT_WHITE, marker="\u2713", marker_color=ACCENT_GREEN)
txt(s, Inches(0.85), Inches(6.55), Inches(7.7), Inches(0.3),
    [{"text": "Note: replace these mockups with actual Remix / MetaMask / portal screenshots.",
      "size": 10, "color": TEXT_GREY, "italic": True}])

# ================================================================ SLIDE 20: COMPARATIVE GRAPH
s = add_slide(); bg(s)
header(s, "Comparative Analysis", "Existing System vs. Proposed System", 20)
chart_data = CategoryChartData()
chart_data.categories = ["Security", "Speed", "Transparency", "Tamper-\nResistance", "Trust", "Global\nAccess"]
chart_data.add_series("Existing System", (4, 3, 2, 2, 3, 3))
chart_data.add_series("UniVerify (Proposed)", (9, 9, 10, 10, 9, 10))
gframe = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.8), Inches(1.9),
                            Inches(8.3), Inches(4.7), chart_data)
chart = gframe.chart
chart.has_title = True
chart.chart_title.text_frame.text = "Capability Score (out of 10)"
chart.chart_title.text_frame.paragraphs[0].font.size = Pt(13)
chart.chart_title.text_frame.paragraphs[0].font.color.rgb = TEXT_WHITE
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
chart.legend.font.color.rgb = TEXT_WHITE
chart.legend.font.size = Pt(11)
plot = chart.plots[0]
plot.gap_width = 80
plot.series[0].format.fill.solid(); plot.series[0].format.fill.fore_color.rgb = ORANGE
plot.series[1].format.fill.solid(); plot.series[1].format.fill.fore_color.rgb = ACCENT_GREEN
cat_ax = chart.category_axis
cat_ax.tick_labels.font.color.rgb = TEXT_WHITE; cat_ax.tick_labels.font.size = Pt(10)
val_ax = chart.value_axis
val_ax.tick_labels.font.color.rgb = TEXT_GREY; val_ax.tick_labels.font.size = Pt(10)
val_ax.maximum_scale = 10
# takeaway panel
card(s, Inches(9.4), Inches(1.95), Inches(3.2), Inches(4.6), fill=BG_PANEL, border=ACCENT_GREEN, bw=Pt(1.25))
txt(s, Inches(9.65), Inches(2.2), Inches(2.8), Inches(0.4),
    [{"text": "Key Insight", "size": 15, "color": ACCENT_GREEN, "bold": True}])
bullets(s, Inches(9.65), Inches(2.75), Inches(2.75), Inches(3.6),
        ["Far higher security & tamper-resistance.",
         "Verification in seconds, not days.",
         "Fully transparent & publicly auditable.",
         "Globally accessible, no middlemen.",
         "Trust shifts from people to cryptography."],
        size=12, gap=13, color=TEXT_WHITE, marker="\u25B8", marker_color=ACCENT_GREEN)

# ================================================================ SLIDE 21: FUTURE + CONCLUSION
s = add_slide(); bg(s)
header(s, "Future Enhancement & Conclusion", "The Road Ahead", 21)
card(s, Inches(0.7), Inches(1.85), Inches(5.85), Inches(4.7), fill=BG_PANEL, border=ACCENT_GREEN, bw=Pt(1.25))
txt(s, Inches(0.95), Inches(2.05), Inches(5.4), Inches(0.4),
    [{"text": "Future Enhancements", "size": 16, "color": ACCENT_GREEN, "bold": True}])
bullets(s, Inches(1.0), Inches(2.6), Inches(5.4), Inches(3.9),
        ["Multi-institution role-based access control.",
         "Full IPFS / decentralized document storage.",
         "QR-code based instant mobile verification.",
         "NFT / soulbound-token based degrees.",
         "Deployment on Ethereum / Polygon mainnet.",
         "Bulk issuance & analytics dashboard."],
        size=13.5, gap=13, color=TEXT_WHITE, marker="\u25B8", marker_color=ACCENT_GREEN)
card(s, Inches(6.8), Inches(1.85), Inches(5.85), Inches(4.7), fill=BG_PANEL, border=BORDER_BLUE, bw=Pt(1.25))
txt(s, Inches(7.05), Inches(2.05), Inches(5.4), Inches(0.4),
    [{"text": "Conclusion", "size": 16, "color": ACCENT_BLUE, "bold": True}])
txt(s, Inches(7.1), Inches(2.6), Inches(5.3), Inches(3.9),
    [{"text": "UniVerify V2.0 shows how blockchain can solve academic credential forgery by making records immutable, transparent and instantly verifiable.",
      "size": 14, "color": TEXT_WHITE, "bold": True, "line_spacing": 1.25, "space_after": 12},
     {"text": "By combining Solidity smart contracts, the Ethereum Sepolia testnet and an Ethers.js Web3 frontend, the system removes slow centralized verification while guaranteeing data integrity through cryptography.",
      "size": 13, "color": TEXT_GREY, "line_spacing": 1.25, "space_after": 12},
     {"text": "Immutable  \u2022  Instant  \u2022  Trustless", "size": 15, "color": ACCENT_GREEN, "bold": True}])

# ================================================================ SLIDE 22: REFERENCES
s = add_slide(); bg(s)
header(s, "References", "Resources", 22)
refs = [
    "Gresch, J. et al. (2018). The Proof is in the Pudding: Blockchain-based University Diplomas. IEEE.",
    "Cheng, J. C. et al. (2018). Blockchain & Smart Contract for Digital Certificate. IEEE ICASI.",
    "Saleh, O. S. et al. (2020). Blockchain + IPFS for Tamper-Proof Document Storage.",
    "Solidity Documentation \u2014 docs.soliditylang.org",
    "Ethers.js v5 Documentation \u2014 docs.ethers.org/v5",
    "Ethereum Sepolia Testnet \u2014 ethereum.org/developers",
    "MetaMask Developer Docs \u2014 docs.metamask.io",
    "Remix IDE \u2014 remix.ethereum.org",
    "Project Repository \u2014 github.com/shaikmofficial-ai/Univerify",
]
card(s, Inches(0.7), Inches(1.85), Inches(11.9), Inches(3.4), fill=BG_PANEL, border=LINE_SUBTLE)
bullets(s, Inches(1.05), Inches(2.15), Inches(11.2), Inches(2.9), refs,
        size=13, gap=8, color=TEXT_GREY, marker="\u25B8", marker_color=ACCENT_BLUE)
card(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.25), fill=BG_PANEL_2, border=ACCENT_BLUE, bw=Pt(1.5))
txt(s, Inches(0.7), Inches(5.58), Inches(11.9), Inches(0.6),
    [{"text": "Thank You!", "size": 30, "color": ACCENT_BLUE, "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_TITLE}],
    align=PP_ALIGN.CENTER)
txt(s, Inches(0.7), Inches(6.18), Inches(11.9), Inches(0.4),
    [{"text": "Questions & Discussion  \u2022  UniVerify V2.0  \u2014  Decentralized Academic Credential Verification",
      "size": 12.5, "color": TEXT_GREY, "align": PP_ALIGN.CENTER}], align=PP_ALIGN.CENTER)

# ----------------------------------------------------------------
prs.save("/projects/sandbox/Univerify/UniVerify_Final_Presentation.pptx")
print("Saved presentation with", len(prs.slides._sldIdLst), "slides.")
