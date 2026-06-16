from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Palette ──────────────────────────────────────────────────────────────────
C_BLUE       = RGBColor(0x00, 0x6E, 0xFF)
C_DARK       = RGBColor(0x0D, 0x1B, 0x2A)
C_NAVY       = RGBColor(0x1A, 0x2E, 0x4A)
C_TEAL       = RGBColor(0x00, 0xC8, 0xAA)
C_ORANGE     = RGBColor(0xFF, 0x6B, 0x35)
C_PURPLE     = RGBColor(0x7B, 0x2D, 0xFF)
C_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
C_LGRAY      = RGBColor(0xF4, 0xF6, 0xFA)
C_MGRAY      = RGBColor(0xD0, 0xD7, 0xE3)
C_DGRAY      = RGBColor(0x55, 0x65, 0x78)
C_YELLOW     = RGBColor(0xFF, 0xD6, 0x00)
C_GREEN      = RGBColor(0x27, 0xAE, 0x60)
C_RED        = RGBColor(0xE7, 0x4C, 0x3C)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ── Primitive helpers ─────────────────────────────────────────────────────────
def rect(slide, l, t, w, h, fill, line_color=None, line_width=None, radius=0):
    s = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line_color:
        s.line.color.rgb = line_color
        if line_width: s.line.width = Pt(line_width)
    else:
        s.line.fill.background()
    if radius:
        # round corners via XML
        sp = s._element
        prstGeom = sp.find('.//' + qn('a:prstGeom'))
        if prstGeom is not None:
            prstGeom.set('prst', 'roundRect')
            avLst = prstGeom.find(qn('a:avLst'))
            if avLst is None:
                avLst = etree.SubElement(prstGeom, qn('a:avLst'))
            gd = etree.SubElement(avLst, qn('a:gd'))
            gd.set('name', 'adj'); gd.set('fmla', f'val {radius}')
    return s

def text_box(slide, txt, l, t, w, h, size=16, bold=False, italic=False,
             color=C_WHITE, align=PP_ALIGN.LEFT, wrap=True, font="Segoe UI"):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = txt
    run.font.size = Pt(size); run.font.bold = bold
    run.font.italic = italic; run.font.name = font
    run.font.color.rgb = color
    return tb

def footer(slide, page_num=None, total=16):
    rect(slide, 0, 7.1, 13.33, 0.4, C_DARK)
    text_box(slide, "WalkMe Builder 1  ·  Instructor-Led Training  ·  Beginner Level",
             0.3, 7.12, 10, 0.32, size=10, color=RGBColor(0x88,0x99,0xAA))
    if page_num:
        text_box(slide, f"{page_num} / {total}", 12.1, 7.12, 1.0, 0.32,
                 size=10, color=C_TEAL, align=PP_ALIGN.RIGHT)

def section_tag(slide, label, color=C_TEAL):
    rect(slide, 0.35, 1.52, 2.2, 0.32, color, radius=15000)
    text_box(slide, label, 0.38, 1.53, 2.15, 0.3, size=11, bold=True,
             color=C_WHITE, align=PP_ALIGN.CENTER)

# ── Slide header ─────────────────────────────────────────────────────────────
def header(slide, title, subtitle="", tag="", tag_color=C_TEAL):
    rect(slide, 0, 0, 13.33, 1.48, C_NAVY)
    rect(slide, 0, 1.48, 13.33, 0.07, C_BLUE)
    # accent bar left
    rect(slide, 0, 0, 0.22, 1.48, C_BLUE)
    text_box(slide, title, 0.42, 0.1, 12.4, 0.85,
             size=30, bold=True, color=C_WHITE)
    if subtitle:
        text_box(slide, subtitle, 0.42, 0.88, 12.4, 0.55,
                 size=14, color=RGBColor(0xAA,0xCC,0xFF), italic=True)
    if tag:
        section_tag(slide, tag, tag_color)

# ── WalkMe editor mock screenshot ─────────────────────────────────────────────
def mock_editor(slide, l, t, w, h, highlight_step=None):
    """Draw a schematic of the WalkMe floating editor panel."""
    # Window chrome
    rect(slide, l, t, w, h, RGBColor(0x1E,0x1E,0x2E), line_color=C_BLUE, line_width=1.5, radius=12000)
    # Title bar
    rect(slide, l, t, w, 0.38, C_BLUE)
    text_box(slide, "⚡  WalkMe Editor", l+0.1, t+0.04, w-0.2, 0.3,
             size=11, bold=True, color=C_WHITE)
    # traffic lights
    for i, c in enumerate([C_RED, C_YELLOW, C_GREEN]):
        cx = Inches(l + w - 0.22 - i*0.22)
        cy = Inches(t + 0.12)
        from pptx.util import Emu
        circ = slide.shapes.add_shape(9,  # oval
            cx, cy, Inches(0.14), Inches(0.14))
        circ.fill.solid(); circ.fill.fore_color.rgb = c
        circ.line.fill.background()

    # Sidebar items
    sidebar_items = ["+ New Item","","▸ Walk-Thrus","  ○ Expense SWT","  ○ Onboarding","▸ SmartTips","  ○ Field Help","▸ ShoutOuts","▸ Launchers","▸ Resources"]
    for i, item in enumerate(sidebar_items):
        iy = t + 0.45 + i*0.28
        if iy + 0.28 > t + h - 0.05: break
        bg = C_BLUE if highlight_step and item.strip().startswith(highlight_step) else RGBColor(0x25,0x25,0x40)
        if item == "+ New Item":
            rect(slide, l+0.05, iy, w-0.1, 0.25, C_TEAL, radius=8000)
            text_box(slide, item, l+0.08, iy+0.02, w-0.15, 0.22, size=9, bold=True, color=C_DARK)
        elif item == "":
            rect(slide, l+0.05, iy+0.1, w-0.1, 0.02, RGBColor(0x33,0x33,0x55))
        else:
            text_box(slide, item, l+0.08, iy+0.02, w-0.15, 0.22, size=9,
                     color=C_WHITE if "▸" in item else RGBColor(0xAA,0xBB,0xCC))

def mock_browser(slide, l, t, w, h, url="https://demo.walkme.com", highlight=None):
    """Draw a schematic browser window with WalkMe bubble overlay."""
    # Window
    rect(slide, l, t, w, h, C_WHITE, line_color=C_MGRAY, line_width=1)
    # Browser chrome
    rect(slide, l, t, w, 0.4, C_LGRAY, line_color=C_MGRAY, line_width=0.5)
    # Traffic lights
    for i, c in enumerate([C_RED, C_YELLOW, C_GREEN]):
        circ = slide.shapes.add_shape(9, Inches(l+0.1+i*0.22), Inches(t+0.13), Inches(0.14), Inches(0.14))
        circ.fill.solid(); circ.fill.fore_color.rgb = c; circ.line.fill.background()
    # URL bar
    rect(slide, l+0.85, t+0.07, w-1.0, 0.26, C_WHITE, line_color=C_MGRAY, line_width=0.5, radius=6000)
    text_box(slide, url, l+0.95, t+0.09, w-1.2, 0.22, size=8, color=C_DGRAY)
    # Page content simulation
    rect(slide, l+0.1, t+0.5, w-0.2, 0.12, C_LGRAY)   # nav bar
    rect(slide, l+0.1, t+0.72, (w-0.3)*0.65, h-1.1, C_LGRAY)  # main
    rect(slide, l+0.1+(w-0.3)*0.68, t+0.72, (w-0.3)*0.3, h-1.1, C_LGRAY)  # sidebar
    # Form fields
    fy = t + 0.85
    for label in ["First Name", "Last Name", "Email"]:
        text_box(slide, label, l+0.18, fy, 1.5, 0.18, size=7, color=C_DGRAY)
        rect(slide, l+0.18, fy+0.18, 2.4, 0.22, C_WHITE, line_color=C_MGRAY, line_width=0.5)
        fy += 0.55

    if highlight == "bubble":
        # WalkMe bubble
        bx, by = l + 2.7, t + 0.75
        rect(slide, bx, by, 2.2, 0.9, C_BLUE, radius=10000)
        text_box(slide, "👋 Click 'First Name'\nand enter your name.", bx+0.1, by+0.07, 2.0, 0.78,
                 size=8, color=C_WHITE, wrap=True)
        # Tail
        tail = slide.shapes.add_shape(5,  # right-triangle
            Inches(bx-0.15), Inches(by+0.3), Inches(0.18), Inches(0.18))
        tail.fill.solid(); tail.fill.fore_color.rgb = C_BLUE; tail.line.fill.background()
        # Step counter
        rect(slide, bx+1.75, by-0.16, 0.55, 0.28, C_TEAL, radius=8000)
        text_box(slide, "1 / 4", bx+1.77, by-0.14, 0.5, 0.24, size=7, bold=True, color=C_DARK, align=PP_ALIGN.CENTER)

    if highlight == "smarttip":
        bx, by = l + 2.7, t + 1.65
        rect(slide, bx, by, 2.0, 0.7, C_DARK, radius=8000)
        rect(slide, bx, by, 2.0, 0.04, C_TEAL)
        text_box(slide, "💡 Use your work email\n(e.g. name@company.com)", bx+0.1, by+0.08, 1.85, 0.6,
                 size=7.5, color=C_WHITE, wrap=True)

    if highlight == "shoutout":
        rect(slide, l+0.05, t+0.45, w-0.1, 0.5, C_BLUE)
        text_box(slide, "🎉  New Feature Alert!  Try the new Dashboard  →  Explore Now",
                 l+0.2, t+0.48, w-0.4, 0.42, size=9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    if highlight == "launcher":
        lbx, lby = l+w-0.65, t+h-0.6
        circ = slide.shapes.add_shape(9, Inches(lbx), Inches(lby), Inches(0.48), Inches(0.48))
        circ.fill.solid(); circ.fill.fore_color.rgb = C_BLUE; circ.line.fill.background()
        text_box(slide, "?", lbx+0.13, lby+0.04, 0.25, 0.38, size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

def step_card(slide, num, title, desc, l, t, w=3.8, h=1.5,
              num_color=C_BLUE, icon=""):
    rect(slide, l, t, w, h, C_WHITE, radius=10000)
    rect(slide, l, t, 0.55, h, num_color, radius=10000)
    # Fix rounded left side visibility
    rect(slide, l+0.35, t, 0.22, h, num_color)
    text_box(slide, str(num), l+0.08, t+h/2-0.28, 0.45, 0.55,
             size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    if icon:
        text_box(slide, icon, l+0.65, t+0.08, 0.45, 0.45, size=20)
    text_box(slide, title, l+0.65+(0.5 if icon else 0), t+0.1, w-1.3, 0.42,
             size=13, bold=True, color=C_NAVY)
    text_box(slide, desc, l+0.65, t+0.52, w-0.85, h-0.62,
             size=10.5, color=C_DGRAY, wrap=True)

def callout_box(slide, icon, title, body, l, t, w, h, bg=C_LGRAY, accent=C_BLUE):
    rect(slide, l, t, w, h, bg, radius=10000)
    rect(slide, l, t, 0.12, h, accent)
    text_box(slide, icon, l+0.2, t+0.1, 0.5, 0.5, size=24)
    text_box(slide, title, l+0.75, t+0.1, w-0.9, 0.4, size=13, bold=True, color=accent)
    text_box(slide, body, l+0.75, t+0.5, w-0.9, h-0.6, size=11, color=C_DGRAY, wrap=True)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1 — Cover
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_DARK)
rect(s, 0, 0, 6.5, 7.5, C_NAVY)
rect(s, 6.5, 0, 6.83, 7.5, C_DARK)

# Diagonal accent
for i in range(8):
    rect(s, 5.8+i*0.12, 0, 0.08, 7.5, RGBColor(0x00,0x6E,0xFF) if i%2==0 else C_DARK)

rect(s, 0, 0, 0.35, 7.5, C_BLUE)
rect(s, 0.35, 0, 0.08, 7.5, C_TEAL)

# Logo area
rect(s, 0.6, 0.55, 2.5, 0.55, C_BLUE, radius=8000)
text_box(s, "  WalkMe™", 0.65, 0.58, 2.4, 0.5, size=16, bold=True, color=C_WHITE)

text_box(s, "Builder 1", 0.6, 1.35, 5.5, 1.1, size=60, bold=True, color=C_WHITE, font="Segoe UI Black")
text_box(s, "Instructor-Led Training", 0.6, 2.4, 5.5, 0.65, size=26, bold=True, color=C_TEAL)
text_box(s, "Beginner Level  ·  Full-Day Course", 0.6, 3.05, 5.5, 0.5, size=16, italic=True,
         color=RGBColor(0x99,0xBB,0xDD))
rect(s, 0.6, 3.65, 5.0, 0.06, C_TEAL)

# Right column — course meta
for icon, label, val in [
    ("🕐", "Duration", "8 Hours"),
    ("📋", "Format",   "Instructor-Led"),
    ("🎯", "Level",    "Beginner"),
    ("💻", "Mode",     "Hands-On Labs"),
]:
    iy = {"🕐":4.1,"📋":4.85,"🎯":5.6,"💻":6.35}[icon]
    rect(s, 7.2, iy, 5.7, 0.68, RGBColor(0x1E,0x2E,0x45), radius=8000)
    text_box(s, icon, 7.35, iy+0.09, 0.5, 0.5, size=20)
    text_box(s, label, 7.9, iy+0.06, 1.3, 0.3, size=11, color=C_TEAL)
    text_box(s, val,   7.9, iy+0.33, 2.5, 0.3, size=14, bold=True, color=C_WHITE)

text_box(s, "Master the WalkMe Editor and build your first\nguided experiences — no coding required!",
         7.2, 1.35, 5.7, 1.2, size=14, color=RGBColor(0xAA,0xCC,0xFF), wrap=True)

footer(s, 1)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2 — Agenda
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "Course Agenda", "What we will cover today — 7 modules + hands-on lab", tag="Overview")

modules = [
    ("1", C_BLUE,   "🌐", "Introduction to WalkMe",           "What WalkMe is, why it matters, key terminology"),
    ("2", C_PURPLE, "🖥️", "The WalkMe Editor",                 "Interface walkthrough & element capture"),
    ("3", C_TEAL,   "🚶", "Smart Walk-Thrus (SWT)",            "Build step-by-step guided flows"),
    ("4", C_ORANGE, "💬", "SmartTips & ShoutOuts",             "Tooltips, banners, and announcements"),
    ("5", C_GREEN,  "🚀", "Launchers & Resources",             "On-demand triggers and help content"),
    ("6", C_BLUE,   "📊", "Publishing, Testing & Analytics",  "Preview, publish, and measure impact"),
    ("7", C_PURPLE, "🔬", "Hands-On Lab & Q&A",               "Build a complete flow from scratch"),
]
positions = [
    (0.35, 1.73), (4.65, 1.73), (8.95, 1.73),
    (0.35, 3.65), (4.65, 3.65), (8.95, 3.65),
    (0.35, 5.57),
]
for idx, (num, col, icon, title, desc) in enumerate(modules):
    l, t = positions[idx]
    w = 8.2 if idx==6 else 3.9
    step_card(s, num, f"{icon}  {title}", desc, l, t, w=w, h=1.6, num_color=col)

footer(s, 2)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3 — What is WalkMe?
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "What is WalkMe?", "Module 1 — Introduction to Digital Adoption", tag="Module 1", tag_color=C_BLUE)

# Big definition card
rect(s, 0.35, 1.75, 12.63, 1.18, C_BLUE, radius=10000)
text_box(s, "💡", 0.55, 1.82, 0.7, 0.9, size=36)
text_box(s, "WalkMe is a Digital Adoption Platform (DAP)", 1.3, 1.82, 9.0, 0.52, size=22, bold=True, color=C_WHITE)
text_box(s, "It overlays real-time guidance on top of any web application — no code changes needed by end users.",
         1.3, 2.32, 11.2, 0.5, size=14, color=RGBColor(0xCC,0xE5,0xFF))

callouts = [
    ("👥", "Who Uses It?",       "HR, IT, Sales Ops, L&D, Customer Success teams who want users to succeed in enterprise software.",       0.35, 3.1, 3.8, 1.55),
    ("🎯", "What Problem?",      "Users forget training, skip steps, and call support. WalkMe shows them exactly what to do — live.",        4.4,  3.1, 3.8, 1.55),
    ("🔧", "How Does It Work?",  "A browser extension (test) or script tag (production) injects WalkMe guidance into any web page.",          8.45, 3.1, 4.55, 1.55),
    ("📦", "Content You Build",  "Smart Walk-Thrus · SmartTips · ShoutOuts · Launchers · Resources — all created in the WalkMe Editor.",     0.35, 4.82, 12.63, 1.25),
]
for icon, title, body, l, t, w, h in callouts:
    callout_box(s, icon, title, body, l, t, w, h)

footer(s, 3)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 4 — Key Terminology
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "Key Terminology", "Module 1 — Know the language before you build", tag="Module 1", tag_color=C_BLUE)

terms = [
    ("Step",        C_BLUE,   "A single instruction inside a Walk-Thru (e.g., 'Click Save', 'Enter your name')."),
    ("Flow",        C_PURPLE, "A sequence of Steps guiding a user through a complete task end-to-end."),
    ("Selector",    C_TEAL,   "The unique fingerprint WalkMe uses to identify a specific element on the page."),
    ("Trigger",     C_ORANGE, "The condition that starts a WalkMe item — auto on page load, on click, by URL, etc."),
    ("Segment",     C_GREEN,  "A rule to show content only to specific users (e.g., role, location, system data)."),
    ("Environment", C_BLUE,   "Separate workspaces: Test (for building) vs. Production (what users see)."),
    ("Publish",     C_PURPLE, "The action of pushing saved Draft content live so end users can see it."),
]
col_w = 6.1
for i, (term, col, defn) in enumerate(terms):
    row, col_idx = divmod(i, 2)
    l = 0.35 + col_idx * (col_w + 0.48)
    t = 1.73 + row * 1.22
    if i == 6:  # last item full width
        l, col_w2 = 0.35, 12.63
    else:
        col_w2 = col_w
    rect(s, l, t, col_w2, 1.1, C_WHITE, radius=8000)
    rect(s, l, t, 1.1, 1.1, col, radius=8000)
    rect(s, l+0.8, t, 0.35, 1.1, col)
    text_box(s, term[0], l+0.2, t+0.22, 0.7, 0.6, size=24, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    text_box(s, term, l+1.2, t+0.06, col_w2-1.4, 0.4, size=14, bold=True, color=col)
    text_box(s, defn, l+1.2, t+0.46, col_w2-1.4, 0.58, size=11, color=C_DGRAY, wrap=True)

footer(s, 4)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 5 — The WalkMe Editor (screenshot mock)
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_DARK)
header(s, "The WalkMe Editor", "Module 2 — Your Command Center", tag="Module 2", tag_color=C_PURPLE)

# Draw the mocked editor
mock_editor(s, 0.35, 1.62, 3.4, 5.5)

# Annotations
annots = [
    (3.9, 1.75, C_TEAL,   "① WalkMe Toolbar", "The floating editor appears as a panel in your browser when the extension is active."),
    (3.9, 2.68, C_BLUE,   "② '+ New Item' Button", "Click this to create any new content: Walk-Thru, SmartTip, ShoutOut, Launcher, or Resource."),
    (3.9, 3.61, C_ORANGE, "③ Content Library", "All your saved items live here, grouped by type. Click any item to edit it."),
    (3.9, 4.54, C_GREEN,  "④ Element Capture", "The crosshair 🎯 icon activates capture mode — hover and click any page element to build a selector."),
    (3.9, 5.47, C_PURPLE, "⑤ Publish Button", "When ready, click Publish to push your content live to the selected environment."),
]
for l, t, col, title, body in annots:
    rect(s, l, t, 9.08, 0.78, RGBColor(0x1A,0x2A,0x3E), radius=8000)
    rect(s, l, t, 0.1, 0.78, col)
    text_box(s, title, l+0.22, t+0.05, 3.5, 0.32, size=13, bold=True, color=col)
    text_box(s, body,  l+0.22, t+0.38, 8.6, 0.35, size=11, color=RGBColor(0xCC,0xDD,0xEE), wrap=True)

footer(s, 5)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 6 — Smart Walk-Thru Overview
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "Smart Walk-Thrus (SWT)", "Module 3 — Step-by-step guided flows", tag="Module 3", tag_color=C_TEAL)

# Left — description cards
cards = [
    ("🚶", "What is an SWT?",    "An interactive overlay that walks users through any task, step by step, live on the page."),
    ("📋", "Step Types",         "Bubble (tooltip), Click, Type, Select, Alert, Hotspot, Frame, and Auto-Step."),
    ("🎯", "Element Capture",    "WalkMe auto-generates selectors. You click the element — it handles the rest."),
    ("⚡", "Smart Actions",      "Steps advance automatically when the user clicks, types, or a condition is met — no back button needed."),
    ("✅", "Best Practice",      "Keep flows to 5–8 steps. Break large processes into multiple focused Walk-Thrus."),
]
for i, (icon, title, body) in enumerate(cards):
    t = 1.72 + i * 1.02
    rect(s, 0.35, t, 6.5, 0.93, C_WHITE, radius=8000)
    text_box(s, icon, 0.48, t+0.18, 0.5, 0.55, size=22)
    text_box(s, title, 1.05, t+0.07, 2.8, 0.38, size=13, bold=True, color=C_NAVY)
    text_box(s, body,  1.05, t+0.45, 5.6, 0.42, size=11, color=C_DGRAY, wrap=True)

# Right — browser mock with SWT bubble
mock_browser(s, 7.0, 1.62, 6.0, 5.5, highlight="bubble")
# Label
rect(s, 7.05, 1.65, 5.9, 0.3, C_TEAL)
text_box(s, "▶ Live preview — SWT Bubble on a form page", 7.15, 1.67, 5.7, 0.26,
         size=10, bold=True, color=C_DARK)

footer(s, 6)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 7 — How to Build an SWT (Step-by-step process)
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_DARK)
header(s, "How to Build a Smart Walk-Thru", "Module 3 — Follow these 8 steps", tag="Module 3", tag_color=C_TEAL)

steps_swt = [
    (C_BLUE,   "Open the WalkMe Editor",       "Click the WalkMe browser extension icon → editor panel opens."),
    (C_TEAL,   "Click  '+ New Item'",           "From the dropdown, choose  'Smart Walk-Thru'."),
    (C_PURPLE, "Name Your Walk-Thru",           "Use an action phrase:  'Submit Expense Report'  (verb + object)."),
    (C_ORANGE, "Click  '+ Add Step'",           "Choose step type — 'Bubble' is the most common starting point."),
    (C_GREEN,  "Activate Capture Mode 🎯",      "Click the crosshair icon then click the target element on the page."),
    (C_BLUE,   "Write the Instruction Text",    "Keep it short:  'Click the Save button.'  Use plain language."),
    (C_TEAL,   "Set the Step Action",           "Choose what triggers the next step: On Click, On Change, etc."),
    (C_PURPLE, "Preview → Publish",             "Click Preview to test, fix issues, then Publish to your environment."),
]
cols = 4
for i, (col, title, desc) in enumerate(steps_swt):
    row, ci = divmod(i, cols)
    l = 0.35 + ci * 3.22
    t = 1.72 + row * 2.55
    rect(s, l, t, 3.05, 2.35, RGBColor(0x1A,0x2E,0x45), radius=10000)
    # Number badge
    badge = slide.shapes if False else s.shapes  # shorthand
    circ = s.shapes.add_shape(9, Inches(l+1.2), Inches(t+0.12), Inches(0.65), Inches(0.65))
    circ.fill.solid(); circ.fill.fore_color.rgb = col; circ.line.fill.background()
    text_box(s, str(i+1), l+1.22, t+0.16, 0.6, 0.5, size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    text_box(s, title, l+0.12, t+0.88, 2.82, 0.52, size=12, bold=True, color=col, align=PP_ALIGN.CENTER)
    text_box(s, desc,  l+0.12, t+1.38, 2.82, 0.85, size=10.5, color=RGBColor(0xCC,0xDD,0xEE),
             align=PP_ALIGN.CENTER, wrap=True)

    # Arrow between steps in same row
    if ci < cols-1:
        text_box(s, "→", l+2.97, t+0.9, 0.32, 0.5, size=18, bold=True, color=col)

footer(s, 7)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 8 — SWT Step Types Reference
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "SWT Step Types — Quick Reference", "Module 3 — Choose the right step for each interaction", tag="Module 3", tag_color=C_TEAL)

step_types = [
    ("💬", "Bubble",    C_BLUE,   "Shows a tooltip balloon pointing at an element. Most common step type. Great for instructions."),
    ("👆", "Click",     C_PURPLE, "Waits for the user to click a specific element before advancing to the next step."),
    ("⌨️", "Type",      C_TEAL,   "Waits for the user to type in a field. Can validate minimum character count."),
    ("☑️", "Select",    C_ORANGE, "Used for dropdown menus — waits for a specific option to be selected."),
    ("🔔", "Alert",     C_RED,    "Pops up a modal message with OK/Cancel — use for warnings or confirmations."),
    ("📍", "Hotspot",   C_GREEN,  "A pulsing beacon that draws attention to an element without requiring action."),
    ("🖼️", "Frame",     C_BLUE,   "Highlights a section of the page with a spotlight frame — ideal for orientation."),
    ("⚡", "Auto-Step", C_PURPLE, "Automatically advances when a condition is detected (URL change, element appears)."),
]
for i, (icon, name, col, desc) in enumerate(step_types):
    row, ci = divmod(i, 4)
    l = 0.35 + ci * 3.22
    t = 1.73 + row * 2.55
    rect(s, l, t, 3.05, 2.35, C_WHITE, radius=10000)
    rect(s, l, t, 3.05, 0.55, col, radius=10000)
    rect(s, l, t+0.3, 3.05, 0.25, col)
    text_box(s, f"{icon}  {name}", l+0.1, t+0.08, 2.85, 0.42, size=15, bold=True, color=C_WHITE)
    text_box(s, desc, l+0.12, t+0.65, 2.82, 1.6, size=11, color=C_DGRAY, wrap=True)

footer(s, 8)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 9 — SmartTips
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "SmartTips", "Module 4 — Contextual help at the point of need", tag="Module 4", tag_color=C_ORANGE)

# Left — How to create
rect(s, 0.35, 1.72, 6.2, 5.4, C_WHITE, radius=8000)
text_box(s, "How to Create a SmartTip", 0.55, 1.82, 5.8, 0.48, size=18, bold=True, color=C_NAVY)
rect(s, 0.55, 2.3, 5.8, 0.05, C_MGRAY)

how_to = [
    ("①", C_ORANGE, "Click  '+ New Item'  → Select  'SmartTip'"),
    ("②", C_BLUE,   "Click the crosshair 🎯 and click the target field/element"),
    ("③", C_TEAL,   "Write your tip text (keep it under 30 words!)"),
    ("④", C_PURPLE, "Set Trigger:  Hover  (mouse over)  or  Focus  (tab into field)"),
    ("⑤", C_GREEN,  "Optionally add an image, link, or video"),
    ("⑥", C_ORANGE, "Click Save → Preview → Publish"),
]
for i, (num, col, text) in enumerate(how_to):
    t = 2.45 + i * 0.73
    rect(s, 0.5, t, 0.42, 0.42, col, radius=20000)
    text_box(s, num, 0.52, t+0.04, 0.38, 0.35, size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    text_box(s, text, 1.0, t+0.04, 5.3, 0.4, size=13, color=C_NAVY)

# Right — browser mock
mock_browser(s, 6.8, 1.72, 6.18, 5.4, highlight="smarttip")
rect(s, 6.85, 1.75, 6.08, 0.3, C_ORANGE)
text_box(s, "▶ SmartTip tooltip on Email field", 6.95, 1.77, 5.9, 0.26, size=10, bold=True, color=C_DARK)

# Best practice banner
rect(s, 0.35, 6.8, 12.63, 0.25, C_ORANGE, radius=6000)  # covered by footer but adds color
footer(s, 9)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 10 — ShoutOuts
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_DARK)
header(s, "ShoutOuts", "Module 4 — Announcements, banners & notifications", tag="Module 4", tag_color=C_ORANGE)

# Browser mock showing shoutout
mock_browser(s, 0.35, 1.72, 6.18, 5.4, highlight="shoutout")
rect(s, 0.4, 1.75, 6.08, 0.3, C_ORANGE)
text_box(s, "▶ ShoutOut banner at top of page", 0.5, 1.77, 5.9, 0.26, size=10, bold=True, color=C_DARK)

# Right side cards
layouts = [
    ("📢", "Banner",  C_ORANGE, "Horizontal bar at top or bottom of the page. Least intrusive — great for announcements."),
    ("🪟", "Modal",   C_BLUE,   "Center-screen overlay with close button. Best for important messages requiring user action."),
    ("🔵", "Beacon",  C_TEAL,   "A pulsing dot that expands to a card when clicked. Subtle, non-disruptive option."),
]
for i, (icon, name, col, desc) in enumerate(layouts):
    t = 1.72 + i * 1.42
    rect(s, 6.85, t, 6.1, 1.3, RGBColor(0x1A,0x2A,0x40), radius=8000)
    rect(s, 6.85, t, 0.55, 1.3, col, radius=8000)
    rect(s, 7.1, t, 0.32, 1.3, col)
    text_box(s, icon, 6.93, t+0.38, 0.42, 0.5, size=22)
    text_box(s, name, 7.52, t+0.12, 2.0, 0.38, size=15, bold=True, color=col)
    text_box(s, desc, 7.52, t+0.52, 5.28, 0.68, size=11, color=RGBColor(0xCC,0xDD,0xEE), wrap=True)

# How-to steps
rect(s, 6.85, 5.0, 6.1, 2.12, RGBColor(0x1A,0x2A,0x40), radius=8000)
text_box(s, "How to Create a ShoutOut", 7.05, 5.1, 5.8, 0.42, size=14, bold=True, color=C_ORANGE)
for i, step in enumerate([
    "① Click '+ New Item' → ShoutOut",
    "② Choose a Layout: Banner / Modal / Beacon",
    "③ Edit text, images, and CTA buttons",
    "④ Set Trigger: Page load / URL / Date range",
    "⑤ Save → Preview → Publish",
]):
    text_box(s, step, 7.05, 5.55+i*0.29, 5.8, 0.28, size=11, color=C_WHITE)

footer(s, 10)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 11 — Launchers
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "Launchers", "Module 5 — On-demand triggers for WalkMe content", tag="Module 5", tag_color=C_GREEN)

# Browser mock
mock_browser(s, 0.35, 1.72, 5.5, 5.4, highlight="launcher")
rect(s, 0.4, 1.75, 5.4, 0.3, C_GREEN)
text_box(s, "▶ Launcher button — bottom right corner", 0.5, 1.77, 5.3, 0.26, size=10, bold=True, color=C_DARK)

# Right — details
rect(s, 6.1, 1.72, 6.88, 5.4, C_WHITE, radius=8000)
text_box(s, "What is a Launcher?", 6.3, 1.82, 6.5, 0.45, size=18, bold=True, color=C_NAVY)
text_box(s, "A clickable icon or button users can click to start a Walk-Thru, open a Resource, or navigate to a URL — on demand, not forced.",
         6.3, 2.28, 6.5, 0.75, size=12, color=C_DGRAY, wrap=True)
rect(s, 6.3, 3.05, 6.5, 0.05, C_MGRAY)

text_box(s, "How to Create a Launcher", 6.3, 3.15, 6.5, 0.42, size=14, bold=True, color=C_GREEN)
for i, step in enumerate([
    "① Click '+ New Item' → Launcher",
    "② Choose an icon (WalkMe library or custom)",
    "③ Set what it launches: Walk-Thru / Resource / URL",
    "④ Anchor it to a page element or set it as floating",
    "⑤ Add a tooltip label (e.g., 'How do I submit?')",
    "⑥ Save → Preview → Publish",
]):
    t = 3.62 + i * 0.41
    col = C_GREEN if i % 2 == 0 else C_TEAL
    text_box(s, step, 6.3, t, 6.5, 0.38, size=12, color=C_NAVY)

callout_box(s, "✅", "Best Practice",
            "Label Launchers clearly using action language: 'How do I submit a report?' "
            "Place them near the task they support — not floating randomly on every page.",
            6.3, 6.1, 6.55, 1.0, bg=RGBColor(0xE8,0xFB,0xF4), accent=C_GREEN)

footer(s, 11)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 12 — Resources
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "Resources & Resource Center", "Module 5 — Surface help content inside your application", tag="Module 5", tag_color=C_GREEN)

res_types = [
    ("📄", "Article",  C_BLUE,   "A text / rich-media help article authored directly inside WalkMe or pulled from a URL."),
    ("🎬", "Video",    C_PURPLE, "Embed a YouTube, Vimeo, or hosted video clip for visual learners."),
    ("🔗", "Link",     C_TEAL,   "A URL to any external page — knowledge base, policy doc, external training."),
    ("🚶", "Walk-Thru",C_ORANGE, "Launch an SWT directly from the Resource Center for hands-on guidance."),
]
for i, (icon, name, col, desc) in enumerate(res_types):
    l = 0.35 + i * 3.22
    rect(s, l, 1.72, 3.05, 1.75, C_WHITE, radius=10000)
    rect(s, l, 1.72, 3.05, 0.62, col, radius=10000)
    rect(s, l, 2.1, 3.05, 0.24, col)
    text_box(s, f"{icon}  {name}", l+0.12, 1.8, 2.82, 0.42, size=15, bold=True, color=C_WHITE)
    text_box(s, desc, l+0.12, 2.45, 2.82, 0.9, size=11, color=C_DGRAY, wrap=True)

# Resource center diagram
rect(s, 0.35, 3.65, 12.63, 3.47, C_WHITE, radius=8000)
text_box(s, "📚  Resource Center — The In-App Help Widget", 0.55, 3.75, 8.0, 0.48, size=16, bold=True, color=C_NAVY)

# Mock resource center widget
rect(s, 0.55, 4.3, 3.8, 2.6, RGBColor(0x1E,0x1E,0x2E), radius=10000)
rect(s, 0.55, 4.3, 3.8, 0.45, C_BLUE, radius=10000)
rect(s, 0.55, 4.55, 3.8, 0.2, C_BLUE)
text_box(s, "🔍  Search help...", 0.65, 4.33, 3.5, 0.38, size=11, color=C_WHITE)
for j, item in enumerate(["📄 Submit an Expense Report","🎬 Getting Started Video","🚶 Create a New Account","🔗 HR Policy Handbook"]):
    ty = 4.84 + j * 0.51
    rect(s, 0.65, ty, 3.6, 0.43, RGBColor(0x25,0x35,0x50), radius=6000)
    text_box(s, item, 0.72, ty+0.07, 3.45, 0.3, size=10, color=C_WHITE)

text_box(s, "How to create a Resource Center:", 4.6, 4.3, 4.5, 0.4, size=13, bold=True, color=C_NAVY)
for i, step in enumerate([
    "① Editor → Widgets → Resource Center",
    "② Click '+ Add Item' → choose type",
    "③ Organize into categories",
    "④ Set the Widget trigger (Launcher / auto-show)",
    "⑤ Save → Publish",
]):
    text_box(s, step, 4.6, 4.75+i*0.47, 4.5, 0.42, size=12, color=C_DGRAY)

callout_box(s, "💡", "Pro Tip",
            "Name Resources with search-friendly titles: 'How to submit an expense' "
            "rather than just 'Expense Form'.",
            9.2, 4.3, 3.6, 2.6, bg=C_LGRAY, accent=C_TEAL)

footer(s, 12)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 13 — Publishing & Testing
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_DARK)
header(s, "Publishing & Testing", "Module 6 — From build to live", tag="Module 6", tag_color=C_BLUE)

pub_steps = [
    (C_BLUE,   "💾", "Save (Draft)",         "Content auto-saves as Draft. Saved ≠ Published. Users cannot see Drafts."),
    (C_PURPLE, "👁️", "Preview",              "Click Preview and walk through the flow as a real user. Fix any broken steps."),
    (C_TEAL,   "🌐", "Publish to Test",      "Push to Test environment. Share with QA teammates using the WalkMe test extension."),
    (C_ORANGE, "🔍", "User Acceptance Test", "Open incognito, navigate to the live app, and experience the flow as an end user would."),
    (C_GREEN,  "🚀", "Publish to Production","Green-light it! End users will see your content on their next page load."),
    (C_BLUE,   "📊", "Monitor Analytics",    "Check WalkMe Insights: completion rate, drop-off steps, unique plays. Iterate!"),
]
# Horizontal pipeline
bar_y = 3.3
rect(s, 0.5, bar_y+0.5, 12.33, 0.08, RGBColor(0x33,0x55,0x88))

for i, (col, icon, title, desc) in enumerate(pub_steps):
    l = 0.35 + i * 2.16
    # Circle node
    circ = s.shapes.add_shape(9, Inches(l+0.72), Inches(bar_y+0.18), Inches(0.65), Inches(0.65))
    circ.fill.solid(); circ.fill.fore_color.rgb = col; circ.line.fill.background()
    text_box(s, icon, l+0.73, bar_y+0.22, 0.62, 0.55, size=20, align=PP_ALIGN.CENTER)
    # Card
    rect(s, l, bar_y+1.02, 2.05, 2.7, RGBColor(0x1A,0x2E,0x45), radius=8000)
    text_box(s, title, l+0.1, bar_y+1.12, 1.85, 0.52, size=12, bold=True, color=col, align=PP_ALIGN.CENTER)
    text_box(s, desc,  l+0.1, bar_y+1.65, 1.85, 1.0, size=10, color=RGBColor(0xCC,0xDD,0xEE),
             wrap=True, align=PP_ALIGN.CENTER)

# Top tip row
rect(s, 0.35, 1.72, 12.63, 1.35, RGBColor(0x1A,0x2A,0x3A), radius=8000)
text_box(s, "⚠️  Golden Rules of Publishing", 0.6, 1.78, 5.0, 0.45, size=14, bold=True, color=C_YELLOW)
for i, rule in enumerate([
    "Always preview before you publish.",
    "Never publish straight to Production — test first.",
    "Check on mobile / different browsers if your users use them.",
]):
    text_box(s, f"  {['①','②','③'][i]}  {rule}", 0.55+i*4.22, 2.22, 4.1, 0.72, size=11, color=C_WHITE, wrap=True)

footer(s, 13)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 14 — Best Practices Cheat Sheet
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_LGRAY)
header(s, "Best Practices Cheat Sheet", "Module 6 — Build right from the start", tag="Module 6", tag_color=C_BLUE)

do_dont = [
    ("✅ DO", C_GREEN,  [
        "Write in plain language: 'Click Save'",
        "One task per Walk-Thru",
        "Preview before every publish",
        "Name items clearly & consistently",
        "Use segments to target the right users",
        "Review analytics monthly",
    ]),
    ("❌ DON'T", C_RED, [
        "Use jargon or acronyms users won't know",
        "Cram 15 steps into one flow",
        "Publish without testing in incognito",
        "Name items 'New Walk-Thru 1'",
        "Show everything to everyone always",
        "Ignore broken selectors after UI changes",
    ]),
]
for i, (title, col, items) in enumerate(do_dont):
    l = 0.35 + i * 6.5
    rect(s, l, 1.72, 6.2, 5.4, C_WHITE, radius=8000)
    rect(s, l, 1.72, 6.2, 0.6, col, radius=8000)
    rect(s, l, 2.1, 6.2, 0.22, col)
    text_box(s, title, l+0.2, 1.78, 5.8, 0.48, size=18, bold=True, color=C_WHITE)
    for j, item in enumerate(items):
        t = 2.45 + j * 0.77
        rect(s, l+0.15, t, 5.9, 0.65, C_LGRAY, radius=6000)
        text_box(s, item, l+0.35, t+0.1, 5.6, 0.48, size=13, color=C_NAVY)

footer(s, 14)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 15 — Hands-On Lab
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_DARK)
header(s, "Hands-On Lab", "Module 7 — Build it yourself on the WalkMe demo site!", tag="Module 7", tag_color=C_PURPLE)

# Timer / Goal
rect(s, 0.35, 1.72, 12.63, 0.82, C_PURPLE, radius=8000)
text_box(s, "⏱️  Time: 45 min build  +  15 min group debrief       🌐  Site: demo.walkme.com (provided by instructor)",
         0.6, 1.85, 12.2, 0.56, size=13, bold=True, color=C_WHITE)

lab_tasks = [
    (C_TEAL,   "Task 1 — Smart Walk-Thru",
     ["Create a Walk-Thru with at least 5 steps",
      "Include a Click step, a Type step, and a Bubble",
      "Use meaningful step text (action verbs!)",
      "Preview and fix any broken selectors"]),
    (C_ORANGE, "Task 2 — SmartTip",
     ["Add a SmartTip to a form field",
      "Set trigger to 'Focus' (fires when user tabs into field)",
      "Write a helpful 1-sentence tip",
      "Preview it by clicking into the field"]),
    (C_BLUE,   "Task 3 — Launcher",
     ["Create a Launcher that starts your Walk-Thru",
      "Choose an icon from the WalkMe library",
      "Add a tooltip label to the Launcher",
      "Anchor it near the relevant form or button"]),
    (C_GREEN,  "Task 4 — Publish & Demo",
     ["Publish all 3 items to Test environment",
      "Open incognito and walk your neighbor through it",
      "Note what worked and what to improve",
      "Be ready to share 1 lesson learned with the group"]),
]
for i, (col, title, tasks) in enumerate(lab_tasks):
    row, ci = divmod(i, 2)
    l = 0.35 + ci * 6.5
    t = 2.72 + row * 2.2
    rect(s, l, t, 6.2, 2.0, RGBColor(0x1A,0x2A,0x3E), radius=10000)
    rect(s, l, t, 6.2, 0.52, col, radius=10000)
    rect(s, l, t+0.3, 6.2, 0.22, col)
    text_box(s, title, l+0.2, t+0.08, 5.8, 0.42, size=14, bold=True, color=C_WHITE)
    for j, task in enumerate(tasks):
        text_box(s, f"  ☐  {task}", l+0.2, t+0.62+j*0.34, 5.8, 0.32, size=11, color=C_WHITE)

footer(s, 15)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 16 — Summary & Next Steps
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_DARK)
rect(s, 0, 0, 13.33, 1.5, C_NAVY)
rect(s, 0, 1.5, 13.33, 0.08, C_BLUE)
rect(s, 0, 0, 0.35, 7.5, C_TEAL)
text_box(s, "Course Summary & Next Steps", 0.55, 0.2, 12.4, 0.82, size=30, bold=True, color=C_WHITE)
text_box(s, "You've completed WalkMe Builder 1 — here's what you've learned and where to go from here.",
         0.55, 0.92, 12.0, 0.52, size=14, italic=True, color=RGBColor(0xAA,0xCC,0xFF))

# Learned column
rect(s, 0.35, 1.75, 6.0, 5.3, RGBColor(0x12,0x22,0x38), radius=8000)
rect(s, 0.35, 1.75, 6.0, 0.55, C_TEAL, radius=8000)
rect(s, 0.35, 2.1, 6.0, 0.2, C_TEAL)
text_box(s, "🎓  What You Mastered", 0.55, 1.82, 5.6, 0.42, size=15, bold=True, color=C_DARK)
for i, item in enumerate([
    "✔  Navigating the WalkMe Editor",
    "✔  Building Smart Walk-Thrus (5 step types)",
    "✔  Creating SmartTips & ShoutOuts",
    "✔  Setting up Launchers & Resources",
    "✔  Publishing to Test & Production",
    "✔  Reading analytics & iterating",
    "✔  Hands-on lab — real WalkMe content!",
]):
    rect(s, 0.5, 2.42+i*0.75, 5.7, 0.65, RGBColor(0x1E,0x33,0x50), radius=6000)
    text_box(s, item, 0.65, 2.49+i*0.75, 5.4, 0.5, size=13, color=C_WHITE)

# Next steps column
rect(s, 6.65, 1.75, 6.33, 5.3, RGBColor(0x12,0x22,0x38), radius=8000)
rect(s, 6.65, 1.75, 6.33, 0.55, C_BLUE, radius=8000)
rect(s, 6.65, 2.1, 6.33, 0.2, C_BLUE)
text_box(s, "🚀  Your Next Steps", 6.85, 1.82, 6.0, 0.42, size=15, bold=True, color=C_WHITE)
next_steps = [
    ("📌", C_TEAL,   "Complete the post-course assessment"),
    ("📌", C_ORANGE, "Build 1 real Walk-Thru in your app this week"),
    ("📌", C_PURPLE, "Enroll in WalkMe Builder 2"),
    ("📌", C_GREEN,  "Join the WalkMe Community (community.walkme.com)"),
    ("📌", C_BLUE,   "Bookmark WalkMe Help Center (help.walkme.com)"),
    ("📌", C_TEAL,   "Schedule 1:1 with your WalkMe admin"),
    ("📌", C_ORANGE, "Share your first Walk-Thru with your team!"),
]
for i, (icon, col, step) in enumerate(next_steps):
    rect(s, 6.8, 2.42+i*0.75, 6.0, 0.65, RGBColor(0x1E,0x33,0x50), radius=6000)
    rect(s, 6.8, 2.42+i*0.75, 0.08, 0.65, col)
    text_box(s, f"{icon}  {step}", 6.95, 2.49+i*0.75, 5.75, 0.5, size=12, color=C_WHITE)

footer(s, 16)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 17 — Q&A / Thank You
# ─────────────────────────────────────────────────────────────────────────────
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, C_DARK)

# Decorative circles
for cx, cy, sz, col in [
    (10.5, -0.5, 5.0, RGBColor(0x00,0x40,0x99)),
    (11.5, 5.5, 3.5, RGBColor(0x00,0x28,0x66)),
    (-0.5, 6.5, 2.5, RGBColor(0x00,0x50,0xAA)),
]:
    c = s.shapes.add_shape(9, Inches(cx), Inches(cy), Inches(sz), Inches(sz))
    c.fill.solid(); c.fill.fore_color.rgb = col; c.line.fill.background()

rect(s, 0, 0, 0.5, 7.5, C_BLUE)
rect(s, 0.5, 0, 0.1, 7.5, C_TEAL)

text_box(s, "Q", 1.2, 0.6, 2.5, 2.5, size=160, bold=True, color=RGBColor(0x00,0x4A,0xCC), font="Segoe UI Black")
text_box(s, "&A", 3.3, 1.5, 2.5, 1.8, size=90, bold=True, color=C_TEAL, font="Segoe UI Black")

text_box(s, "Questions?", 1.2, 3.6, 7.0, 0.85, size=40, bold=True, color=C_WHITE)
text_box(s, "Raise your hand, type in the chat, or ask your instructor!",
         1.2, 4.45, 7.5, 0.65, size=18, color=RGBColor(0xAA,0xCC,0xFF))

rect(s, 1.2, 5.3, 7.8, 0.06, C_TEAL)

res = [
    ("📖", "Help Center",  "help.walkme.com"),
    ("🌐", "Community",    "community.walkme.com"),
    ("🎓", "WalkMe U",     "university.walkme.com"),
]
for i, (icon, label, url) in enumerate(res):
    l = 1.2 + i * 2.65
    text_box(s, icon, l, 5.5, 0.45, 0.5, size=22)
    text_box(s, label, l+0.48, 5.5, 2.1, 0.3, size=13, bold=True, color=C_TEAL)
    text_box(s, url, l+0.48, 5.8, 2.1, 0.3, size=11, italic=True, color=RGBColor(0x99,0xBB,0xDD))

rect(s, 0, 7.1, 13.33, 0.4, RGBColor(0x06,0x0F,0x1A))
text_box(s, "Thank you for attending WalkMe Builder 1!  |  See you in Builder 2!", 0.3, 7.13, 12.5, 0.3,
         size=12, color=C_TEAL, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────────────────────────────────────
out = "/home/user/nishant/WalkMe_Builder1_ILT_v2.pptx"
prs.save(out)
print(f"✅  Saved: {out}  ({len(prs.slides)} slides)")
