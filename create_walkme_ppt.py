from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree

# WalkMe Brand Colors
WALKME_BLUE = RGBColor(0x00, 0x6E, 0xFF)
WALKME_DARK = RGBColor(0x1A, 0x1A, 0x2E)
WALKME_LIGHT_BLUE = RGBColor(0xE8, 0xF1, 0xFF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x5A, 0x5A, 0x5A)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
ACCENT = RGBColor(0x00, 0xC8, 0xAA)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]


def add_rect(slide, left, top, width, height, fill_color, alpha=None):
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    return shape


def add_text(slide, text, left, top, width, height, font_size=18, bold=False,
             color=None, align=PP_ALIGN.LEFT, italic=False, font_name="Calibri"):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font_name
    if color:
        run.font.color.rgb = color
    return txBox


def add_bullet_slide(slide, title, bullets, subtitle=None):
    # Header bar
    add_rect(slide, 0, 0, 13.33, 1.3, WALKME_BLUE)
    add_rect(slide, 0, 1.3, 13.33, 0.08, ACCENT)
    # Title
    add_text(slide, title, 0.4, 0.18, 12, 1.0, font_size=32, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_text(slide, subtitle, 0.4, 0.72, 12, 0.5, font_size=16, color=RGBColor(0xCC, 0xE5, 0xFF), align=PP_ALIGN.LEFT)
    # Background
    add_rect(slide, 0, 1.38, 13.33, 6.12, LIGHT_GRAY)
    # Content card
    add_rect(slide, 0.35, 1.6, 12.63, 5.65, WHITE)
    y = 1.85
    for bullet in bullets:
        if isinstance(bullet, dict):
            icon = bullet.get("icon", "•")
            text = bullet.get("text", "")
            sub = bullet.get("sub", None)
            add_text(slide, icon, 0.5, y, 0.4, 0.5, font_size=20, bold=True, color=WALKME_BLUE)
            add_text(slide, text, 0.95, y, 11.8, 0.5, font_size=20, bold=bullet.get("bold", False), color=WALKME_DARK)
            y += 0.52
            if sub:
                add_text(slide, sub, 1.1, y - 0.1, 11.5, 0.45, font_size=15, color=GRAY, italic=True)
                y += 0.38
        else:
            add_text(slide, "▸", 0.5, y, 0.4, 0.5, font_size=18, bold=True, color=WALKME_BLUE)
            add_text(slide, bullet, 0.95, y, 11.8, 0.5, font_size=19, color=WALKME_DARK)
            y += 0.52
    # Footer
    add_rect(slide, 0, 7.1, 13.33, 0.4, WALKME_DARK)
    add_text(slide, "WalkMe Builder 1  |  Instructor-Led Training  |  For Beginners", 0.3, 7.12, 10, 0.35,
             font_size=11, color=RGBColor(0xAA, 0xAA, 0xAA))
    add_text(slide, "CONFIDENTIAL", 10.5, 7.12, 2.5, 0.35, font_size=11, color=RGBColor(0x88, 0x88, 0x88),
             align=PP_ALIGN.RIGHT)


# ─────────────────────────────────────────────
# SLIDE 1 — Title / Cover
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, 13.33, 7.5, WALKME_DARK)
add_rect(slide, 0, 0, 5.5, 7.5, WALKME_BLUE)
add_rect(slide, 5.5, 3.4, 7.83, 0.06, ACCENT)

add_text(slide, "WalkMe", 0.5, 1.2, 4.8, 1.0, font_size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(slide, "Builder 1", 0.5, 2.1, 4.8, 0.9, font_size=44, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
add_text(slide, "Instructor-Led Training", 0.5, 3.05, 4.8, 0.6, font_size=20, color=WHITE, align=PP_ALIGN.CENTER)
add_text(slide, "For Beginners", 0.5, 3.55, 4.8, 0.5, font_size=17, italic=True, color=RGBColor(0xCC, 0xE5, 0xFF), align=PP_ALIGN.CENTER)
add_rect(slide, 0.5, 4.2, 4.3, 0.04, ACCENT)

add_text(slide, "Course Overview", 5.9, 1.3, 7.0, 0.6, font_size=24, bold=True, color=ACCENT)
lines = [
    "Duration: 1 Day (8 hours)",
    "Format: Instructor-Led (ILT)",
    "Level: Beginner",
    "",
    "You will learn to:",
    "  ✔  Navigate the WalkMe Editor",
    "  ✔  Build Smart Walk-Thrus",
    "  ✔  Create SmartTips & ShoutOuts",
    "  ✔  Use Launchers & Resources",
    "  ✔  Publish & test content",
]
y = 2.05
for line in lines:
    bold = line.startswith("You will") or line.startswith("Duration") or line.startswith("Format") or line.startswith("Level")
    col = ACCENT if line.startswith("You") else (WHITE if line.startswith("  ✔") else RGBColor(0xCC, 0xE5, 0xFF))
    add_text(slide, line, 5.9, y, 7.0, 0.45, font_size=16, bold=bold, color=col)
    y += 0.44

add_rect(slide, 0, 7.1, 13.33, 0.4, RGBColor(0x0D, 0x0D, 0x1A))
add_text(slide, "WalkMe Builder 1  |  ILT  |  Beginner", 0.3, 7.13, 9, 0.3, font_size=11, color=RGBColor(0x99, 0x99, 0x99))


# ─────────────────────────────────────────────
# SLIDE 2 — Agenda
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Course Agenda", subtitle="What we'll cover today", bullets=[
    {"icon": "1", "text": "Welcome & Introduction to WalkMe", "bold": True, "sub": "Who uses WalkMe and why? Key terminology."},
    {"icon": "2", "text": "The WalkMe Editor — Your Command Center", "bold": True, "sub": "Navigating the editor, sidebar, and main menu."},
    {"icon": "3", "text": "Smart Walk-Thrus (SWT)", "bold": True, "sub": "Step-by-step guided flows — the core WalkMe content type."},
    {"icon": "4", "text": "SmartTips & ShoutOuts", "bold": True, "sub": "Contextual tooltips and announcement banners."},
    {"icon": "5", "text": "Launchers & Resources", "bold": True, "sub": "Trigger flows and surface help content on demand."},
    {"icon": "6", "text": "Publishing, Testing & Best Practices", "bold": True, "sub": "Preview, publish, and validate your content."},
    {"icon": "7", "text": "Hands-On Lab & Q&A", "bold": True, "sub": "Build your first Walk-Thru from scratch!"},
])


# ─────────────────────────────────────────────
# SLIDE 3 — What is WalkMe?
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "What is WalkMe?", subtitle="Module 1 — Introduction", bullets=[
    {"icon": "💡", "text": "WalkMe is a Digital Adoption Platform (DAP)", "bold": True,
     "sub": "It sits on top of any web application and guides users through tasks in real time."},
    {"icon": "🎯", "text": "Who uses WalkMe?", "bold": True,
     "sub": "HR teams, IT, Sales Ops, Customer Success, L&D — anyone who wants users to succeed in software."},
    {"icon": "🔧", "text": "What does WalkMe Builder do?", "bold": True,
     "sub": "The Builder (Editor) lets you create guided content — no coding required!"},
    {"icon": "🌐", "text": "How is it deployed?", "bold": True,
     "sub": "As a browser extension (for testing) or embedded script tag on your web application."},
    {"icon": "📦", "text": "Core content types you'll build today:", "bold": True,
     "sub": "Smart Walk-Thrus  |  SmartTips  |  ShoutOuts  |  Launchers  |  Resources"},
])


# ─────────────────────────────────────────────
# SLIDE 4 — Key Terminology
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Key Terminology", subtitle="Module 1 — Know the language before you build", bullets=[
    {"icon": "📖", "text": "Step  —  A single instruction in a Walk-Thru (click here, fill this field…)", "bold": False},
    {"icon": "📖", "text": "Flow  —  A sequence of steps that guides a user to complete a task", "bold": False},
    {"icon": "📖", "text": "Selector  —  The code 'fingerprint' WalkMe uses to identify an element on-screen", "bold": False},
    {"icon": "📖", "text": "Trigger  —  The condition or action that starts a WalkMe item (auto, click, URL…)", "bold": False},
    {"icon": "📖", "text": "Segment  —  A rule to show content only to specific users or in specific contexts", "bold": False},
    {"icon": "📖", "text": "Publish  —  Pushing your saved content live so end users can see it", "bold": False},
    {"icon": "📖", "text": "Environment  —  Separate spaces (Test vs. Production) for safe building", "bold": False},
])


# ─────────────────────────────────────────────
# SLIDE 5 — The WalkMe Editor
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "The WalkMe Editor", subtitle="Module 2 — Your Command Center", bullets=[
    {"icon": "🖥️", "text": "Access the Editor via the WalkMe browser extension", "bold": True,
     "sub": "Click the WalkMe icon in your browser toolbar → opens the floating editor panel."},
    {"icon": "📋", "text": "Left Sidebar — Content List", "bold": True,
     "sub": "All your Walk-Thrus, SmartTips, ShoutOuts, and Launchers live here."},
    {"icon": "➕", "text": "The '+' Button — Create New Content", "bold": True,
     "sub": "Use this to create any new item. Choose the content type from the dropdown."},
    {"icon": "🔍", "text": "Element Capture Mode", "bold": True,
     "sub": "Click the crosshair icon to let WalkMe capture a page element and build a selector."},
    {"icon": "⚙️", "text": "Settings Panel", "bold": True,
     "sub": "Each item has its own settings: name, trigger, segmentation, and display options."},
    {"icon": "👁️", "text": "Preview Mode", "bold": True,
     "sub": "Test your content in real time before publishing — always preview first!"},
])


# ─────────────────────────────────────────────
# SLIDE 6 — Smart Walk-Thrus Overview
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Smart Walk-Thrus (SWT)", subtitle="Module 3 — Step-by-step guided flows", bullets=[
    {"icon": "🚶", "text": "What is a Smart Walk-Thru?", "bold": True,
     "sub": "An interactive overlay that guides users step-by-step through a task on any webpage."},
    {"icon": "1️⃣", "text": "Step Types you can add:", "bold": True,
     "sub": "Bubble (tooltip), Click, Type, Select, Hotspot, Alert, Frame — each fits a different interaction."},
    {"icon": "🎯", "text": "Capturing Elements", "bold": True,
     "sub": "Use the crosshair to click any page element. WalkMe auto-generates a selector for it."},
    {"icon": "✏️", "text": "Writing Step Text", "bold": True,
     "sub": "Keep instructions short: 'Click the Save button.' Use action verbs. Avoid jargon."},
    {"icon": "🔀", "text": "Step Settings", "bold": True,
     "sub": "Action (what triggers moving to next step), placement (where bubble appears), and conditions."},
    {"icon": "✅", "text": "Best Practice", "bold": True,
     "sub": "Limit flows to 7–10 steps. Break long processes into multiple Walk-Thrus."},
])


# ─────────────────────────────────────────────
# SLIDE 7 — Building a Walk-Thru (Step by Step)
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Building Your First Walk-Thru", subtitle="Module 3 — Step-by-step process", bullets=[
    {"icon": "①", "text": "Click '+' in the Editor → Choose 'Smart Walk-Thru'", "bold": True},
    {"icon": "②", "text": "Give it a clear name  (e.g., 'Submit a New Expense Report')", "bold": True},
    {"icon": "③", "text": "Click '+ Add Step' → Select step type (usually 'Bubble')", "bold": True},
    {"icon": "④", "text": "Use the crosshair to capture the first element the user needs to interact with", "bold": True},
    {"icon": "⑤", "text": "Write the instruction text in the bubble editor", "bold": True},
    {"icon": "⑥", "text": "Set the Action (e.g., 'On Click' to auto-advance when user clicks)", "bold": True},
    {"icon": "⑦", "text": "Repeat ③–⑥ for each step → then click Preview to test!", "bold": True},
])


# ─────────────────────────────────────────────
# SLIDE 8 — SmartTips
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "SmartTips", subtitle="Module 4 — Contextual tooltips at the point of need", bullets=[
    {"icon": "💬", "text": "What is a SmartTip?", "bold": True,
     "sub": "A small tooltip that appears when a user hovers over or focuses on a specific UI element."},
    {"icon": "🎯", "text": "Best used for:", "bold": True,
     "sub": "Field-level help, definitions, policy reminders, warnings, and inline instructions."},
    {"icon": "🔧", "text": "How to create:", "bold": True,
     "sub": "Click '+' → SmartTip → Capture the element → Add your tip text → Set trigger (Hover/Focus)."},
    {"icon": "🖼️", "text": "You can add:", "bold": True,
     "sub": "Text, images, links, and even embedded videos inside a SmartTip balloon."},
    {"icon": "✅", "text": "Best Practice:", "bold": True,
     "sub": "Keep SmartTip text under 30 words. Link to longer documentation if needed."},
])


# ─────────────────────────────────────────────
# SLIDE 9 — ShoutOuts
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "ShoutOuts", subtitle="Module 4 — Announcements & Notifications", bullets=[
    {"icon": "📢", "text": "What is a ShoutOut?", "bold": True,
     "sub": "A full-screen or banner overlay used to announce updates, new features, or important messages."},
    {"icon": "📐", "text": "Layout options:", "bold": True,
     "sub": "Banner (top/bottom), Modal (center), or Beacon (dot that expands on click)."},
    {"icon": "🔧", "text": "How to create:", "bold": True,
     "sub": "Click '+' → ShoutOut → Choose template → Edit text, images, CTA buttons → Set trigger."},
    {"icon": "⏱️", "text": "Trigger Options:", "bold": True,
     "sub": "On page load, on URL match, scheduled date range, or after X days since last seen."},
    {"icon": "🚫", "text": "Avoid overuse:", "bold": True,
     "sub": "ShoutOuts interrupt workflows. Only use for high-priority, time-sensitive announcements."},
])


# ─────────────────────────────────────────────
# SLIDE 10 — Launchers
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Launchers", subtitle="Module 5 — On-demand triggers for WalkMe content", bullets=[
    {"icon": "🚀", "text": "What is a Launcher?", "bold": True,
     "sub": "A clickable button or icon that users can click to start a Walk-Thru or open a Resource on demand."},
    {"icon": "📍", "text": "Where do Launchers appear?", "bold": True,
     "sub": "Attached to a specific element on the page, or floating anywhere on screen."},
    {"icon": "🔧", "text": "How to create:", "bold": True,
     "sub": "Click '+' → Launcher → Choose an icon/style → Set what it launches (Walk-Thru, URL, Resource)."},
    {"icon": "🎨", "text": "Customization:", "bold": True,
     "sub": "Use your company icon, custom text, or WalkMe's built-in icon library."},
    {"icon": "✅", "text": "Best Practice:", "bold": True,
     "sub": "Place Launchers near the task they support. Label them clearly: 'How do I submit a report?'"},
])


# ─────────────────────────────────────────────
# SLIDE 11 — Resources
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Resources", subtitle="Module 5 — Surfacing help content inside your application", bullets=[
    {"icon": "📚", "text": "What is a Resource?", "bold": True,
     "sub": "A help article, video, PDF, or link surfaced inside your app via WalkMe's Resource panel."},
    {"icon": "🗂️", "text": "Resource Center:", "bold": True,
     "sub": "A searchable help widget users can open anytime — like an in-app knowledge base."},
    {"icon": "🔧", "text": "How to create:", "bold": True,
     "sub": "Click '+' → Resource → Choose type (Article / Video / Link) → Add content → Save."},
    {"icon": "🔗", "text": "Linking Resources to Launchers:", "bold": True,
     "sub": "Set a Launcher to open a Resource so users get help exactly when and where they need it."},
    {"icon": "✅", "text": "Best Practice:", "bold": True,
     "sub": "Organize Resources into categories. Keep titles searchable: 'How to submit an expense.'"},
])


# ─────────────────────────────────────────────
# SLIDE 12 — Publishing & Testing
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Publishing & Testing Your Content", subtitle="Module 6 — From build to live", bullets=[
    {"icon": "👁️", "text": "Step 1 — Preview", "bold": True,
     "sub": "Always click 'Preview' first. Walk through your own flow as an end user to catch errors."},
    {"icon": "💾", "text": "Step 2 — Save", "bold": True,
     "sub": "Content is saved as a Draft automatically. Saved ≠ Published — users can't see it yet."},
    {"icon": "🌐", "text": "Step 3 — Publish", "bold": True,
     "sub": "Click 'Publish' to push live. Choose environment: Test first, then Production."},
    {"icon": "🔄", "text": "Step 4 — Test in Production", "bold": True,
     "sub": "Open a private/incognito window and experience the flow as a real user would."},
    {"icon": "📊", "text": "Step 5 — Check Analytics", "bold": True,
     "sub": "WalkMe Insights shows completions, drop-offs, and engagement — use data to improve."},
    {"icon": "♻️", "text": "Iterate!", "bold": True,
     "sub": "Good content is never 'done.' Review analytics monthly and update as the app changes."},
])


# ─────────────────────────────────────────────
# SLIDE 13 — Best Practices
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Best Practices for Beginners", subtitle="Module 6 — Build right from the start", bullets=[
    {"icon": "✍️", "text": "Write for your audience — use plain language, avoid IT jargon", "bold": False},
    {"icon": "🎯", "text": "One task, one Walk-Thru — don't try to do everything in a single flow", "bold": False},
    {"icon": "📱", "text": "Test on the same browser your users use — selectors can vary", "bold": False},
    {"icon": "🏷️", "text": "Name everything clearly — future-you will thank present-you", "bold": False},
    {"icon": "🔄", "text": "Re-capture selectors when the app UI changes — don't let them break", "bold": False},
    {"icon": "📊", "text": "Monitor analytics weekly during the first month after launch", "bold": False},
    {"icon": "🤝", "text": "Collaborate — share your Editor access with SMEs who know the process", "bold": False},
])


# ─────────────────────────────────────────────
# SLIDE 14 — Hands-On Lab
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_bullet_slide(slide, "Hands-On Lab", subtitle="Module 7 — Build it yourself!", bullets=[
    {"icon": "🎯", "text": "Lab Goal:", "bold": True,
     "sub": "Build a complete guided flow on the WalkMe demo website from scratch."},
    {"icon": "①", "text": "Create a Smart Walk-Thru with at least 4 steps", "bold": False},
    {"icon": "②", "text": "Add one SmartTip to a form field on the page", "bold": False},
    {"icon": "③", "text": "Create a Launcher that triggers your Walk-Thru", "bold": False},
    {"icon": "④", "text": "Preview your content — fix any broken steps", "bold": False},
    {"icon": "⑤", "text": "Publish to Test environment and walk a partner through it", "bold": False},
    {"icon": "⏱️", "text": "Time: 45 minutes | Then: 15 min group debrief", "bold": True},
])


# ─────────────────────────────────────────────
# SLIDE 15 — Summary & Next Steps
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, 13.33, 7.5, WALKME_DARK)
add_rect(slide, 0, 0, 13.33, 1.4, WALKME_BLUE)
add_rect(slide, 0, 1.4, 13.33, 0.08, ACCENT)
add_text(slide, "Course Summary & Next Steps", 0.4, 0.2, 12.5, 1.0, font_size=32, bold=True, color=WHITE)

# Two columns
# Left — Summary
add_rect(slide, 0.3, 1.65, 6.0, 5.5, RGBColor(0x22, 0x22, 0x40))
add_text(slide, "What You Learned Today", 0.55, 1.8, 5.5, 0.5, font_size=18, bold=True, color=ACCENT)
learned = [
    "✔  Navigating the WalkMe Editor",
    "✔  Building Smart Walk-Thrus",
    "✔  Creating SmartTips & ShoutOuts",
    "✔  Using Launchers & Resources",
    "✔  Publishing and testing content",
    "✔  Best practices for beginners",
]
y = 2.4
for item in learned:
    add_text(slide, item, 0.6, y, 5.5, 0.45, font_size=16, color=WHITE)
    y += 0.48

# Right — Next Steps
add_rect(slide, 6.8, 1.65, 6.2, 5.5, RGBColor(0x22, 0x22, 0x40))
add_text(slide, "Your Next Steps", 7.05, 1.8, 5.7, 0.5, font_size=18, bold=True, color=ACCENT)
steps = [
    "📌  Complete the post-course assessment",
    "📌  Build 1 real Walk-Thru this week",
    "📌  Enroll in WalkMe Builder 2",
    "📌  Join the WalkMe Community forum",
    "📌  Bookmark WalkMe's Help Center",
    "📌  Schedule a 1:1 with your WalkMe admin",
]
y = 2.4
for step in steps:
    add_text(slide, step, 7.05, y, 5.7, 0.45, font_size=16, color=WHITE)
    y += 0.48

add_rect(slide, 0, 7.1, 13.33, 0.4, RGBColor(0x0D, 0x0D, 0x1A))
add_text(slide, "Thank you for attending WalkMe Builder 1!  🎉  Questions?  Raise your hand.", 0.4, 7.12, 12.5, 0.3,
         font_size=13, color=ACCENT, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────
# SLIDE 16 — Q&A / Thank You
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, 13.33, 7.5, WALKME_BLUE)
add_rect(slide, 0, 0, 13.33, 7.5, WALKME_DARK)
# Big circle decoration
add_rect(slide, 8.5, -1.5, 7, 7, RGBColor(0x00, 0x5A, 0xD4))

add_text(slide, "Questions &", 1.5, 1.8, 9, 1.2, font_size=64, bold=True, color=WHITE)
add_text(slide, "Answers", 1.5, 2.9, 9, 1.2, font_size=64, bold=True, color=ACCENT)
add_rect(slide, 1.5, 4.15, 4.5, 0.08, ACCENT)
add_text(slide, "Raise your hand or type in the chat!", 1.5, 4.35, 8, 0.55, font_size=22, color=WHITE)
add_text(slide, "WalkMe Help Center:  help.walkme.com", 1.5, 5.0, 8, 0.5, font_size=17, italic=True,
         color=RGBColor(0xCC, 0xE5, 0xFF))
add_text(slide, "WalkMe Community:  community.walkme.com", 1.5, 5.5, 8, 0.5, font_size=17, italic=True,
         color=RGBColor(0xCC, 0xE5, 0xFF))

add_rect(slide, 0, 7.1, 13.33, 0.4, RGBColor(0x0D, 0x0D, 0x1A))
add_text(slide, "WalkMe Builder 1  |  ILT  |  Beginner", 0.3, 7.12, 12.5, 0.3, font_size=11,
         color=RGBColor(0x99, 0x99, 0x99), align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────
output = "/home/user/nishant/WalkMe_Builder1_ILT_Beginners.pptx"
prs.save(output)
print(f"Saved: {output}  ({prs.slides.__len__()} slides)")
