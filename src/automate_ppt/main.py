"""
Korea-Nepal Relationship PowerPoint Creator
Uses win32com to control PowerPoint LIVE on screen.

Requirements:
    pip install pywin32

Run:
    python korea_nepal_ppt.py
"""

import win32com.client
import time
import os

# ── Helpers ──────────────────────────────────────────────────────────────────

def rgb(r, g, b):
    """Convert RGB to the integer format win32com/PowerPoint expects."""
    return r + (g * 256) + (b * 256 * 256)

def add_textbox(slide, left, top, width, height, text,
                font_name="Calibri", font_size=14, bold=False,
                color=(255, 255, 255), align=None, italic=False):
    """Add a text box to a slide and return the shape."""
    inches = 72  # EMUs per inch
    tb = slide.Shapes.AddTextbox(
        1,  # msoTextOrientationHorizontal
        left * inches, top * inches,
        width * inches, height * inches
    )

    tf = tb.TextFrame
    tf.WordWrap = True
    tf.AutoSize = 0  # ppAutoSizeNone

    # Set full text at once
    tf.TextRange.Text = text

    tr = tf.TextRange
    tr.Font.Name = font_name
    tr.Font.Size = font_size
    tr.Font.Bold = bold
    tr.Font.Italic = italic
    tr.Font.Color.RGB = rgb(*color)

    if align:
        # 1=left, 2=center, 3=right
        tr.ParagraphFormat.Alignment = align

    return tb

def add_shape_rect(slide, left, top, width, height, fill_color, line_visible=False):
    """Add a solid rectangle shape."""
    inches = 72
    # COM requires left/top/width/height all > 0
    left  = max(left,  0.01)
    top   = max(top,   0.01)
    width = max(width, 0.01)
    height= max(height,0.01)
    shape = slide.Shapes.AddShape(
        1,  # msoShapeRectangle
        left * inches, top * inches,
        width * inches, height * inches
    )
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = rgb(*fill_color)
    if not line_visible:
        shape.Line.Visible = False
    return shape

# ── Main script ───────────────────────────────────────────────────────────────

def build_presentation():
    print("🚀 Launching PowerPoint...")
    def get_or_create_presentation():

        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            print("✅ Attached to running PowerPoint.")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            # ppt_app.Visible = True
            # ppt_app.WindowState = 2
            print("🚀 Started new PowerPoint instance.")

        time.sleep(0.5)


        # Force foreground
        import win32gui, win32con, win32api

        hwnd = win32gui.FindWindow(None, ppt_app.Caption)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32api.keybd_event(0x12, 0, 0, 0)
            win32gui.SetForegroundWindow(hwnd)
            win32api.keybd_event(0x12, 0, win32con.KEYEVENTF_KEYUP, 0)
            
        target_name = "Presentation1"
        pres = None

        for p in ppt_app.Presentations:
            if p.Name.replace(".pptx", "") == target_name:
                pres = p
                print(f"📂 Using existing presentation: {p.Name}")
                break

        if pres is None:
            print("📄 Creating new presentation.")
            pres = ppt_app.Presentations.Add()
            # pres.SaveAs(f"{target_name}.pptx")

        return ppt_app, pres

    ppt_app, pres = get_or_create_presentation()

    ppt_app.Activate()

    # Slide dimensions (widescreen 13.33 x 7.5 inches)
    pres.PageSetup.SlideWidth  = 13.33 * 72
    pres.PageSetup.SlideHeight = 7.5  * 72

    time.sleep(0.5)

    # ── SLIDE 1: Title Slide ──────────────────────────────────────────────────
    print("📄 Building Slide 1: Title...")

    slide1 = pres.Slides.Add(1, 12)  # ppLayoutBlank
    slide1.FollowMasterBackground = False

    # Dark navy background
    slide1.Background.Fill.Solid()
    slide1.Background.Fill.ForeColor.RGB = rgb(13, 27, 62)

    # Left accent bar (Korean red)
    add_shape_rect(slide1, 0, 0, 0.35, 7.5, (200, 18, 46))

    # Right accent bar (Nepalese crimson)
    add_shape_rect(slide1, 12.98, 0, 0.35, 7.5, (169, 29, 58))

    # Decorative center horizontal line
    add_shape_rect(slide1, 0.6, 3.6, 12.1, 0.04, (255, 255, 255))

    # Flag emoji labels
    add_textbox(slide1, 1.5, 1.0, 4.0, 1.0,
                "🇰🇷  South Korea", font_size=20, bold=True,
                color=(255, 220, 0), align=2)

    add_textbox(slide1, 7.5, 1.0, 4.0, 1.0,
                "🇳🇵  Nepal", font_size=20, bold=True,
                color=(255, 220, 0), align=2)

    # Main title
    add_textbox(slide1, 0.6, 2.2, 12.1, 1.2,
                "A Partnership Across the Himalayas",
                font_name="Georgia", font_size=38, bold=True,
                color=(255, 255, 255), align=2)

    # Subtitle
    add_textbox(slide1, 0.6, 3.8, 12.1, 0.7,
                "Korea–Nepal Bilateral Relations: History, Trade & Future Cooperation",
                font_size=16, italic=True,
                color=(180, 200, 240), align=2)

    # Year tag
    add_textbox(slide1, 0.6, 6.8, 12.1, 0.5,
                "Established 1974  •  50+ Years of Diplomacy",
                font_size=12, color=(140, 160, 200), align=2)

    time.sleep(0.8)

    # ── SLIDE 2: Key Pillars ──────────────────────────────────────────────────
    print("📄 Building Slide 2: Key Pillars...")

    slide2 = pres.Slides.Add(2, 12)  # ppLayoutBlank
    slide2.FollowMasterBackground = False

    # Light cream background
    slide2.Background.Fill.Solid()
    slide2.Background.Fill.ForeColor.RGB = rgb(245, 245, 250)

    # Top header bar (dark navy)
    add_shape_rect(slide2, 0, 0, 13.33, 1.3, (13, 27, 62))

    # Slide title
    add_textbox(slide2, 0.4, 0.15, 12.5, 1.0,
                "Key Pillars of Korea–Nepal Relations",
                font_name="Georgia", font_size=28, bold=True,
                color=(255, 255, 255), align=1)

    # ── Four content cards ────────────────────────────────────────────────────
    cards = [
        {
            "icon": "🤝",
            "title": "Diplomatic Ties",
            "color": (200, 18, 46),      # Korean red
            "body": (
                "Diplomatic relations established in 1974. "
                "Korea maintains an embassy in Kathmandu. "
                "Both nations cooperate at the UN and ASEAN forums."
            ),
            "stat": "50+ Years",
        },
        {
            "icon": "💰",
            "title": "Trade & Investment",
            "color": (13, 100, 180),
            "body": (
                "Korea exports machinery, electronics & vehicles to Nepal. "
                "Two-way trade exceeds $100M annually. "
                "KOICA supports private-sector development projects."
            ),
            "stat": "$100M+ Trade",
        },
        {
            "icon": "🎓",
            "title": "Education & ODA",
            "color": (169, 29, 58),      # Nepalese crimson
            "body": (
                "Korea's ODA (KOICA) funds infrastructure, health & education. "
                "GKS Scholarships bring Nepali students to Korean universities. "
                "Technical training programs in IT and vocational skills."
            ),
            "stat": "1,000+ Scholars",
        },
        {
            "icon": "🌏",
            "title": "People & Culture",
            "color": (0, 120, 80),
            "body": (
                "~50,000 Nepali migrants work legally in South Korea via EPS. "
                "Growing Hallyu (K-pop/K-drama) fanbase in Nepal. "
                "Annual cultural exchange festivals in both countries."
            ),
            "stat": "~50,000 Workers",
        },
    ]

    card_positions = [
        (0.3,  1.5),
        (6.85, 1.5),
        (0.3,  4.3),
        (6.85, 4.3),
    ]

    for card, (cx, cy) in zip(cards, card_positions):
        # Card background
        add_shape_rect(slide2, cx, cy, 6.2, 2.6, (255, 255, 255))

        # Colored top strip
        add_shape_rect(slide2, cx, cy, 6.2, 0.08, card["color"])

        # Icon + title
        add_textbox(slide2, cx + 0.15, cy + 0.12, 4.5, 0.5,
                    f"{card['icon']}  {card['title']}",
                    font_size=15, bold=True,
                    color=(20, 20, 50), align=1)

        # Stat callout (right side)
        add_textbox(slide2, cx + 3.8, cy + 0.12, 2.2, 0.45,
                    card["stat"],
                    font_size=11, bold=True, italic=True,
                    color=card["color"], align=3)

        # Body text
        add_textbox(slide2, cx + 0.15, cy + 0.65, 5.9, 1.8,
                    card["body"],
                    font_size=12, color=(60, 60, 80), align=1)

    time.sleep(0.8)

    # ── Save ──────────────────────────────────────────────────────────────────
    # save_path = os.path.join(os.path.expanduser("~"), "Desktop", "Korea_Nepal_Relations.pptx")
    # print(f"\n💾 Saving to: {save_path}")
    # pres.SaveAs(save_path)
    # print("✅ Done! File saved to your Desktop.")
    print("\n💡 PowerPoint is still open so you can review and edit live.")


if __name__ == "__main__":
    build_presentation()
    