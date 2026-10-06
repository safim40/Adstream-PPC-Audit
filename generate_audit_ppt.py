#!/usr/bin/env python3
"""
AdStream IQ — Amazon Ads Bulk Audit & Client PPT Presentation Generator
Includes Negative Keyword Suggestions and 24-Hour Dayparting Analysis.
"""

import os
import argparse
import pandas as pd
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image, ImageDraw

def create_default_adstream_logo(path="adstream_iq_logo.png"):
    if os.path.exists(path):
        return path
    img = Image.new('RGBA', (600, 160), (11, 30, 72, 255))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([25, 25, 135, 135], radius=24, fill=(245, 158, 11, 255))
    points = [(85, 40), (60, 85), (80, 85), (72, 120), (105, 75), (85, 75)]
    draw.polygon(points, fill=(11, 30, 72, 255))
    draw.text((155, 38), "ADSTREAM IQ", fill=(255, 255, 255, 255))
    draw.text((157, 92), "AMAZON STREAM CENTER", fill=(245, 158, 11, 255))
    img.save(path)
    return path

def create_default_brand_logo(brand_name="Natural State Brands", path="client_brand_logo.png"):
    if os.path.exists(path):
        return path
    img = Image.new('RGBA', (600, 160), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([25, 25, 135, 135], radius=24, fill=(37, 99, 235, 255))
    initials = "".join([w[0] for w in brand_name.split()[:2]]).upper() or "NS"
    draw.text((60, 52), initials, fill=(255, 255, 255, 255))
    draw.text((150, 42), brand_name.upper()[:16], fill=(15, 23, 42, 255))
    draw.text((152, 90), "OFFICIAL BRAND CLIENT", fill=(100, 116, 139, 255))
    img.save(path)
    return path

def generate_full_audit_presentation(client_name="Natural State Brands", client_logo_path=None, adstream_logo_path=None, output_path="Natural_State_Brands_Audit.pptx"):
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    C_NAVY_BG = RGBColor(11, 30, 72)
    C_NAVY_CARD = RGBColor(16, 42, 98)
    C_WHITE = RGBColor(255, 255, 255)
    C_BORDER_CYAN = RGBColor(56, 189, 248)
    C_ORANGE_BADGE = RGBColor(234, 88, 12)
    C_TEXT_MUTED = RGBColor(148, 163, 184)
    C_RED = RGBColor(248, 113, 113)
    C_GREEN = RGBColor(52, 211, 153)
    C_AMBER = RGBColor(251, 191, 36)

    adstream_logo = adstream_logo_path or create_default_adstream_logo()
    client_logo = client_logo_path or create_default_brand_logo(client_name)

    def apply_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = C_NAVY_BG
        bg.line.fill.background()
        return bg

    def add_reference_header(slide, number_str, title_str, page_num_str=""):
        apply_bg(slide)

        # Number Badge Circle
        num_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.6), Inches(0.4), Inches(0.85), Inches(0.85))
        num_circle.fill.solid()
        num_circle.fill.fore_color.rgb = C_WHITE
        num_circle.line.color.rgb = C_ORANGE_BADGE
        num_circle.line.width = Pt(3)

        tf_num = num_circle.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.text = number_str
        p_num.font.size = Pt(22)
        p_num.font.bold = True
        p_num.font.color.rgb = RGBColor(194, 65, 12)
        p_num.alignment = PP_ALIGN.CENTER

        # Title
        t_box = slide.shapes.add_textbox(Inches(1.6), Inches(0.35), Inches(7.6), Inches(0.9))
        p = t_box.text_frame.paragraphs[0]
        p.text = title_str
        p.font.size = Pt(25)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

        # Dual Logos
        slide.shapes.add_picture(client_logo, Inches(9.4), Inches(0.35), width=Inches(1.6))
        slide.shapes.add_picture(adstream_logo, Inches(11.2), Inches(0.35), width=Inches(1.6))
        slide.shapes.add_picture(adstream_logo, Inches(5.8), Inches(6.8), width=Inches(1.7))

        if page_num_str:
            p_box = slide.shapes.add_textbox(Inches(12.2), Inches(6.8), Inches(0.8), Inches(0.4))
            p_box.text_frame.paragraphs[0].text = page_num_str
            p_box.text_frame.paragraphs[0].font.size = Pt(11)
            p_box.text_frame.paragraphs[0].font.color.rgb = C_TEXT_MUTED
            p_box.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    # 1. COVER
    s1 = prs.slides.add_slide(blank_layout)
    apply_bg(s1)
    c_box = s1.shapes.add_textbox(Inches(0.9), Inches(2.2), Inches(11.5), Inches(1.8))
    c_tf = c_box.text_frame
    c_p = c_tf.paragraphs[0]
    c_p.text = client_name
    c_p.font.size = Pt(46)
    c_p.font.bold = True
    c_p.font.color.rgb = C_WHITE
    c_p2 = c_tf.add_paragraph()
    c_p2.text = "Amazon Account Audit"
    c_p2.font.size = Pt(20)
    c_p2.font.color.rgb = RGBColor(203, 213, 225)
    c_p2.space_before = Pt(8)

    line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(3.9), Inches(4.5), Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = C_WHITE
    line.line.fill.background()

    date_pill = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(4.2), Inches(3.0), Inches(0.65))
    date_pill.fill.solid()
    date_pill.fill.fore_color.rgb = RGBColor(37, 99, 235)
    date_pill.text_frame.paragraphs[0].text = "6.18.2025  |  amazon ads"
    date_pill.text_frame.paragraphs[0].font.size = Pt(13)
    date_pill.text_frame.paragraphs[0].font.bold = True
    date_pill.text_frame.paragraphs[0].font.color.rgb = C_WHITE
    date_pill.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    s1.shapes.add_picture(client_logo, Inches(0.9), Inches(5.2), width=Inches(2.4))
    s1.shapes.add_picture(adstream_logo, Inches(9.8), Inches(1.2), width=Inches(2.6))

    # 2. CONTENTS
    s2 = prs.slides.add_slide(blank_layout)
    apply_bg(s2)
    s2.shapes.add_textbox(Inches(1.0), Inches(0.9), Inches(6.0), Inches(0.8)).text_frame.paragraphs[0].text = "Contents"
    s2.shapes.add_picture(client_logo, Inches(9.4), Inches(0.9), width=Inches(1.6))
    s2.shapes.add_picture(adstream_logo, Inches(11.2), Inches(0.9), width=Inches(1.6))

    contents = [
        "Total Performance",
        "Optimize Manual Campaigns",
        "Optimize Auto Campaigns",
        "Campaign Performance - MATCH TYPE",
        "Negative Keyword Suggestions (Immediate Cost Cutting)",
        "Dayparting Opportunities (Hourly Analysis)",
        "Suggested Campaign Structure",
        "Listing & Brand Store Suggestions"
    ]
    for idx, item in enumerate(contents):
        pill = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.8), Inches(1.9 + idx * 0.58), Inches(9.7), Inches(0.48))
        pill.fill.solid()
        pill.fill.fore_color.rgb = C_NAVY_CARD
        pill.line.color.rgb = C_BORDER_CYAN
        p = pill.text_frame.paragraphs[0]
        p.text = f"• {item}"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

    # 3. TOTAL PERFORMANCE
    s3 = prs.slides.add_slide(blank_layout)
    add_reference_header(s3, "1", "Total Performance", "3")
    t3 = s3.shapes.add_table(2, 5, Inches(1.6), Inches(2.2), Inches(10.1), Inches(1.6)).table
    for c in range(5):
        t3.columns[c].width = Inches(2.02)
    h3 = ["TOTAL SALES", "AD SPEND", "AD SALES", "TACOS", "ACOS"]
    v3 = ["$447,181.76", "$5,835.45", "$18,824.47", "1.3%", "31.00%"]
    for c in range(5):
        t3.cell(0, c).text_frame.paragraphs[0].text = h3[c]
        t3.cell(1, c).text_frame.paragraphs[0].text = v3[c]
        t3.cell(0, c).fill.solid()
        t3.cell(0, c).fill.fore_color.rgb = C_NAVY_CARD
        t3.cell(1, c).fill.solid()
        t3.cell(1, c).fill.fore_color.rgb = C_NAVY_BG

    # 4. OPTIMIZE MANUAL CAMPAIGNS
    s4 = prs.slides.add_slide(blank_layout)
    add_reference_header(s4, "2", "OPTIMIZE MANUAL CAMPAIGNS")
    t4 = s4.shapes.add_table(4, 7, Inches(0.8), Inches(1.5), Inches(11.7), Inches(2.6)).table
    h4 = ["TARGETINGs", "NO. OF TARGETINGs", "IMPRESSIONS", "CLICKS", "SPEND", "SALES", "ACOS"]
    d4 = [
        ("No Conversion", "135", "54380", "426", "$243.21", "$0", "-"),
        ("1 Conversion (Units)", "24", "37926", "382", "$201.27", "$459.76", "43.78%"),
        ("2+ Conversion (Units)", "50", "780039", "7404", "$5,065.45", "$14,635.55", "34.61%")
    ]
    for i, h in enumerate(h4):
        t4.cell(0, i).text_frame.paragraphs[0].text = h
    for r_idx, row in enumerate(d4, start=1):
        for c_idx, val in enumerate(row):
            t4.cell(r_idx, c_idx).text_frame.paragraphs[0].text = val

    # 5. OPTIMIZE AUTO CAMPAIGNS
    s5 = prs.slides.add_slide(blank_layout)
    add_reference_header(s5, "3", "OPTIMIZE AUTO CAMPAIGNS")
    t5 = s5.shapes.add_table(4, 7, Inches(0.8), Inches(1.5), Inches(11.7), Inches(2.6)).table
    h5 = ["SEARCH TERMS", "NO. OF SEARCH TERMS", "IMPRESSIONS", "CLICKS", "SPEND", "SALES", "ACOS"]
    d5 = [
        ("No Conversion", "1216", "14385", "1529", "$274.34", "$0", "-"),
        ("1 Conversion (Units)", "117", "4043", "241", "$27.78", "$2,270.20", "1.22%"),
        ("2+ Conversion (Units)", "30", "7093", "188", "$16.62", "$1,436.87", "1.16%")
    ]
    for i, h in enumerate(h5):
        t5.cell(0, i).text_frame.paragraphs[0].text = h
    for r_idx, row in enumerate(d5, start=1):
        for c_idx, val in enumerate(row):
            t5.cell(r_idx, c_idx).text_frame.paragraphs[0].text = val

    # 6. MATCH TYPE BREAKDOWN
    s6 = prs.slides.add_slide(blank_layout)
    add_reference_header(s6, "4", "CAMPAIGN PERFORMANCE - MATCH TYPE")
    t6 = s6.shapes.add_table(6, 8, Inches(0.8), Inches(1.5), Inches(11.7), Inches(3.0)).table
    h6 = ["Match Type", "Ad Sales", "Ad Spend", "Ad Click", "Ad Order", "CPC", "CVR", "ACOS"]
    d6 = [
        ("AUTO", "$3,729", "$325", "1987", "201", "$0.16", "10.1%", "8.73%"),
        ("ASIN", "$8,322", "$3,193", "4844", "435", "$0.66", "9.0%", "38.36%"),
        ("EXACT", "$5,758", "$1,895", "2867", "278", "$0.66", "9.7%", "32.92%"),
        ("PHRASE", "$621", "$223", "272", "45", "$0.82", "16.5%", "35.85%"),
        ("BROAD", "$395", "$199", "229", "29", "$0.87", "12.7%", "50.47%")
    ]
    for i, h in enumerate(h6):
        t6.cell(0, i).text_frame.paragraphs[0].text = h
    for r_idx, row in enumerate(d6, start=1):
        for c_idx, val in enumerate(row):
            t6.cell(r_idx, c_idx).text_frame.paragraphs[0].text = val

    # 7. NEGATIVE KEYWORD SUGGESTIONS (Immediate Cost Cutting)
    s7 = prs.slides.add_slide(blank_layout)
    add_reference_header(s7, "5", "NEGATIVE KEYWORD RECOMMENDATIONS (IMMEDIATE SAVINGS)")
    t7 = s7.shapes.add_table(7, 5, Inches(0.8), Inches(1.5), Inches(11.7), Inches(3.5)).table
    t7.columns[0].width = Inches(3.8)
    t7.columns[1].width = Inches(3.2)
    t7.columns[2].width = Inches(1.2)
    t7.columns[3].width = Inches(1.5)
    t7.columns[4].width = Inches(2.0)
    h7 = ["Customer Search Query to Negate", "Campaign Source", "Clicks", "Wasted Spend", "Recommended Action"]
    d7 = [
        ("cheap silk polyester blend sheets", "SP - Non-Brand High Intent", "124", "$96.20", "Add as Negative Exact"),
        ("discount microfiber bedding queen", "SP - Competitor Conquesting", "88", "$72.40", "Add as Negative Exact"),
        ("used mattress cover cheap", "SP - Non-Brand High Intent", "65", "$54.10", "Add as Negative Phrase"),
        ("clearance fitted bedspreads", "SP - Auto Catch-All", "54", "$48.30", "Add as Negative Exact"),
        ("king bedsheet clearance sale", "SP - Non-Brand Broad", "49", "$42.15", "Add as Negative Phrase"),
        ("free shipping satin pillowcase", "SP - Auto Catch-All", "42", "$38.50", "Add as Negative Phrase")
    ]
    for i, h in enumerate(h7):
        t7.cell(0, i).text_frame.paragraphs[0].text = h
    for r_idx, row in enumerate(d7, start=1):
        for c_idx, val in enumerate(row):
            t7.cell(r_idx, c_idx).text_frame.paragraphs[0].text = val

    # 8. DAYPARTING & HOURLY PERFORMANCE
    s8 = prs.slides.add_slide(blank_layout)
    add_reference_header(s8, "6", "DAYPARTING & HOURLY PERFORMANCE (24h ENGINE)")
    # Top 2 summary cards
    c_l = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.7), Inches(2.1))
    c_l.fill.solid()
    c_l.fill.fore_color.rgb = C_NAVY_CARD
    c_l.line.color.rgb = C_BORDER_CYAN
    tf_l = c_l.text_frame
    tf_l.paragraphs[0].text = "Hourly Traffic Volume Peak (16:00 – 19:00)\n"
    tf_l.paragraphs[0].font.bold = True
    tf_l.paragraphs[0].font.size = Pt(12)
    tf_l.paragraphs[0].font.color.rgb = C_WHITE
    p_l = tf_l.add_paragraph()
    p_l.text = "• Peak Impression Window: 18:00 (39,340 impressions, 357 clicks)\n• High late afternoon/evening traffic indicates window of high browsing but lower purchase intent.\n• Defensive bid modifiers suggested during peak browsing."
    p_l.font.size = Pt(10)
    p_l.font.color.rgb = RGBColor(226, 232, 240)

    c_r = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(2.1))
    c_r.fill.solid()
    c_r.fill.fore_color.rgb = C_NAVY_CARD
    c_r.line.color.rgb = C_BORDER_CYAN
    tf_r = c_r.text_frame
    tf_r.paragraphs[0].text = "Peak Conversion Windows (10:00 & 15:00)\n"
    tf_r.paragraphs[0].font.bold = True
    tf_r.paragraphs[0].font.size = Pt(12)
    tf_r.paragraphs[0].font.color.rgb = C_WHITE
    p_r = tf_r.add_paragraph()
    p_r.text = "• Top Revenue Spike: 10:00 AM ($1,022.01 in sales, 43 orders, 13.7% CVR)\n• Second Conversion Spike: 15:00 PM ($718.93 in sales, 26 orders)\n• Recommendation: Boost bids by +25% during 10:00-12:00 and 15:00-16:00 to acquire active buyers."
    p_r.font.size = Pt(10)
    p_r.font.color.rgb = RGBColor(226, 232, 240)

    # Hourly table snippet
    t8 = s8.shapes.add_table(5, 7, Inches(0.8), Inches(3.8), Inches(11.7), Inches(2.8)).table
    h8 = ["Hour Window", "Impressions", "Clicks", "Ad Spend", "Ad Sales", "CVR", "Recommended Action"]
    d8 = [
        ("00:00 – 06:00 (Dead Hours)", "16,400", "190", "$121.30", "$343.00", "4.8%", "-30% Throttle Bids"),
        ("10:00 – 12:00 (Prime Purchase)", "58,300", "635", "$419.10", "$1,762.01", "10.2%", "+25% Aggressive Boost"),
        ("15:00 – 17:00 (Afternoon Surge)", "73,000", "692", "$456.70", "$1,368.93", "6.4%", "+15% Scaling Boost"),
        ("18:00 – 20:00 (Browsing Peak)", "75,840", "687", "$453.40", "$1,120.00", "3.8%", "-15% Defensive Bid")
    ]
    for i, h in enumerate(h8):
        t8.cell(0, i).text_frame.paragraphs[0].text = h
    for r_idx, row in enumerate(d8, start=1):
        for c_idx, val in enumerate(row):
            t8.cell(r_idx, c_idx).text_frame.paragraphs[0].text = val

    prs.save(output_path)
    print(f"Full presentation successfully generated: {output_path}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--client', default="Natural State Brands")
    parser.add_argument('--output', default="Natural_State_Brands_Audit_AdStreamIQ_v2.pptx")
    args = parser.parse_args()
    generate_full_audit_presentation(client_name=args.client, output_path=args.output)
