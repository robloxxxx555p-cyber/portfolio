#!/usr/bin/env python3
"""
Generate PowerPoint Presentation for SANZ GOKU Business Portfolio
Following Persona 5 Aesthetic and Exact 8-Slide Requirements
Updated: Brand changed to 'SANZ GOKU', Bot '@Sanzstrnokosmumerbot', Photo embedded.
"""

import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Initialize Presentation
prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Persona 5 Theme Palette
COLOR_BG_DARK = RGBColor(14, 14, 17)        # Pure dark ink
COLOR_BG_CARD = RGBColor(26, 26, 32)        # Card dark grey
COLOR_RED_PRIMARY = RGBColor(230, 0, 18)    # P5 Iconic Crimson Red
COLOR_RED_DEEP = RGBColor(163, 0, 12)       # P5 Deep Wine Red
COLOR_WHITE = RGBColor(246, 244, 239)       # P5 Off-white / Cream
COLOR_TEXT_MUTED = RGBColor(170, 170, 180)  # Light grey
COLOR_ACCENT_YELLOW = RGBColor(241, 224, 90)# P5 Gold / Yellow star
COLOR_BORDER = RGBColor(60, 60, 70)         # Subtle border

FONT_TITLE = "Arial Black"
FONT_BODY = "Arial"

ASSETS_DIR = "/home/s4nzcyber/.gemini/antigravity/scratch/persona5-style-portfolio/assets"
PHOTO_PATH = os.path.join(ASSETS_DIR, "me.jpg")

def add_full_bg(slide):
    """Draw dark canvas background"""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BG_DARK
    bg.line.fill.background()
    return bg

def add_header(slide, slide_num_str, category_str, title_str):
    """Draw stylish Persona 5 Header"""
    # Red top accent strip
    top_strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    top_strip.fill.solid()
    top_strip.fill.fore_color.rgb = COLOR_RED_PRIMARY
    top_strip.line.fill.background()

    # Category Tag Box (Red badge with white text)
    badge = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.4), Inches(2.8), Inches(0.38))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_RED_PRIMARY
    badge.line.fill.background()
    tf_b = badge.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = Inches(0.1)
    tf_b.margin_top = Inches(0.04)
    p_b = tf_b.paragraphs[0]
    p_b.text = f"{slide_num_str} // {category_str}"
    p_b.font.name = FONT_TITLE
    p_b.font.size = Pt(11)
    p_b.font.bold = True
    p_b.font.color.rgb = COLOR_WHITE

    # Main Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.85), Inches(11.8), Inches(0.7))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = title_str.upper()
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(24)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_WHITE

    # Bottom decorative slash / line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.58), Inches(11.733), Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_RED_DEEP
    line.line.fill.background()

def add_footer(slide, current_idx):
    """Subdued HUD bottom bar"""
    footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.35))
    tf_f = footer_box.text_frame
    tf_f.word_wrap = True
    p_f = tf_f.paragraphs[0]
    p_f.text = f"SANZ GOKU  ·  Satu Bot, Banyak Solusi  ·  snz.web.id                           SLIDE {current_idx:02d} / 08"
    p_f.font.name = FONT_BODY
    p_f.font.size = Pt(10)
    p_f.font.color.rgb = COLOR_TEXT_MUTED

# ==========================================
# SLIDE 1 — COVER
# ==========================================
slide1 = prs.slides.add_slide(blank_layout)
add_full_bg(slide1)

# Diagonal Red Graphic on Right Side
red_panel = slide1.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(8.5), Inches(0), Inches(4.833), Inches(7.5))
red_panel.fill.solid()
red_panel.fill.fore_color.rgb = COLOR_RED_PRIMARY
red_panel.line.fill.background()
red_panel.rotation = 180

# Dark accent overlay on right
tri2 = slide1.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(10.5), Inches(0), Inches(2.833), Inches(7.5))
tri2.fill.solid()
tri2.fill.fore_color.rgb = COLOR_RED_DEEP
tri2.line.fill.background()
tri2.rotation = 180

# Top Tag
top_tag = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.8), Inches(3.4), Inches(0.42))
top_tag.fill.solid()
top_tag.fill.fore_color.rgb = COLOR_RED_PRIMARY
top_tag.line.fill.background()
tf_tag = top_tag.text_frame
p_tag = tf_tag.paragraphs[0]
p_tag.text = "PORTOFOLIO BISNIS DIGITAL"
p_tag.font.name = FONT_TITLE
p_tag.font.size = Pt(12)
p_tag.font.bold = True
p_tag.font.color.rgb = COLOR_WHITE

# Main Title (SANZ GOKU)
title_cover = slide1.shapes.add_textbox(Inches(0.75), Inches(1.4), Inches(8.5), Inches(1.8))
tf_c = title_cover.text_frame
tf_c.word_wrap = True
p_c1 = tf_c.paragraphs[0]
p_c1.text = "SANZ GOKU"
p_c1.font.name = FONT_TITLE
p_c1.font.size = Pt(50)
p_c1.font.bold = True
p_c1.font.color.rgb = COLOR_WHITE

# Tagline Badge
tagline_box = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.2), Inches(5.8), Inches(0.55))
tagline_box.fill.solid()
tagline_box.fill.fore_color.rgb = COLOR_BG_CARD
tagline_box.line.color.rgb = COLOR_ACCENT_YELLOW
tagline_box.line.width = Pt(1.5)
tf_tb = tagline_box.text_frame
p_tb = tf_tb.paragraphs[0]
p_tb.text = "« Satu Bot, Banyak Solusi »"
p_tb.font.name = FONT_TITLE
p_tb.font.size = Pt(16)
p_tb.font.bold = True
p_tb.font.color.rgb = COLOR_ACCENT_YELLOW

# Description
desc_cover = slide1.shapes.add_textbox(Inches(0.75), Inches(3.95), Inches(7.5), Inches(1.2))
tf_dc = desc_cover.text_frame
tf_dc.word_wrap = True
p_dc = tf_dc.paragraphs[0]
p_dc.text = "Penyedia produk & layanan digital, infrastruktur server Linux, panel Pterodactyl, nomor virtual OTP, serta solusi otomatisasi transaksi berbasis bot Telegram (@Sanzstrnokosmumerbot)."
p_dc.font.name = FONT_BODY
p_dc.font.size = Pt(13)
p_dc.font.color.rgb = COLOR_WHITE

# Presenter / Identity Box
id_box = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(5.3), Inches(7.2), Inches(1.4))
id_box.fill.solid()
id_box.fill.fore_color.rgb = COLOR_BG_CARD
id_box.line.color.rgb = COLOR_RED_PRIMARY
id_box.line.width = Pt(2)
tf_id = id_box.text_frame
tf_id.margin_left = Inches(0.25)
tf_id.margin_top = Inches(0.2)
p_id1 = tf_id.paragraphs[0]
p_id1.text = "DISUSUN OLEH:"
p_id1.font.name = FONT_TITLE
p_id1.font.size = Pt(11)
p_id1.font.color.rgb = COLOR_ACCENT_YELLOW
p_id2 = tf_id.add_paragraph()
p_id2.text = "Muhammad Sandika Akbar  (Sanz)"
p_id2.font.name = FONT_TITLE
p_id2.font.size = Pt(15)
p_id2.font.bold = True
p_id2.font.color.rgb = COLOR_WHITE
p_id3 = tf_id.add_paragraph()
p_id3.text = "Domain: snz.web.id  ·  Bot: @Sanzstrnokosmumerbot  ·  Admin: @Sanzxgoku"
p_id3.font.name = FONT_BODY
p_id3.font.size = Pt(11)
p_id3.font.color.rgb = COLOR_TEXT_MUTED

add_footer(slide1, 1)

# ==========================================
# SLIDE 2 — PROFIL BISNIS (with User Photo)
# ==========================================
slide2 = prs.slides.add_slide(blank_layout)
add_full_bg(slide2)
add_header(slide2, "02", "IDENTITAS & NILAI", "PROFIL SANZ GOKU")

# Left Column: User Polaroid Photo Card
if os.path.exists(PHOTO_PATH):
    # Polaroid White Container
    photo_card = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(2.9), Inches(4.95))
    photo_card.fill.solid()
    photo_card.fill.fore_color.rgb = COLOR_WHITE
    photo_card.line.color.rgb = COLOR_RED_PRIMARY
    photo_card.line.width = Pt(2)
    
    # Insert 1:1 image
    slide2.shapes.add_picture(PHOTO_PATH, Inches(0.95), Inches(1.95), Inches(2.6), Inches(2.6))
    
    # Polaroid caption text frame
    tf_pc = photo_card.text_frame
    tf_pc.margin_left = Inches(0.15)
    tf_pc.margin_right = Inches(0.15)
    tf_pc.margin_top = Inches(2.9)
    p_pc1 = tf_pc.paragraphs[0]
    p_pc1.text = "MUHAMMAD SANDIKA AKBAR"
    p_pc1.font.name = FONT_TITLE
    p_pc1.font.size = Pt(10)
    p_pc1.font.bold = True
    p_pc1.font.color.rgb = RGBColor(17, 17, 17)
    
    p_pc2 = tf_pc.add_paragraph()
    p_pc2.text = "Founder / Owner (Sanz)\nDomain: snz.web.id\nBot: @Sanzstrnokosmumerbot\nAdmin: @Sanzxgoku"
    p_pc2.font.name = FONT_BODY
    p_pc2.font.size = Pt(9.5)
    p_pc2.font.color.rgb = RGBColor(60, 60, 60)
    p_pc2.space_before = Pt(4)

# Right Column: Executive Summary + 4 Pillars
# Executive Summary
card_top = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.9), Inches(1.8), Inches(8.633), Inches(1.3))
card_top.fill.solid()
card_top.fill.fore_color.rgb = COLOR_BG_CARD
card_top.line.color.rgb = COLOR_RED_PRIMARY
card_top.line.width = Pt(1.5)
tf_ct = card_top.text_frame
tf_ct.word_wrap = True
tf_ct.margin_left = Inches(0.25)
tf_ct.margin_right = Inches(0.25)
tf_ct.margin_top = Inches(0.15)
p_ct1 = tf_ct.paragraphs[0]
p_ct1.text = "RINGKASAN EKSEKUTIF"
p_ct1.font.name = FONT_TITLE
p_ct1.font.size = Pt(11)
p_ct1.font.color.rgb = COLOR_ACCENT_YELLOW
p_ct2 = tf_ct.add_paragraph()
p_ct2.text = "SANZ GOKU adalah usaha digital aktif yang berfokus pada penyediaan produk & layanan digital, layanan server, serta solusi teknologi praktis. Usaha ini hadir untuk menjawab kebutuhan pelanggan akan infrastruktur hosting stabil, nomor virtual OTP cepat via bot Telegram, dan alur transaksi yang jelas serta transparan."
p_ct2.font.name = FONT_BODY
p_ct2.font.size = Pt(11)
p_ct2.font.color.rgb = COLOR_WHITE

# 4 Key Pillars (2x2 grid in right column)
pillars = [
    ("BIDANG USAHA", [
        "Penjualan & konfigurasi VPS Linux.",
        "Instalasi panel Pterodactyl & hosting.",
        "Layanan nomor virtual OTP (RumahOTP).",
        "Script, subdomain, dan produk digital."
    ]),
    ("KEBUTUHAN PELANGGAN", [
        "Infrastruktur server fleksibel & stabil.",
        "Verifikasi akun instan tanpa nomor pribadi.",
        "Kebutuhan tools & script digital praktis.",
        "Pemesanan responsif dibantu admin."
    ]),
    ("TUJUAN PENDIRIAN", [
        "Solusi teknologi terjangkau & mudah diakses.",
        "Ekosistem mandiri berbasis efisiensi server.",
        "Otomatisasi pelayanan pelanggan bertahap.",
        "Penyedia layanan digital yang dapat diandalkan."
    ]),
    ("NILAI UTAMA BISNIS", [
        "Transparansi spesifikasi tanpa klaim palsu.",
        "Kecepatan penanganan pesanan via Telegram.",
        "Inovasi bot Telegram (@Sanzstrnokosmumerbot).",
        "Dukungan teknis responsif saat ada kendala."
    ]),
]

xp = [Inches(3.9), Inches(8.3)]
yp = [Inches(3.3), Inches(5.15)]

for idx, (p_title, p_items) in enumerate(pillars):
    col = idx % 2
    row = idx // 2
    w_box = Inches(4.233)
    h_box = Inches(1.68)
    
    box = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, xp[col], yp[row], w_box, h_box)
    box.fill.solid()
    box.fill.fore_color.rgb = COLOR_BG_CARD
    box.line.color.rgb = COLOR_BORDER
    box.line.width = Pt(1)
    
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.12)
    
    p0 = tf.paragraphs[0]
    p0.text = f"■  {p_title}"
    p0.font.name = FONT_TITLE
    p0.font.size = Pt(11)
    p0.font.bold = True
    p0.font.color.rgb = COLOR_RED_PRIMARY
    
    for item in p_items:
        pi = tf.add_paragraph()
        pi.text = f"• {item}"
        pi.font.name = FONT_BODY
        pi.font.size = Pt(9.5)
        pi.font.color.rgb = COLOR_WHITE

add_footer(slide2, 2)

# ==========================================
# SLIDE 3 — PRODUK DAN LAYANAN
# ==========================================
slide3 = prs.slides.add_slide(blank_layout)
add_full_bg(slide3)
add_header(slide3, "03", "KATALOG BISNIS", "PRODUK DAN LAYANAN")

# Status Warning Bar
status_bar = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(11.733), Inches(0.42))
status_bar.fill.solid()
status_bar.fill.fore_color.rgb = COLOR_BG_CARD
status_bar.line.color.rgb = COLOR_ACCENT_YELLOW
status_bar.line.width = Pt(1)
tf_sb = status_bar.text_frame
tf_sb.margin_left = Inches(0.25)
tf_sb.margin_top = Inches(0.08)
p_sb = tf_sb.paragraphs[0]
p_sb.text = "CATATAN STATUS: Layanan dikelompokkan berdasarkan kategori. Membedakan layanan aktif berjualan vs tahap rencana/pengembangan."
p_sb.font.name = FONT_BODY
p_sb.font.size = Pt(10)
p_sb.font.color.rgb = COLOR_ACCENT_YELLOW

services = [
    ("VPS & SERVER LINUX", "STATUS: AKTIF BERJUALAN", COLOR_RED_PRIMARY, [
        "Penjualan VPS Linux (berbagai pilihan spek).",
        "Setup & lisensi/manajemen panel Pterodactyl.",
        "Pengelolaan konfigurasi dasar server Linux.",
        "Hosting server sesuai ketersediaan stok."
    ]),
    ("NOKOS & LAYANAN OTP", "STATUS: AKTIF BERJUALAN", COLOR_RED_PRIMARY, [
        "Nomor virtual verifikasi akun aplikasi digital.",
        "Tersedia melalui Bot @Sanzstrnokosmumerbot.",
        "Integrasi penyedia terpercaya (RumahOTP).",
        "Aktivasi akun aman tanpa nomor pribadi."
    ]),
    ("PRODUK DIGITAL & LAINNYA", "STATUS: AKTIF / ON-DEMAND", COLOR_RED_PRIMARY, [
        "Penjualan script dan produk digital terpilih.",
        "Layanan penyediaan custom subdomain.",
        "Layanan boost media sosial & kebutuhan akun.",
        "Top up saldo dan layanan digital pelengkap."
    ]),
    ("BOT TELEGRAM (@Sanzstrnokosmumerbot)", "STATUS: PENGEMBANGAN FITUR", COLOR_ACCENT_YELLOW, [
        "Katalog interaktif otomatis via Telegram.",
        "Pemeriksaan status transaksi & saldo real-time.",
        "Rencana integrasi payment gateway QRIS.",
        "Mempermudah alur pemesanan secara mandiri."
    ]),
]

w_card = Inches(2.78)
gap = Inches(0.2)
y_card = Inches(2.35)
h_card = Inches(4.5)

for i, (title_s, status_s, badge_color, items_s) in enumerate(services):
    x_card = Inches(0.8) + i * (w_card + gap)
    card = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_card, y_card, w_card, h_card)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_BG_CARD
    card.line.color.rgb = COLOR_BORDER
    card.line.width = Pt(1)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.18)
    
    p1 = tf.paragraphs[0]
    p1.text = title_s
    p1.font.name = FONT_TITLE
    p1.font.size = Pt(11.5)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_WHITE
    
    p2 = tf.add_paragraph()
    p2.text = f"[{status_s}]"
    p2.font.name = FONT_TITLE
    p2.font.size = Pt(9)
    p2.font.bold = True
    p2.font.color.rgb = badge_color
    
    p3 = tf.add_paragraph()
    p3.text = " "
    p3.font.size = Pt(4)
    
    for it in items_s:
        p_it = tf.add_paragraph()
        p_it.text = f"• {it}"
        p_it.font.name = FONT_BODY
        p_it.font.size = Pt(10.5)
        p_it.font.color.rgb = COLOR_WHITE
        p_it.space_after = Pt(6)

add_footer(slide3, 3)

# ==========================================
# SLIDE 4 — SISTEM OPERASIONAL BISNIS
# ==========================================
slide4 = prs.slides.add_slide(blank_layout)
add_full_bg(slide4)
add_header(slide4, "04", "ALUR PELAYANAN", "SISTEM OPERASIONAL BISNIS")

box_op = slide4.shapes.add_textbox(Inches(0.75), Inches(1.7), Inches(11.8), Inches(0.55))
tf_op = box_op.text_frame
p_op = tf_op.paragraphs[0]
p_op.text = "Alur Pelayanan Pelanggan 6 Langkah — Menggabungkan penanganan admin aktif dengan persiapan otomasi bot (@Sanzstrnokosmumerbot)."
p_op.font.name = FONT_BODY
p_op.font.size = Pt(12)
p_op.font.color.rgb = COLOR_TEXT_MUTED

steps = [
    ("01", "KATALOG & INFO", "Pelanggan melihat daftar layanan, deskripsi, dan ketersediaan stok melalui kanal resmi atau bot Telegram."),
    ("02", "PEMILIHAN LAYANAN", "Pelanggan memilih jenis layanan (VPS/OTP/Produk Digital) dan menyampaikan kebutuhan ke admin/bot."),
    ("03", "VERIFIKASI ADMIN", "Admin memeriksa rincian pesanan, kesiapan server/nomor, dan memberikan instruksi pembayaran."),
    ("04", "PEMBAYARAN", "Pelanggan melakukan pembayaran melalui metode yang tersedia dan diverifikasi secara transparan."),
    ("05", "PEMROSESAN SISTEM", "Pesanan diproses langsung melalui sistem internal server Linux atau penyedia layanan terkait."),
    ("06", "SERAH TERIMA & CS", "Hasil layanan diserahkan ke pelanggan (kredensial/kode OTP) disertai bantuan jika berkendala."),
]

for i, (num_s, title_step, desc_step) in enumerate(steps):
    col = i % 3
    row = i // 3
    x_s = Inches(0.8) + col * Inches(3.98)
    y_s = Inches(2.35) + row * Inches(2.15)
    w_s = Inches(3.78)
    h_s = Inches(1.95)
    
    card = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_s, y_s, w_s, h_s)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_BG_CARD
    card.line.color.rgb = COLOR_RED_DEEP
    card.line.width = Pt(1.5)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.15)
    
    p_num = tf.paragraphs[0]
    p_num.text = f"LANGKAH {num_s}"
    p_num.font.name = FONT_TITLE
    p_num.font.size = Pt(11)
    p_num.font.bold = True
    p_num.font.color.rgb = COLOR_RED_PRIMARY
    
    p_t = tf.add_paragraph()
    p_t.text = title_step
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_WHITE
    
    p_d = tf.add_paragraph()
    p_d.text = desc_step
    p_d.font.name = FONT_BODY
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = COLOR_TEXT_MUTED

note_box = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.58), Inches(11.733), Inches(0.38))
note_box.fill.solid()
note_box.fill.fore_color.rgb = COLOR_BG_CARD
note_box.line.color.rgb = COLOR_BORDER
tf_nb = note_box.text_frame
tf_nb.margin_left = Inches(0.2)
tf_nb.margin_top = Inches(0.06)
p_nb = tf_nb.paragraphs[0]
p_nb.text = "STATUS OPERASIONAL: Saat ini pesanan ditangani via admin. Integrasi Bot @Sanzstrnokosmumerbot terus disempurnakan untuk otomasi alur transaksi."
p_nb.font.name = FONT_BODY
p_nb.font.size = Pt(9.5)
p_nb.font.color.rgb = COLOR_ACCENT_YELLOW

add_footer(slide4, 4)

# ==========================================
# SLIDE 5 — STRATEGI PEMASARAN DAN CS
# ==========================================
slide5 = prs.slides.add_slide(blank_layout)
add_full_bg(slide5)
add_header(slide5, "05", "PEMASARAN & PELAYANAN", "STRATEGI PEMASARAN & PELAYANAN PELANGGAN")

# Left Column: Pemasaran (Width: 5.733)
box_mkt = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.733), Inches(3.9))
box_mkt.fill.solid()
box_mkt.fill.fore_color.rgb = COLOR_BG_CARD
box_mkt.line.color.rgb = COLOR_RED_PRIMARY
box_mkt.line.width = Pt(1.5)

tf_m = box_mkt.text_frame
tf_m.word_wrap = True
tf_m.margin_left = Inches(0.25)
tf_m.margin_right = Inches(0.25)
tf_m.margin_top = Inches(0.2)

pm_head = tf_m.paragraphs[0]
pm_head.text = "STRATEGI PEMASARAN DIGITAL"
pm_head.font.name = FONT_TITLE
pm_head.font.size = Pt(14)
pm_head.font.bold = True
pm_head.font.color.rgb = COLOR_RED_PRIMARY

mkt_points = [
    ("Pemanfaatan Kanal Digital", "Mengenalkan produk dan layanan melalui kanal media sosial dan komunitas online teknologi/gaming."),
    ("Katalog Produk yang Jelas", "Menyajikan spesifikasi layanan, batas penggunaan, serta panduan pemesanan tanpa klaim berlebihan."),
    ("Optimalisasi Telegram", "Menjadikan bot (@Sanzstrnokosmumerbot) dan chat admin (@Sanzxgoku) sebagai pusat interaksi harian yang cepat."),
    ("Harga Kompetitif & Realistis", "Menetapkan harga rasional sesuai spesifikasi server dan biaya penyedia tanpa perang harga merusak.")
]

for title_m, desc_m in mkt_points:
    p_t = tf_m.add_paragraph()
    p_t.text = f"■ {title_m}"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(11)
    p_t.font.color.rgb = COLOR_WHITE
    p_t.space_before = Pt(6)
    
    p_d = tf_m.add_paragraph()
    p_d.text = desc_m
    p_d.font.name = FONT_BODY
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = COLOR_TEXT_MUTED

# Right Column: Pelayanan Pelanggan (Width: 5.733)
box_cs = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), Inches(1.85), Inches(5.733), Inches(3.9))
box_cs.fill.solid()
box_cs.fill.fore_color.rgb = COLOR_BG_CARD
box_cs.line.color.rgb = COLOR_RED_PRIMARY
box_cs.line.width = Pt(1.5)

tf_c = box_cs.text_frame
tf_c.word_wrap = True
tf_c.margin_left = Inches(0.25)
tf_c.margin_right = Inches(0.25)
tf_c.margin_top = Inches(0.2)

pc_head = tf_c.paragraphs[0]
pc_head.text = "PELAYANAN & KEPERCAYAAN PELANGGAN"
pc_head.font.name = FONT_TITLE
pc_head.font.size = Pt(14)
pc_head.font.bold = True
pc_head.font.color.rgb = COLOR_RED_PRIMARY

cs_points = [
    ("Komunikasi Responsif", "Admin merespon pesan, pertanyaan spesifikasi, dan pesanan secara cepat serta lugas."),
    ("Bantuan & Troubleshooting", "Membantu pelanggan jika menemui kendala instalasi panel atau kendala kode OTP yang belum masuk."),
    ("Transparansi Stok & Server", "Memberitahukan ketersediaan kapasitas server apa adanya demi menjaga kepuasan pelanggan."),
    ("Membangun Loyalitas", "Pelayanan yang jujur dan tuntas menjadi fondasi repeat order tanpa perlu manipulasi testimoni.")
]

for title_c, desc_c in cs_points:
    p_t = tf_c.add_paragraph()
    p_t.text = f"■ {title_c}"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(11)
    p_t.font.color.rgb = COLOR_WHITE
    p_t.space_before = Pt(6)
    
    p_d = tf_c.add_paragraph()
    p_d.text = desc_c
    p_d.font.name = FONT_BODY
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = COLOR_TEXT_MUTED

# Bottom Contact Bar
contact_box = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(5.95), Inches(11.733), Inches(0.85))
contact_box.fill.solid()
contact_box.fill.fore_color.rgb = COLOR_BG_CARD
contact_box.line.color.rgb = COLOR_ACCENT_YELLOW
contact_box.line.width = Pt(1.5)

tf_cb = contact_box.text_frame
tf_cb.word_wrap = True
tf_cb.margin_left = Inches(0.3)
tf_cb.margin_top = Inches(0.12)

p_cb1 = tf_cb.paragraphs[0]
p_cb1.text = "KANAL RESMI KOMUNIKASI & TRANSAKSI:"
p_cb1.font.name = FONT_TITLE
p_cb1.font.size = Pt(11)
p_cb1.font.color.rgb = COLOR_ACCENT_YELLOW

p_cb2 = tf_cb.add_paragraph()
p_cb2.text = "Bot Telegram: @Sanzstrnokosmumerbot  |  Admin: @Sanzxgoku  |  Instagram: @sanzshun  |  Domain: snz.web.id"
p_cb2.font.name = FONT_TITLE
p_cb2.font.size = Pt(12)
p_cb2.font.bold = True
p_cb2.font.color.rgb = COLOR_WHITE

add_footer(slide5, 5)

# ==========================================
# SLIDE 6 — TEKNOLOGI PENDUKUNG
# ==========================================
slide6 = prs.slides.add_slide(blank_layout)
add_full_bg(slide6)
add_header(slide6, "06", "STACK & INFRASTRUKTUR", "TEKNOLOGI PENDUKUNG BISNIS")

box_act = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.733), Inches(4.95))
box_act.fill.solid()
box_act.fill.fore_color.rgb = COLOR_BG_CARD
box_act.line.color.rgb = COLOR_RED_PRIMARY
box_act.line.width = Pt(1.5)

tf_a = box_act.text_frame
tf_a.word_wrap = True
tf_a.margin_left = Inches(0.25)
tf_a.margin_right = Inches(0.25)
tf_a.margin_top = Inches(0.18)

pa_h = tf_a.paragraphs[0]
pa_h.text = "TEKNOLOGI AKTIF DIGUNAKAN"
pa_h.font.name = FONT_TITLE
pa_h.font.size = Pt(13)
pa_h.font.bold = True
pa_h.font.color.rgb = COLOR_RED_PRIMARY

tech_active = [
    ("Linux VPS (Ubuntu / Debian)", "Infrastruktur dasar server untuk hosting layanan, menjalankan script, dan pengelolaan port jaringan."),
    ("JavaScript & Node.js", "Bahasa pemrograman dan runtime utama dalam membangun logika bot, skrip otomatisasi, serta komunikasi API."),
    ("Telegram Bot API", "Protokol penghubung untuk interaksi chat bot @Sanzstrnokosmumerbot, menu tombol, dan notifikasi."),
    ("Pterodactyl Panel", "Sistem panel web berbasis container Docker untuk mempermudah kontrol server game dan bot secara mandiri."),
    ("Cloudflare DNS & Proxy", "Manajemen domain snz.web.id, proteksi terhadap serangan DDoS, serta pengaturan routing subdomain.")
]

for t_name, t_desc in tech_active:
    p = tf_a.add_paragraph()
    p.text = f"■ {t_name}"
    p.font.name = FONT_TITLE
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_WHITE
    p.space_before = Pt(6)
    
    pd = tf_a.add_paragraph()
    pd.text = t_desc
    pd.font.name = FONT_BODY
    pd.font.size = Pt(9.5)
    pd.font.color.rgb = COLOR_TEXT_MUTED

box_pln = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), Inches(1.85), Inches(5.733), Inches(4.95))
box_pln.fill.solid()
box_pln.fill.fore_color.rgb = COLOR_BG_CARD
box_pln.line.color.rgb = COLOR_ACCENT_YELLOW
box_pln.line.width = Pt(1.5)

tf_p = box_pln.text_frame
tf_p.word_wrap = True
tf_p.margin_left = Inches(0.25)
tf_p.margin_right = Inches(0.25)
tf_p.margin_top = Inches(0.18)

pp_h = tf_p.paragraphs[0]
pp_h.text = "INTEGRASI API & PROYEK PENGEMBANGAN"
pp_h.font.name = FONT_TITLE
pp_h.font.size = Pt(13)
pp_h.font.bold = True
pp_h.font.color.rgb = COLOR_ACCENT_YELLOW

tech_planned = [
    ("Integrasi API Pihak Ketiga (RumahOTP, FayuPedia)", "Rencana sinkronisasi stok dan eksekusi nomor OTP serta produk digital secara otomatis melalui REST API."),
    ("Pakasir (Gateway Pembayaran QRIS)", "Perencanaan sistem verifikasi pembayaran QRIS instan untuk bot Telegram dan web store."),
    ("Proyek SOTP (sxgotp)", "Pengembangan terpisah website layanan OTP modern menggunakan Next.js, React, Tailwind CSS, Vercel, dan Supabase."),
    ("Manfaat Utama Teknologi", "• Efisiensi Operasional: Mengurangi pekerjaan admin berulang.\n• Kecepatan Layanan: Transaksi diproses dalam hitungan detik.\n• Skalabilitas: Kapasitas sistem siap menangani lonjakan pesanan.")
]

for t_name, t_desc in tech_planned:
    p = tf_p.add_paragraph()
    p.text = f"◆ {t_name}"
    p.font.name = FONT_TITLE
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_WHITE
    p.space_before = Pt(6)
    
    pd = tf_p.add_paragraph()
    pd.text = t_desc
    pd.font.name = FONT_BODY
    pd.font.size = Pt(9.5)
    pd.font.color.rgb = COLOR_TEXT_MUTED

add_footer(slide6, 6)

# ==========================================
# SLIDE 7 — VISI DAN RENCANA PENGEMBANGAN
# ==========================================
slide7 = prs.slides.add_slide(blank_layout)
add_full_bg(slide7)
add_header(slide7, "07", "ROADMAP BISNIS", "VISI DAN RENCANA PENGEMBANGAN")

box_v = slide7.shapes.add_textbox(Inches(0.75), Inches(1.7), Inches(11.8), Inches(0.5))
tf_v = box_v.text_frame
p_v = tf_v.paragraphs[0]
p_v.text = "Rencana Strategis Masa Depan — Disajikan sebagai visi pengembangan jangka pendek dan menengah."
p_v.font.name = FONT_BODY
p_v.font.size = Pt(12)
p_v.font.color.rgb = COLOR_TEXT_MUTED

plans = [
    ("01", "EKSPANSI LAYANAN DIGITAL", "Memperluas katalog produk digital, menambah varian spesifikasi server VPS, dan menyesuaikan layanan dengan tren kebutuhan pasar."),
    ("02", "PENINGKATAN MUTU LAYANAN", "Menjaga kestabilan uptime server, kecepatan delivery pesanan, serta evaluasi kepuasan pelanggan secara berkala."),
    ("03", "OTOMASI BOT (@Sanzstrnokosmumerbot)", "Menyelesaikan integrasi pemesanan mandiri, verifikasi QRIS otomatis, dan penyerahan kredensial/layanan langsung via bot."),
    ("04", "SISTEM PENCATATAN & PEMBUKUAN", "Menerapkan pencatatan transaksi terstruktur untuk memantau performa layanan, inventaris stok server, dan arus keuangan rapi."),
    ("05", "KEAMANAN & KEANDALAN SISTEM", "Memperkuat proteksi firewall server Linux, manajemen akses API yang aman, serta cadangan data (backup) rutin."),
    ("06", "EKOSISTEM BERKELANJUTAN", "Mengintegrasikan website (snz.web.id), bot Telegram, dan panel server menjadi satu ekosistem bisnis digital yang stabil.")
]

for i, (num_p, title_p, desc_p) in enumerate(plans):
    col = i % 2
    row = i // 2
    x_p = Inches(0.8) + col * Inches(6.0)
    y_p = Inches(2.35) + row * Inches(1.48)
    w_p = Inches(5.733)
    h_p = Inches(1.35)
    
    card = slide7.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_p, y_p, w_p, h_p)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_BG_CARD
    card.line.color.rgb = COLOR_BORDER
    card.line.width = Pt(1)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.12)
    
    p_t = tf.paragraphs[0]
    p_t.text = f"{num_p}. {title_p}"
    p_t.font.name = FONT_TITLE
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_RED_PRIMARY
    
    p_d = tf.add_paragraph()
    p_d.text = desc_p
    p_d.font.name = FONT_BODY
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = COLOR_WHITE

add_footer(slide7, 7)

# ==========================================
# SLIDE 8 — PENUTUP & KESIMPULAN
# ==========================================
slide8 = prs.slides.add_slide(blank_layout)
add_full_bg(slide8)
add_header(slide8, "08", "PENUTUP", "KESIMPULAN BISNIS")

card_close = slide8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(11.733), Inches(2.6))
card_close.fill.solid()
card_close.fill.fore_color.rgb = COLOR_BG_CARD
card_close.line.color.rgb = COLOR_RED_PRIMARY
card_close.line.width = Pt(2)

tf_cl = card_close.text_frame
tf_cl.word_wrap = True
tf_cl.margin_left = Inches(0.4)
tf_cl.margin_right = Inches(0.4)
tf_cl.margin_top = Inches(0.3)

p_cl_tag = tf_cl.paragraphs[0]
p_cl_tag.text = "KESIMPULAN EKSEKUTIF"
p_cl_tag.font.name = FONT_TITLE
p_cl_tag.font.size = Pt(12)
p_cl_tag.font.color.rgb = COLOR_ACCENT_YELLOW

p_cl_body = tf_cl.add_paragraph()
p_cl_body.text = "SANZ GOKU merupakan usaha digital yang berfokus pada produk dan layanan teknologi. Melalui pengembangan layanan server, produk digital, dan otomatisasi menggunakan bot Telegram (@Sanzstrnokosmumerbot), bisnis ini diarahkan untuk memberikan pelayanan yang praktis, transparan, dan terus berkembang sesuai kebutuhan pelanggan."
p_cl_body.font.name = FONT_BODY
p_cl_body.font.size = Pt(13.5)
p_cl_body.font.color.rgb = COLOR_WHITE
p_cl_body.space_before = Pt(8)

p_cl_motto = tf_cl.add_paragraph()
p_cl_motto.text = "“ SANZ GOKU — Satu Bot, Banyak Solusi. ”"
p_cl_motto.font.name = FONT_TITLE
p_cl_motto.font.size = Pt(16)
p_cl_motto.font.bold = True
p_cl_motto.font.color.rgb = COLOR_ACCENT_YELLOW
p_cl_motto.space_before = Pt(12)

# Bottom Row: Closing remark and Official Contacts
box_ty = slide8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(4.7), Inches(4.5), Inches(2.1))
box_ty.fill.solid()
box_ty.fill.fore_color.rgb = COLOR_RED_PRIMARY
box_ty.line.fill.background()

tf_ty = box_ty.text_frame
tf_ty.word_wrap = True
tf_ty.margin_left = Inches(0.3)
tf_ty.margin_top = Inches(0.45)
p_ty1 = tf_ty.paragraphs[0]
p_ty1.text = "TERIMA KASIH"
p_ty1.font.name = FONT_TITLE
p_ty1.font.size = Pt(24)
p_ty1.font.bold = True
p_ty1.font.color.rgb = COLOR_WHITE

p_ty2 = tf_ty.add_paragraph()
p_ty2.text = "Presentasi Portofolio Bisnis Selesai."
p_ty2.font.name = FONT_BODY
p_ty2.font.size = Pt(12)
p_ty2.font.color.rgb = COLOR_WHITE
p_ty2.space_before = Pt(6)

box_ct = slide8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(4.7), Inches(7.033), Inches(2.1))
box_ct.fill.solid()
box_ct.fill.fore_color.rgb = COLOR_BG_CARD
box_ct.line.color.rgb = COLOR_BORDER
box_ct.line.width = Pt(1)

tf_ct2 = box_ct.text_frame
tf_ct2.word_wrap = True
tf_ct2.margin_left = Inches(0.3)
tf_ct2.margin_top = Inches(0.2)

p_ct_t = tf_ct2.paragraphs[0]
p_ct_t.text = "INFORMASI KONTAK & IDENTITAS RESMI"
p_ct_t.font.name = FONT_TITLE
p_ct_t.font.size = Pt(12)
p_ct_t.font.bold = True
p_ct_t.font.color.rgb = COLOR_ACCENT_YELLOW

contacts = [
    ("Nama Pemilik", "Muhammad Sandika Akbar (Sanz)"),
    ("Bot Telegram", "@Sanzstrnokosmumerbot  (t.me/Sanzstrnokosmumerbot)"),
    ("Telegram Admin", "@Sanzxgoku  (t.me/Sanzxgoku)"),
    ("Instagram", "@sanzshun  (instagram.com/sanzshun)"),
    ("Domain Bisnis", "snz.web.id")
]

for label_c, val_c in contacts:
    p_row = tf_ct2.add_paragraph()
    p_row.text = f"• {label_c}: {val_c}"
    p_row.font.name = FONT_BODY
    p_row.font.size = Pt(10.5)
    p_row.font.color.rgb = COLOR_WHITE
    p_row.space_before = Pt(2)

add_footer(slide8, 8)

# Output Paths
output_dir = "/home/s4nzcyber/.gemini/antigravity/scratch/persona5-style-portfolio"
pptx_filename = "Portofolio_Bisnis_SANZ_GOKU.pptx"
output_path = os.path.join(output_dir, pptx_filename)
prs.save(output_path)

# Also save copy as previous name just in case
prs.save(os.path.join(output_dir, "Portofolio_Bisnis_SANZ_GOKU_STORE.pptx"))

print(f"Presentation successfully created at: {output_path}")
print(f"Total slides: {len(prs.slides)}")
