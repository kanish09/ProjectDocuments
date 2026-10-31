#!/usr/bin/env python3
"""Generate PowerPoint presentation for KhelCircle – Fantasy Esports Platform."""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree

OUT = "/home/abcom/myproject/Full_project/projectdocuments/presentation"

# ── Palette ──────────────────────────────────────────────────────────────────
C_DARK   = RGBColor(0x0D, 0x1B, 0x2A)
C_NAVY   = RGBColor(0x10, 0x25, 0x3C)
C_CARD   = RGBColor(0x16, 0x2A, 0x42)
C_ACCENT = RGBColor(0x1E, 0x90, 0xFF)
C_GOLD   = RGBColor(0xFF, 0xC7, 0x00)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_LGRAY  = RGBColor(0xCC, 0xDD, 0xFF)
C_GREEN  = RGBColor(0x22, 0xAA, 0x44)
C_RED    = RGBColor(0xCC, 0x33, 0x33)
C_PURPLE = RGBColor(0x88, 0x44, 0xCC)
C_ORANGE = RGBColor(0xFF, 0x88, 0x22)
C_MGRAY  = RGBColor(0x88, 0xBB, 0xFF)

# cm helpers
def cm(v): return Cm(v)

# ── Helpers ───────────────────────────────────────────────────────────────────

def bg(slide, r, g, b):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(r, g, b)

def rect(slide, l, t, w, h, fill_c, line_c=None, lw=0):
    from pptx.util import Pt as PPt
    s = slide.shapes.add_shape(1, cm(l), cm(t), cm(w), cm(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill_c
    if line_c:
        s.line.color.rgb = line_c; s.line.width = PPt(lw)
    else:
        s.line.fill.background()
    return s

def txb(slide, text, l, t, w, h, sz=14, bold=False, italic=False,
        col=None, align=PP_ALIGN.LEFT, wrap=True):
    col = col or C_WHITE
    box = slide.shapes.add_textbox(cm(l), cm(t), cm(w), cm(h))
    tf = box.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(sz); run.font.bold = bold
    run.font.italic = italic; run.font.color.rgb = col
    run.font.name = "Calibri"
    return box

def add_transition(slide):
    ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    xml = (f'<p:transition xmlns:p="{ns}" spd="med" advClick="1" advTm="0">'
           f'<p:fade/></p:transition>')
    existing = slide._element.find(f'{{{ns}}}transition')
    if existing is not None:
        slide._element.remove(existing)
    slide._element.append(parse_xml(xml))

def header(slide, title, subtitle=None):
    rect(slide, 0, 0, 33.87, 2.8, C_DARK)
    rect(slide, 0, 2.8, 33.87, 0.28, C_ACCENT)
    txb(slide, title, 0.8, 0.25, 32, 1.6, sz=30, bold=True, col=C_WHITE)
    if subtitle:
        txb(slide, subtitle, 0.8, 1.75, 32, 0.95, sz=14, italic=True, col=C_MGRAY)

def card(slide, title, body, l, t, w, h, hcol=None, bcol=None, tsz=13, bsz=12):
    hcol = hcol or C_ACCENT
    bcol = bcol or C_CARD
    rect(slide, l, t, w, h, bcol)
    rect(slide, l, t, w, 0.65, hcol)
    txb(slide, title, l+0.15, t+0.05, w-0.3, 0.58, sz=tsz, bold=True, col=C_WHITE)
    txb(slide, body,  l+0.15, t+0.75, w-0.3, h-0.9, sz=bsz, col=C_LGRAY)

def bullets(slide, items, l, t, w, fsz=13, gap=0.72, col=None):
    col = col or C_WHITE
    y = t
    for item in items:
        txb(slide, f"▸  {item}", l, y, w, gap+0.1, sz=fsz, col=col)
        y += gap
    return y

# ─────────────────────────────────────────────────────────────────────────────

def build():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 1 – TITLE
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    rect(sl, 0, 0, 33.87, 0.5, C_ACCENT)
    rect(sl, 0, 18.5, 33.87, 0.6, C_ACCENT)
    # decorative glow rect
    glow = sl.shapes.add_shape(9, cm(20), cm(1.5), cm(12), cm(12))
    glow.fill.solid(); glow.fill.fore_color.rgb = RGBColor(0x05,0x2A,0x55)
    glow.line.fill.background()
    txb(sl,"🎮", 9, 1.5, 16, 3.5, sz=68, align=PP_ALIGN.CENTER)
    txb(sl,"KhelCircle", 0, 5.5, 33.87, 3.2,
        sz=76, bold=True, col=C_ACCENT, align=PP_ALIGN.CENTER)
    txb(sl,"Fantasy Esports Platform", 0, 8.8, 33.87, 2,
        sz=32, italic=True, col=C_WHITE, align=PP_ALIGN.CENTER)
    txb(sl,"CS2  ·  Valorant  ·  BGMI", 0, 10.8, 33.87, 1.4,
        sz=20, col=C_MGRAY, align=PP_ALIGN.CENTER)
    txb(sl,"Angular 17  ·  Node.js  ·  MongoDB  ·  Socket.io  ·  Ionic  ·  Capacitor",
        0, 12.3, 33.87, 1.4, sz=15, col=RGBColor(0x66,0x99,0xCC), align=PP_ALIGN.CENTER)
    txb(sl,"Comprehensive Project Presentation  |  September 2026",
        0, 14.2, 33.87, 1.2, sz=13, italic=True,
        col=RGBColor(0x77,0x88,0xAA), align=PP_ALIGN.CENTER)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 2 – AGENDA
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Agenda","What We Will Cover Today")
    left_items = [
        "01  Problem Statement",
        "02  Objectives & Scope",
        "03  Existing vs. Proposed System",
        "04  System Architecture",
        "05  Technologies Used",
        "06  Core Modules",
    ]
    right_items = [
        "07  Scoring Algorithm",
        "08  Database Design",
        "09  UI/UX Walkthrough",
        "10  Implementation Details",
        "11  Testing & Results",
        "12  Future Scope & Conclusion",
    ]
    for i, item in enumerate(left_items):
        y = 3.5 + i*1.35
        rect(sl, 0.8, y, 14.8, 1.1, C_CARD)
        txb(sl, item, 1.1, y+0.18, 14.3, 0.78, sz=14, col=C_WHITE)
    for i, item in enumerate(right_items):
        y = 3.5 + i*1.35
        rect(sl, 17.4, y, 14.8, 1.1, C_CARD)
        txb(sl, item, 17.7, y+0.18, 14.3, 0.78, sz=14, col=C_WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 3 – PROJECT OVERVIEW
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Project Overview","KhelCircle at a Glance")
    cards_data = [
        ("🎮  Platform","Fantasy Esports Gaming\nWeb + Android + iOS"),
        ("🏆  Games Covered","CS2  ·  Valorant  ·  BGMI"),
        ("⚡  Real-Time","Live Leaderboard via Socket.io\n(updates every 2 minutes)"),
        ("📱  Mobile App","Native Android & iOS\nvia Ionic + Capacitor"),
        ("🔒  Security","JWT + bcrypt + Helmet.js\n+ Rate Limiting"),
        ("☁️  Cloud Stack","MongoDB Atlas + Cloudinary CDN\n+ SMS OTP Gateway"),
    ]
    cols_c = [C_ACCENT, C_GREEN, C_PURPLE, C_ORANGE, C_RED, RGBColor(0x00,0x99,0xCC)]
    for idx,(t,b) in enumerate(cards_data):
        row = idx//3; col_i = idx%3
        card(sl, t, b, 0.6+col_i*11.1, 3.5+row*4.0, 10.5, 3.6, hcol=cols_c[idx], tsz=14, bsz=13)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 4 – PROBLEM STATEMENT
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Problem Statement","The Gap We Are Filling")
    rect(sl, 0.6, 3.1, 32.6, 16.2, C_NAVY)
    problems = [
        "No unified platform covers CS2, Valorant AND BGMI fantasy gaming",
        "Existing platforms update scores only AFTER match completion — zero real-time data",
        "Fantasy mechanics not tailored to tactical esports (ACS, KAST, first-blood, clutch)",
        "Poor mobile experience — not optimised for gaming-first audiences",
        "Opaque scoring systems — users cannot verify how their points are calculated",
        "Admins lack efficient tools to manage matches, players, and contests",
    ]
    for i, prob in enumerate(problems):
        y = 3.5 + i*2.55
        rect(sl, 1.0, y, 1.2, 2.1, C_RED)
        txb(sl, "⚠", 1.0, y, 1.2, 2.1, sz=24, align=PP_ALIGN.CENTER, col=C_WHITE)
        txb(sl, prob, 2.5, y+0.35, 30.2, 1.5, sz=14, col=C_WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 5 – OBJECTIVES
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Objectives","What KhelCircle Aims to Achieve")
    objs = [
        ("🎯","Multi-game platform — CS2, Valorant & BGMI in one app"),
        ("📊","Transparent scoring engine with documented, verifiable rules"),
        ("⚡","Real-time leaderboards via WebSocket (Socket.io) every 2 min"),
        ("📱","Cross-platform: Web + Android + iOS via Ionic & Capacitor"),
        ("🔐","Secure auth: JWT + bcrypt + OTP SMS verification"),
        ("🛠","Full admin panel — match, player & contest management"),
        ("☁️","Cloudinary CDN for scalable player & team image delivery"),
        ("📰","Esports news & live-scores module for user engagement"),
    ]
    for i,(icon,text) in enumerate(objs):
        row = i//2; col_i = i%2
        x = 0.6 + col_i*16.6; y = 3.3 + row*2.65
        rect(sl, x, y, 15.8, 2.25, C_CARD)
        rect(sl, x, y, 2.8, 2.25, C_ACCENT)
        txb(sl, icon, x, y, 2.8, 2.25, sz=28, align=PP_ALIGN.CENTER, col=C_WHITE)
        txb(sl, text, x+3.0, y+0.62, 12.5, 1.05, sz=13.5, col=C_WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 6 – EXISTING vs PROPOSED
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Existing vs. Proposed System","Why KhelCircle Is Different")
    # left – existing
    rect(sl, 0.4, 3.2, 15.5, 16.2, RGBColor(0x2A,0x10,0x10))
    rect(sl, 0.4, 3.2, 15.5, 1.3, C_RED)
    txb(sl,"❌  Existing Systems", 0.7, 3.3, 15, 1.1, sz=16, bold=True, col=C_WHITE)
    exist_items = ["Cricket/Football only (Dream11, MyTeam11)",
                   "No real-time score updates during match",
                   "Desktop-only experiences",
                   "Opaque, unverifiable point systems",
                   "No unified multi-game platform",
                   "No dedicated admin management tools"]
    for i, itm in enumerate(exist_items):
        txb(sl, f"✗  {itm}", 0.7, 4.8+i*2.2, 15, 2.0, sz=13,
            col=RGBColor(0xFF,0xAA,0xAA))
    # right – proposed
    rect(sl, 17.0, 3.2, 15.5, 16.2, RGBColor(0x0A,0x2A,0x12))
    rect(sl, 17.0, 3.2, 15.5, 1.3, C_GREEN)
    txb(sl,"✅  KhelCircle", 17.3, 3.3, 15, 1.1, sz=16, bold=True, col=C_WHITE)
    prop_items = ["CS2 + Valorant + BGMI in one unified app",
                  "Live WebSocket leaderboard (every 2 min)",
                  "Web + Android + iOS via Capacitor",
                  "Fully transparent scoring rules — published",
                  "Unified multi-game fantasy ecosystem",
                  "Comprehensive admin panel & Excel import"]
    for i, itm in enumerate(prop_items):
        txb(sl, f"✓  {itm}", 17.3, 4.8+i*2.2, 15, 2.0, sz=13,
            col=RGBColor(0x88,0xFF,0xAA))
    # VS badge
    rect(sl, 15.5, 9.5, 3.0, 3.0, C_GOLD)
    txb(sl,"VS", 15.5, 9.5, 3.0, 3.0, sz=30, bold=True, col=C_DARK, align=PP_ALIGN.CENTER)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 7 – SYSTEM ARCHITECTURE
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"System Architecture","Three-Tier + Real-Time WebSocket Layer")
    tiers = [
        ("📱  CLIENT TIER",
         "Angular 17 + Ionic\nCapacitor (Android/iOS)\nSocket.io Client\nAngular Router + Guards",
         C_ACCENT, 0.5),
        ("⚙️  APPLICATION TIER",
         "Node.js + Express.js\nJWT Auth Middleware\nHelmet · Rate Limiter\nSocket.io Server · node-cron",
         C_PURPLE, 8.8),
        ("🗄️  DATA TIER",
         "MongoDB Atlas\nMongoose v8 ODM\nCloudinary CDN\nCompound Indexes",
         C_GREEN, 17.1),
        ("📡  SERVICES TIER",
         "SMS Gateway (OTP)\nCloudinary API\nExternal Game APIs\n(future Riot / FACEIT)",
         C_ORANGE, 25.4),
    ]
    for title, body, col, x in tiers:
        rect(sl, x, 3.5, 7.9, 11.5, C_NAVY)
        rect(sl, x, 3.5, 7.9, 1.6, col)
        txb(sl, title, x+0.2, 3.6, 7.5, 1.4, sz=13, bold=True, col=C_WHITE)
        txb(sl, body,  x+0.25, 5.3, 7.4, 9.5, sz=12.5, col=C_LGRAY)
    # arrows
    for ax in [8.5, 16.8, 25.1]:
        txb(sl,"→", ax, 8.3, 0.75, 1.8, sz=32, bold=True, col=C_GOLD, align=PP_ALIGN.CENTER)
    # protocol labels
    txb(sl,"HTTPS REST\n+ WebSocket",  8.3, 10.4, 1.6, 2.0, sz=9,
        col=RGBColor(0xAA,0xBB,0xCC), align=PP_ALIGN.CENTER)
    txb(sl,"Mongoose\nODM",           16.6, 10.4, 1.6, 1.8, sz=9,
        col=RGBColor(0xAA,0xBB,0xCC), align=PP_ALIGN.CENTER)
    txb(sl,"HTTPS\nSDK Calls",        24.9, 10.4, 1.6, 1.8, sz=9,
        col=RGBColor(0xAA,0xBB,0xCC), align=PP_ALIGN.CENTER)
    txb(sl,"Security: Helmet.js · CORS · express-rate-limit · express-validator · bcryptjs",
        0.4, 15.3, 33.1, 1.1, sz=12, col=C_GOLD, align=PP_ALIGN.CENTER)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 8 – TECHNOLOGIES
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Technologies Used","Full-Stack Technology Stack")
    tech_cols = [
        ("Frontend", ["Angular 17 (Standalone Components)","Ionic Framework (Mobile UI)",
                      "Capacitor (Native Bridge)","TypeScript + SCSS",
                      "RxJS · Angular Router","Socket.io-client"], C_ACCENT, 0.5),
        ("Backend",  ["Node.js + Express.js v4","Socket.io v4 (WebSocket)",
                      "JWT + bcryptjs","Helmet · express-rate-limit",
                      "Multer · Axios · node-cron","dotenv (Config Mgmt)"], C_PURPLE, 9.0),
        ("Database", ["MongoDB (NoSQL Document DB)","Mongoose v8 ODM",
                      "MongoDB Atlas (Cloud)","Cloudinary (Media CDN)",
                      "Compound Indexing Strategy","GridFS (large file support)"], C_GREEN, 17.5),
        ("Deployment",["Android APK via Capacitor","PWA (Progressive Web App)",
                       "Render / Heroku (Backend)","Vercel / Netlify (Frontend)",
                       "MongoDB Atlas + Cloudinary","nginx Reverse Proxy"], C_ORANGE, 26.0),
    ]
    for g_title, items, col, x in tech_cols:
        rect(sl, x, 3.3, 8.1, 16.0, C_NAVY)
        rect(sl, x, 3.3, 8.1, 1.4, col)
        txb(sl, g_title, x+0.2, 3.38, 7.7, 1.25, sz=17, bold=True, col=C_WHITE)
        for j, itm in enumerate(items):
            txb(sl, f"▸  {itm}", x+0.3, 5.0+j*2.35, 7.5, 2.1, sz=12.5, col=C_WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 9 – MODULES
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Core Modules","Key Components of KhelCircle")
    modules = [
        ("🔐","Authentication","JWT Login · OTP SMS\nbcrypt Hashing · Admin Guard",C_ACCENT),
        ("🏟","Match Management","CS2 / Valorant / BGMI\nCRUD + Excel Import",C_PURPLE),
        ("👤","Player & Teams","Stats DB · Roles · Pricing\nCloudinary Images",C_GREEN),
        ("⚽","Fantasy Team Builder","5-player · 100-credit cap\nCaptain / Vice-Captain",C_ORANGE),
        ("🏆","Contest System","Free / Paid / Premium\nPrize Pool + Distribution",C_RED),
        ("📡","Real-Time Scoring","ScoringService · node-cron\nSocket.io Broadcast",C_GOLD),
        ("🛠","Admin Panel","Full CRUD · Stats Entry\nMatch Stats · Articles",RGBColor(0x00,0x99,0xCC)),
        ("📰","News Module","Articles · Live Scores\nBulk JSON Import",C_GREEN),
    ]
    for idx,(icon,title,body,col) in enumerate(modules):
        row = idx//4; col_i = idx%4
        x = 0.5 + col_i*8.35; y = 3.5 + row*5.5
        rect(sl, x, y, 7.9, 5.1, C_CARD)
        rect(sl, x, y, 7.9, 0.8, col)
        txb(sl, icon,  x+0.2, y+0.05, 1.2, 0.72, sz=18, col=C_WHITE)
        txb(sl, title, x+1.5, y+0.08, 6.2, 0.65, sz=14, bold=True, col=C_WHITE)
        txb(sl, body,  x+0.2, y+1.0,  7.5, 3.8,  sz=12.5, col=C_LGRAY)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 10 – DETAILED WORKING PROCESS
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Detailed Working Process","From Registration to Live Leaderboard")
    steps = [
        ("1","Register / Login","User signs up with email + phone. Password hashed with bcrypt. JWT token issued.",C_ACCENT),
        ("2","Browse Matches","User browses CS2/Valorant/BGMI matches filtered by status and game tab.",C_PURPLE),
        ("3","Build Fantasy Team","Select 5 players within 100-credit budget. Assign Captain (2×) and Vice-Captain (1.5×).",C_GREEN),
        ("4","Join Contest","Enter a free or paid contest for the chosen match. Team locked at match start.",C_ORANGE),
        ("5","Match Goes Live","Admin sets status → live. Admin enters per-player stats in real time.",C_RED),
        ("6","Score Calculation","node-cron triggers ScoringService every 2 min. Points calculated per player.",C_GOLD),
        ("7","WebSocket Broadcast","broadcastLeaderboardUpdate() emits ranked list to all match-room subscribers.",RGBColor(0x00,0xCC,0x99)),
        ("8","Results & Prizes","Match completed. Final ranks locked. Prize distribution executed by admin.",RGBColor(0xCC,0x44,0x99)),
    ]
    for i,(num,title,body,col) in enumerate(steps):
        row = i//4; col_i = i%4
        x = 0.4 + col_i*8.35; y = 3.5 + row*5.5
        rect(sl, x, y, 7.9, 5.1, C_NAVY)
        rect(sl, x, y, 2.0, 5.1, col)
        txb(sl, num, x, y, 2.0, 1.5, sz=32, bold=True, col=C_WHITE, align=PP_ALIGN.CENTER)
        txb(sl, title, x+2.1, y+0.2, 5.6, 0.9, sz=13, bold=True, col=C_WHITE)
        txb(sl, body,  x+2.1, y+1.2, 5.6, 3.7, sz=11.5, col=C_LGRAY)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 11 – SCORING ALGORITHM
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Scoring Algorithm","Transparent, Rule-Based Fantasy Point System")
    # left panel
    rect(sl, 0.5, 3.2, 15.8, 16.3, C_NAVY)
    txb(sl,"Point Rules",0.8,3.35,15.2,0.9,sz=16,bold=True,col=C_GOLD)
    rules = [
        ("Kill","+2 pts",C_GREEN),("Death","−1 pt",C_RED),("Assist","+2 pts",C_GREEN),
        ("MVP (CS2)","+5 pts",C_ACCENT),("First Blood (Val)","+3 pts",C_ACCENT),
        ("Clutch Win (Val)","+5 pts",C_ACCENT),("Spike Plant (Val)","+2 pts",C_ACCENT),
        ("Spike Defuse (Val)","+3 pts",C_ACCENT),
        ("Rating > 1.3 (CS2)","+10 bonus",C_GOLD),("Rating < 0.9 (CS2)","−5 pts",C_RED),
    ]
    for i,(action,pts,col) in enumerate(rules):
        y = 4.4 + i*1.45
        rect(sl, 0.7, y, 10.5, 1.2, C_CARD)
        txb(sl, action, 0.9, y+0.15, 7.5, 0.9, sz=12.5, col=C_WHITE)
        txb(sl, pts, 8.5, y+0.15, 2.5, 0.9, sz=12.5, bold=True, col=col, align=PP_ALIGN.RIGHT)

    # right panel – multipliers + example
    rect(sl, 17.2, 3.2, 15.8, 7.0, C_NAVY)
    txb(sl,"Role Multipliers",17.5,3.35,15.2,0.9,sz=16,bold=True,col=C_GOLD)
    mults = [("Captain (C)","× 2.0",C_GOLD),("Vice-Captain (VC)","× 1.5",C_LGRAY),("Regular","× 1.0",C_WHITE)]
    for i,(role,mult,col) in enumerate(mults):
        y = 4.4 + i*1.75
        rect(sl, 17.4, y, 15.2, 1.5, C_CARD)
        txb(sl, role, 17.6, y+0.2, 8, 1.0, sz=14, col=C_WHITE)
        txb(sl, mult, 24.2, y+0.2, 3, 1.0, sz=18, bold=True, col=col, align=PP_ALIGN.RIGHT)

    rect(sl, 17.2, 10.5, 15.8, 9.0, C_NAVY)
    txb(sl,"Example — Valorant Captain",17.5,10.65,15.2,0.9,sz=15,bold=True,col=C_GOLD)
    ex = [("20 kills × 2","= 40 pts"),("8 deaths × −1","= −8 pts"),
          ("5 assists × 2","= 10 pts"),("3 first bloods × 3","= 9 pts"),
          ("2 clutches × 5","= 10 pts"),("Base total","= 61 pts"),
          ("Captain (×2)","= 122 pts ★")]
    for i,(left,right) in enumerate(ex):
        y = 11.75 + i*1.05
        txb(sl, left,  17.5, y, 8.5, 0.9, sz=11.5, col=C_LGRAY)
        txb(sl, right, 26.5, y, 6.3, 0.9, sz=12, bold=(i==6), col=C_GOLD if i==6 else C_WHITE, align=PP_ALIGN.RIGHT)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 12 – DATABASE DESIGN
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Database Design","MongoDB Collections and Relationships")
    collections = [
        ("users","username · email · phone · password (bcrypt) · role · stats"),
        ("matches","game · team1/2 · tournament · status · format · matchContests[]"),
        ("bgmimatches","teams[] · standings (placement, kills, points) · map"),
        ("players","nickname · game · team · role · price · stats (K/D/A …)"),
        ("teams","name · game · region · achievements[] · socialMedia"),
        ("tournaments","name · game · tier · prizePool · startDate · endDate"),
        ("contests","entryFee · prizePool · participants[] · prizeDistribution[]"),
        ("fantasyteams","user · match · players[5] · captain · VC · totalPoints · rank"),
        ("matchplayerstats","match · player · kills · deaths · assists · fantasyPoints"),
        ("articles","title · content · author · tags · image · publishedAt"),
        ("complaints","userId · subject · body · status · adminResponse"),
        ("pendingregistrations","phone · otp · otpExpiry"),
    ]
    col_count = 2
    per_col = 6
    for i,(col_name,fields) in enumerate(collections):
        col_i = i // per_col
        row_i = i % per_col
        x = 0.5 + col_i*16.7; y = 3.4 + row_i*2.55
        rect(sl, x, y, 15.8, 2.25, C_CARD)
        rect(sl, x, y, 2.8, 2.25, C_ACCENT)
        txb(sl, col_name, x+0.15, y+0.55, 2.55, 1.1, sz=11, bold=True,
            col=C_WHITE, align=PP_ALIGN.CENTER)
        txb(sl, fields, x+3.0, y+0.4, 12.6, 1.5, sz=10.5, col=C_LGRAY)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 13 – UI/UX
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"UI/UX Overview","Key Screens and Design Philosophy")
    # design principles left
    rect(sl, 0.5, 3.2, 13.5, 16.3, C_NAVY)
    txb(sl,"Design Principles",0.8,3.38,13,0.9,sz=15,bold=True,col=C_GOLD)
    principles = ["Mobile-First (360–414 px target)","Dark gaming-centric theme + electric blue",
                  "Progressive Disclosure — step-by-step UX","Real-time feedback with toast notifications",
                  "Lazy-loaded routes → fast initial load","Cloudinary CDN for all media assets"]
    for i,p in enumerate(principles):
        txb(sl,f"▸  {p}", 0.7, 4.5+i*2.15, 13, 1.9, sz=12.5, col=C_WHITE)
    # screens right
    rect(sl, 14.8, 3.2, 18.5, 16.3, C_NAVY)
    txb(sl,"Key Screens",15.1,3.38,18,0.9,sz=15,bold=True,col=C_GOLD)
    screens = [
        ("/home","Landing — featured matches, news"),
        ("/matches","Tabbed CS2 | Valorant | BGMI match list"),
        ("/matches/:id","Match detail + contests + player pool"),
        ("/team-builder/:id","Interactive 5-player selection + budget"),
        ("/contests","Contest cards — entry fee, prize pool"),
        ("/contests/:id","Live leaderboard (WebSocket updates)"),
        ("/leaderboard/:id","Full ranked participant list"),
        ("/dashboard","Personal stats — ranks, points, contests"),
        ("/admin","Comprehensive admin management panel"),
        ("/news","Esports articles + live scores ticker"),
    ]
    for i,(route,desc) in enumerate(screens):
        y = 4.5 + i*1.47
        rect(sl, 15.0, y, 18.0, 1.28, C_CARD)
        txb(sl, route, 15.2, y+0.12, 5.5, 1.0, sz=10.5, bold=True, col=C_ACCENT)
        txb(sl, desc,  20.9, y+0.12, 11.8, 1.0, sz=10.5, col=C_WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 14 – IMPLEMENTATION
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Implementation Details","Project Structure and Security Stack")
    # backend
    rect(sl, 0.5, 3.2, 15.8, 16.3, C_NAVY)
    txb(sl,"Backend Structure",0.8,3.38,15.2,0.9,sz=15,bold=True,col=C_GOLD)
    be_files = [("server.js","Express init, Socket.io, cron jobs, routes"),
                ("controllers/","Business logic per resource"),
                ("models/","Mongoose schemas — 14 collections"),
                ("routes/","Express Router URL → controller"),
                ("middlewares/","JWT auth, errorHandler, validator"),
                ("services/","scoringService, smsService"),
                ("utils/constants.js","Scoring rules, budget limits, enums")]
    for i,(f,d) in enumerate(be_files):
        y = 4.5+i*2.05
        rect(sl,0.7,y,15.2,1.8,C_CARD)
        txb(sl,f, 0.9,y+0.18,5.0,1.3,sz=11,bold=True,col=C_ACCENT)
        txb(sl,d, 6.1,y+0.18,9.5,1.3,sz=11,col=C_WHITE)
    # security right
    rect(sl,17.2,3.2,15.8,7.8,C_NAVY)
    txb(sl,"Security Stack",17.5,3.38,15.2,0.9,sz=15,bold=True,col=C_GOLD)
    sec = ["JWT — stateless session tokens","bcrypt — password hashing (10 rounds)",
           "Helmet.js — 14 HTTP security headers","CORS — strict origin allowlist in prod",
           "Rate Limiter — 100 req / 15 min / IP","Input Validation — express-validator"]
    for i,s in enumerate(sec):
        txb(sl,f"🔒  {s}", 17.4, 4.5+i*1.58, 15.4, 1.4, sz=12.5, col=C_WHITE)
    # frontend
    rect(sl,17.2,11.3,15.8,8.2,C_NAVY)
    txb(sl,"Frontend Structure",17.5,11.48,15.2,0.9,sz=15,bold=True,col=C_GOLD)
    fe_files = [("app-routing.module.ts","Lazy routes + guards"),
                ("services/auth.service.ts","Login, JWT, user state"),
                ("services/api.service.ts","HTTP wrapper for all REST calls"),
                ("services/socket.service.ts","Socket.io client connection"),
                ("guards/","AuthGuard, AdminGuard"),
                ("components/","team-builder, leaderboard, admin…")]
    for i,(f,d) in enumerate(fe_files):
        y = 12.6+i*1.1
        txb(sl,f"▸  {f}  —  {d}", 17.4, y, 15.4, 1.0, sz=11, col=C_LGRAY)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 15 – TESTING AND RESULTS
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Testing & Results","Validated End-to-End with 100% Pass Rate")
    # test cases table
    rect(sl,0.5,3.2,32.8,16.3,C_NAVY)
    rect(sl,0.5,3.2,32.8,1.0,C_ACCENT)
    for i,(lbl,w) in enumerate(zip(["TC-ID","Test Case","Input","Expected","Result"],
                                    [2.5,8.0,8.5,9.8,3.0])):
        x = 0.5 + sum([2.5,8.0,8.5,9.8,3.0][:i])
        txb(sl, lbl, x+0.1, 3.28, w-0.2, 0.85, sz=11, bold=True, col=C_WHITE)
    tcs = [
        ("TC-001","User Registration","POST /api/auth/register — valid","201 Created + JWT token","✅ Pass"),
        ("TC-002","Duplicate Email","POST with existing email","400 Bad Request error","✅ Pass"),
        ("TC-003","Valid Login","POST correct credentials","200 + JWT token","✅ Pass"),
        ("TC-004","Invalid Login","POST wrong password","401 Unauthorized","✅ Pass"),
        ("TC-005","Create Fantasy Team","5 players + C + VC, budget OK","201 team created","✅ Pass"),
        ("TC-006","Over-Budget Team","Players sum > 100 credits","400 budget exceeded","✅ Pass"),
        ("TC-007","Missing Captain","Team submitted without C flag","400 validation error","✅ Pass"),
        ("TC-008","Join Contest","POST join with valid teamId","200 participant added","✅ Pass"),
        ("TC-009","Duplicate Contest Join","Same user joins twice","400 already joined","✅ Pass"),
        ("TC-010","Live Leaderboard","Admin enters stats → 2 min wait","WS update received","✅ Pass"),
        ("TC-011","Admin Without Role","GET /api/admin — regular JWT","403 Forbidden","✅ Pass"),
        ("TC-012","Rate Limit","101 requests / 15 min (prod)","429 Too Many Requests","✅ Pass"),
    ]
    widths = [2.5,8.0,8.5,9.8,3.0]
    for ri, row in enumerate(tcs):
        bg_c = C_CARD if ri%2==0 else C_NAVY
        x_start = 0.5
        for ci, (val, w) in enumerate(zip(row, widths)):
            x = x_start + sum(widths[:ci])
            y = 4.35 + ri*1.18
            col = RGBColor(0x44,0xFF,0x88) if "Pass" in val else C_WHITE
            txb(sl, val, x+0.1, y, w-0.2, 1.05, sz=10, col=col)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 16 – ADVANTAGES AND LIMITATIONS
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Advantages & Limitations","Strengths and Known Constraints")
    rect(sl,0.5,3.2,15.8,16.3,RGBColor(0x0A,0x22,0x12))
    rect(sl,0.5,3.2,15.8,1.2,C_GREEN)
    txb(sl,"✅  Advantages",0.8,3.28,15.2,1.0,sz=16,bold=True,col=C_WHITE)
    advs = ["CS2 + Valorant + BGMI — unified platform",
            "Cross-platform: Web + Android + iOS",
            "Real-time WebSocket leaderboards",
            "Transparent, verifiable scoring rules",
            "Production-grade security stack",
            "Admin Excel import — fast data entry",
            "Scalable: stateless API + Atlas auto-scale",
            "Extensible — new game = new model + route"]
    for i,a in enumerate(advs):
        txb(sl,f"▸  {a}", 0.7, 4.7+i*1.85, 15.4, 1.65, sz=13, col=RGBColor(0x88,0xFF,0xAA))
    rect(sl,17.2,3.2,15.8,16.3,RGBColor(0x2A,0x10,0x10))
    rect(sl,17.2,3.2,15.8,1.2,C_RED)
    txb(sl,"⚠  Current Limitations",17.5,3.28,15.2,1.0,sz=16,bold=True,col=C_WHITE)
    lims = ["Manual stat entry (scraping disabled)",
            "No payment gateway (fees tracked logically)",
            "No push notifications (FCM not yet added)",
            "Single-server WebSocket (no Redis yet)",
            "Limited player form analytics / charts",
            "BGMI scoring less granular than CS2/Val"]
    for i,l in enumerate(lims):
        txb(sl,f"▸  {l}", 17.4, 4.7+i*1.85, 15.4, 1.65, sz=13, col=RGBColor(0xFF,0xAA,0xAA))

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 17 – FUTURE SCOPE
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Future Scope","Planned Enhancements and Roadmap")
    future = [
        ("💳","Payment Gateway","Razorpay / Stripe for real-money entry, wallet & prize payout"),
        ("🎮","Official Game API","Riot Games / FACEIT API for automated real-time stat ingestion"),
        ("🤖","ML Recommendations","AI player-pick suggestions based on historical fantasy data"),
        ("🔔","Push Notifications","Firebase FCM for match-start alerts and prize-win notifications"),
        ("👥","Social Features","Friend lists, private contests, head-to-head challenges"),
        ("📊","Advanced Analytics","Player form charts, win-rate histograms, points history"),
        ("🌐","Multi-Language","Hindi + regional language localisation for tier-2/3 audience"),
        ("⚡","Redis + Socket.io","Horizontal WebSocket scaling via Redis pub/sub adapter"),
        ("🏅","Tournament Bracket","Visual bracket & round-robin schedule manager"),
        ("🧠","AI Chat Assistant","Conversational AI for optimal fantasy team building"),
        ("🎬","Video Highlights","Embedded match clips from YouTube / Twitch"),
        ("🎁","Referral Program","In-app coins for referrals and engagement milestones"),
    ]
    for idx,(icon,title,body) in enumerate(future):
        row=idx//4; col_i=idx%4
        x=0.4+col_i*8.35; y=3.5+row*3.7
        rect(sl,x,y,7.9,3.4,C_CARD)
        txb(sl,icon,  x+0.15,y+0.15,1.2,0.85,sz=18,col=C_WHITE)
        txb(sl,title, x+1.5, y+0.17,6.2,0.8, sz=13,bold=True,col=C_GOLD)
        txb(sl,body,  x+0.2, y+1.1, 7.5,2.1, sz=11,col=C_LGRAY)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 18 – REAL-WORLD APPLICATIONS
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Real-World Applications","Where KhelCircle Can Be Deployed")
    apps = [
        ("🎮","Esports Events","Overlay fantasy competitions on live esports tournaments to boost viewer engagement."),
        ("💰","Commercial Platform","Standalone revenue-generating fantasy product in India's growing gaming market."),
        ("🏫","College Esports","Universities host inter-college tournaments with integrated fantasy leagues for students."),
        ("🏢","Corporate Events","Gamify internal company esports events with fantasy scoring and leaderboards."),
        ("📈","Analytics Tool","The player-stats DB and scoring engine repurposed for esports performance scouting."),
        ("🎓","Portfolio Project","Demonstrates full-stack, real-time, and cross-platform mobile development expertise."),
    ]
    for idx,(icon,title,body) in enumerate(apps):
        row=idx//2; col_i=idx%2
        x=0.5+col_i*16.7; y=3.5+row*4.0
        rect(sl,x,y,15.8,3.6,C_CARD)
        rect(sl,x,y,2.5,3.6,C_ACCENT)
        txb(sl,icon, x,    y,    2.5,3.6,sz=30,align=PP_ALIGN.CENTER,col=C_WHITE)
        txb(sl,title,x+2.7,y+0.3,12.8,0.9,sz=15,bold=True,col=C_GOLD)
        txb(sl,body, x+2.7,y+1.4,12.8,2.0,sz=12.5,col=C_WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 19 – CONCLUSION
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    header(sl,"Conclusion","KhelCircle — Ready to Compete")
    rect(sl,0.5,3.2,32.8,11.0,C_NAVY)
    concl = [
        "KhelCircle delivers a production-ready Fantasy Esports Platform covering CS2, Valorant, and BGMI in one unified application.",
        "The architecture — Angular 17 + Ionic, Node.js + Express, MongoDB, Socket.io — is modern, scalable, and fully extensible.",
        "Real-time scoring, transparent rules, cross-platform deployment, and a comprehensive admin panel set it apart from existing solutions.",
        "All 12 test cases passed. Android APK deployed. APIs secured. Leaderboards broadcast live via WebSocket.",
        "Future roadmap includes payment gateway, official game API integration, ML recommendations, and push notifications.",
    ]
    for i, line in enumerate(concl):
        y = 3.6 + i*2.1
        rect(sl, 0.7, y, 0.8, 1.75, C_ACCENT)
        txb(sl, str(i+1), 0.7, y, 0.8, 1.75, sz=16, bold=True,
            col=C_WHITE, align=PP_ALIGN.CENTER)
        txb(sl, line, 1.7, y+0.2, 31.2, 1.4, sz=13, col=C_WHITE)
    txb(sl,"\"Transforming esports spectators into active, skill-based competitors.\"",
        0.5, 14.5, 32.8, 2.0, sz=18, italic=True, col=C_GOLD, align=PP_ALIGN.CENTER)
    txb(sl,"Platform ready for college presentation · evaluation · viva · commercial deployment",
        0.5, 17.0, 32.8, 1.5, sz=13, col=C_MGRAY, align=PP_ALIGN.CENTER)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 20 – Q&A / THANK YOU
    # ══════════════════════════════════════════════════════════════════════
    sl = prs.slides.add_slide(blank); bg(sl,0x0D,0x1B,0x2A); add_transition(sl)
    rect(sl,0,0,33.87,0.55,C_ACCENT)
    rect(sl,0,18.9,33.87,0.55,C_ACCENT)
    # big glow circle
    glow=sl.shapes.add_shape(9,cm(10.5),cm(3.5),cm(13),cm(13))
    glow.fill.solid(); glow.fill.fore_color.rgb=RGBColor(0x05,0x25,0x50)
    glow.line.fill.background()
    txb(sl,"🎮", 0, 4.5, 33.87, 3.5, sz=72, align=PP_ALIGN.CENTER)
    txb(sl,"Thank You!", 0, 8.0, 33.87, 3.2,
        sz=72, bold=True, col=C_GOLD, align=PP_ALIGN.CENTER)
    txb(sl,"Questions & Answers", 0, 11.2, 33.87, 2.0,
        sz=30, col=C_WHITE, align=PP_ALIGN.CENTER)
    txb(sl,"KhelCircle — Fantasy Esports Platform", 0, 13.3, 33.87, 1.5,
        sz=18, italic=True, col=C_MGRAY, align=PP_ALIGN.CENTER)
    # tech tag strip
    rect(sl, 3.0, 15.2, 28.0, 1.8, C_CARD)
    txb(sl,"Angular 17  ·  Node.js  ·  MongoDB  ·  Socket.io  ·  Ionic  ·  Capacitor  ·  Cloudinary",
        3.2, 15.3, 27.6, 1.5, sz=13.5, col=C_ACCENT, align=PP_ALIGN.CENTER)

    out_path = os.path.join(OUT, "KhelCircle_Presentation.pptx")
    prs.save(out_path)
    print(f"PowerPoint saved: {out_path}")

if __name__=="__main__":
    build()
