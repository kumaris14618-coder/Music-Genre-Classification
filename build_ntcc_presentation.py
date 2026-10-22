import os
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Theme Colors
    BG_DARK = RGBColor(15, 23, 42)       # Slate 900
    BG_LIGHT = RGBColor(248, 250, 252)   # Slate 50
    CARD_BG = RGBColor(255, 255, 255)    # White
    CARD_BORDER = RGBColor(226, 232, 240)# Slate 200
    PRIMARY = RGBColor(30, 58, 138)      # Blue 900 (Amity Navy)
    ACCENT_BLUE = RGBColor(37, 99, 235)  # Blue 600
    ACCENT_CYAN = RGBColor(6, 182, 212)  # Cyan 500
    ACCENT_GOLD = RGBColor(217, 119, 6)  # Amber 600
    TEXT_DARK = RGBColor(30, 41, 59)     # Slate 800
    TEXT_MUTED = RGBColor(100, 116, 139) # Slate 500
    TEXT_LIGHT = RGBColor(255, 255, 255) # White

    def set_slide_background(slide, color):
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_header(slide, category, title):
        # Header banner shape
        banner = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1)
        )
        banner.fill.solid()
        banner.fill.fore_color.rgb = PRIMARY
        banner.line.color.rgb = ACCENT_CYAN
        banner.line.width = Pt(1.5)

        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.7), Inches(0.85))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        
        # Category / Breadcrumb
        p_cat = tf.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = ACCENT_CYAN

        # Title
        p_title = tf.add_paragraph()
        p_title.text = title
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_LIGHT

    def add_card(slide, left, top, width, height, title="", bg_color=CARD_BG, border_color=CARD_BORDER):
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
        )
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)

        if title:
            tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), Inches(0.45))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = PRIMARY

        return card

    # ==========================================
    # SLIDE 1: TITLE / COVER SLIDE
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, BG_DARK)

    # Accent decorative bars
    top_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = ACCENT_CYAN
    top_bar.line.fill.background()

    # University header card
    uni_card = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(0.6), Inches(10.333), Inches(1.0))
    uni_card.fill.solid()
    uni_card.fill.fore_color.rgb = RGBColor(30, 41, 59)
    uni_card.line.color.rgb = ACCENT_CYAN
    uni_card.line.width = Pt(1)

    tb = s1.shapes.add_textbox(Inches(1.5), Inches(0.65), Inches(10.333), Inches(0.9))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "AMITY SCHOOL OF ENGINEERING AND TECHNOLOGY"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(226, 232, 240)
    p.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = "AMITY UNIVERSITY JHARKHAND, RANCHI"
    p2.font.size = Pt(14)
    p2.font.bold = True
    p2.font.color.rgb = ACCENT_CYAN
    p2.alignment = PP_ALIGN.CENTER

    # Project Title Box
    title_box = s1.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(11.333), Inches(2.2))
    tf = title_box.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "NON-TEACHING CREDIT COURSE (NTCC) REPORT"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GOLD
    p.alignment = PP_ALIGN.CENTER

    p = tf.add_paragraph()
    p.text = "Music Genre Classification using Deep Learning\nand Python"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT
    p.alignment = PP_ALIGN.CENTER

    p = tf.add_paragraph()
    p.text = "An End-to-End Deep CNN Acoustic Framework with Mel-Spectrogram Processing and Streamlit Web Application"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(148, 163, 184)
    p.alignment = PP_ALIGN.CENTER

    # Details Cards (Student & Guide)
    # Student Card
    s_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.8), Inches(4.4), Inches(4.6), Inches(2.3))
    s_card.fill.solid()
    s_card.fill.fore_color.rgb = RGBColor(30, 41, 59)
    s_card.line.color.rgb = ACCENT_BLUE
    s_card.line.width = Pt(1.5)

    tb = s1.shapes.add_textbox(Inches(2.0), Inches(4.55), Inches(4.2), Inches(2.0))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "SUBMITTED BY:"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN
    
    p = tf.add_paragraph()
    p.text = "Karan"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT

    p = tf.add_paragraph()
    p.text = "Program: Bachelor of Technology (B.Tech - CSE)\nSemester: 3rd Semester\nBatch: 2024 - 2028\nDepartment of Computer Science & Engineering"
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(203, 213, 225)

    # Guide Card
    g_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(4.4), Inches(4.6), Inches(2.3))
    g_card.fill.solid()
    g_card.fill.fore_color.rgb = RGBColor(30, 41, 59)
    g_card.line.color.rgb = ACCENT_GOLD
    g_card.line.width = Pt(1.5)

    tb = s1.shapes.add_textbox(Inches(7.1), Inches(4.55), Inches(4.2), Inches(2.0))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "UNDER THE GUIDANCE OF:"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GOLD

    p = tf.add_paragraph()
    p.text = "Mr. Mohan Kumar Dehury"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT

    p = tf.add_paragraph()
    p.text = "Designation: Assistant Professor\nDepartment of Computer Science & Engineering\nAmity School of Engineering and Technology\nAmity University Jharkhand, Ranchi"
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(203, 213, 225)

    # Bottom bar
    bot_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.3), Inches(13.333), Inches(0.2))
    bot_bar.fill.solid()
    bot_bar.fill.fore_color.rgb = PRIMARY
    bot_bar.line.fill.background()

    # ==========================================
    # SLIDE 2: DECLARATION & CERTIFICATE
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2, BG_LIGHT)
    add_header(s2, "Academic Credentials", "Declaration and Faculty Certificate")

    # Left: Declaration
    add_card(s2, Inches(0.8), Inches(1.3), Inches(5.6), Inches(5.7), "STUDENT DECLARATION")
    tb = s2.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.2), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "I hereby declare that the NTCC Project Report titled \"Music Genre Classification using Deep Learning and Python\", submitted by me to the Department of Computer Science & Engineering, Amity School of Engineering and Technology, Amity University Jharkhand, in partial fulfillment of the requirements for the award of the degree of Bachelor of Technology, is an authentic record of original research work carried out by me under the esteemed supervision of Mr. Mohan Kumar Dehury.\n\nI further affirm that this work has not previously formed the basis for the award of any degree, diploma, associate-ship, or other similar academic title to this or any other university or institution."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    p = tf.add_paragraph()
    p.text = "\n\nStudent Name: Karan\nDegree: B.Tech (Computer Science & Engineering)\nAmity University Jharkhand, Ranchi\nDate: 22nd October 2026"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = PRIMARY

    # Right: Certificate
    add_card(s2, Inches(6.9), Inches(1.3), Inches(5.6), Inches(5.7), "FACULTY SUPERVISOR CERTIFICATE")
    tb = s2.shapes.add_textbox(Inches(7.1), Inches(1.85), Inches(5.2), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "This is to certify that the NTCC Project Report entitled \"Music Genre Classification using Deep Learning and Python\", submitted by Karan, student of Bachelor of Technology in Computer Science & Engineering, Amity School of Engineering and Technology, Amity University Jharkhand, is a faithful record of project work carried out under my personal guidance and supervision.\n\nThe project demonstrates technical rigor, sound experimental methodology in audio digital signal processing and deep convolutional neural networks, and successful software implementation.\n\nTo the best of my knowledge, this work has not been submitted elsewhere for any academic degree or diploma."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    p = tf.add_paragraph()
    p.text = "\n\nMr. Mohan Kumar Dehury\nAssistant Professor, Dept. of Computer Science & Engineering\nAmity School of Engineering and Technology\nAmity University Jharkhand, Ranchi"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = PRIMARY

    # ==========================================
    # SLIDE 3: ACKNOWLEDGEMENT
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3, BG_LIGHT)
    add_header(s3, "Formal Recognition", "Acknowledgements")

    # Card 1: Guide
    add_card(s3, Inches(0.8), Inches(1.4), Inches(11.7), Inches(1.6), "FACULTY GUIDANCE & MENTORSHIP")
    tb = s3.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "I express my profound gratitude and deepest indebtedness to my esteemed project guide, Mr. Mohan Kumar Dehury, Assistant Professor, Department of Computer Science & Engineering, Amity University Jharkhand, for his invaluable guidance, inspiring encouragement, and technical recommendations throughout the conception and implementation of this project."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # Card 2: Institution
    add_card(s3, Inches(0.8), Inches(3.2), Inches(11.7), Inches(1.6), "INSTITUTIONAL SUPPORT & DEPARTMENT")
    tb = s3.shapes.add_textbox(Inches(1.0), Inches(3.7), Inches(11.3), Inches(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "I extend my sincere thanks to the Department of Computer Science & Engineering and the Amity School of Engineering and Technology, Amity University Jharkhand, for providing modern computational infrastructure, laboratory resources, and an academically enriching environment that made this machine learning endeavor feasible."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # Card 3: Peers & Family
    add_card(s3, Inches(0.8), Inches(5.0), Inches(11.7), Inches(1.8), "PEERS, FAMILY & SCIENTIFIC COMMUNITY")
    tb = s3.shapes.add_textbox(Inches(1.0), Inches(5.5), Inches(11.3), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "I am deeply thankful to my family members and peers whose unwavering moral support, constructive feedback, and encouragement sustained my commitment during challenging phases of model tuning.\n\nFinally, I acknowledge the global open-source AI and MIR research community behind Librosa, TensorFlow, and the GTZAN dataset, whose pioneering open-access contributions continue to propel audio artificial intelligence forward."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 4: TABLE OF CONTENTS / OUTLINE
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4, BG_LIGHT)
    add_header(s4, "Project Roadmap", "Presentation Table of Contents")

    sections = [
        ("01", "Abstract & Executive Summary", "Project scope, high-level approach & core impact"),
        ("02", "Introduction & Motivation", "Audio AI revolution, problem statement & objectives"),
        ("03", "Literature Review", "Evolution from handcrafted MFCCs to Deep Spectrogram CNNs"),
        ("04", "Methodology & Architecture", "GTZAN dataset, Librosa Mel-Spectrograms & 5-Block CNN"),
        ("05", "System Analysis & Challenges", "Acoustic polyphony, genre overlap & Python ecosystem"),
        ("06", "Results & Empirical Findings", "98.56% Train / 89.25% Val Acc, Confusion Matrix & Streamlit App"),
        ("07", "Limitations & Future Scope", "Audio compression, multi-label tagging & Audio Transformers"),
        ("08", "Conclusions & References", "Key findings, engineering takeaways & academic literature"),
        ("09", "Weekly Progress Reports (WPR)", "5-week comprehensive NTCC execution logs & deliverables")
    ]

    for idx, (num, title, desc) in enumerate(sections):
        row = idx // 3
        col = idx % 3
        left = Inches(0.8 + col * 4.0)
        top = Inches(1.4 + row * 1.85)

        card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(3.7), Inches(1.65))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.2)

        # Number badge
        badge = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.2), top + Inches(0.2), Inches(0.55), Inches(0.35))
        badge.fill.solid()
        badge.fill.fore_color.rgb = PRIMARY
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = num
        p_b.font.size = Pt(12)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s4.shapes.add_textbox(left + Inches(0.85), top + Inches(0.15), Inches(2.7), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(10)
        p_d.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 5: ABSTRACT
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5, BG_LIGHT)
    add_header(s5, "Executive Summary", "Project Abstract")

    add_card(s5, Inches(0.8), Inches(1.4), Inches(8.0), Inches(5.5), "RESEARCH ABSTRACT")
    tb = s5.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(7.6), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "Music genre classification represents an essential pillar of Music Information Retrieval (MIR), serving as the foundational engine for automated tagging, personalized recommendation, and catalog indexing in modern streaming platforms.\n\nThis project presents an end-to-end deep learning framework developed in Python that converts one-dimensional raw audio signals into two-dimensional Mel-Frequency Spectrograms, treating acoustic classification as an advanced computer vision task.\n\nUtilizing the benchmark GTZAN dataset (1,000 tracks across 10 distinct genres: blues, classical, country, disco, hiphop, jazz, metal, pop, reggae, and rock), the audio signals are chunked into 4-second windows with 2-second overlap and transformed via Librosa into 150x150 decibel-scaled Mel-filterbank representations.\n\nA custom 5-block Deep Convolutional Neural Network (CNN) is designed, incorporating progressive filter depths (32 to 512), spatial max-pooling, and strategic dropout regularization (0.3 to 0.45) with Adam optimization.\n\nThe system attains 98.56% training accuracy and 89.25% validation accuracy, successfully deployed into an interactive Streamlit web application providing instant playback, dynamic spectrogram rendering, and multi-class genre probability distribution."
    p.font.size = Pt(11.5)
    p.font.color.rgb = TEXT_DARK

    # Right side metric highlights
    metrics = [
        ("10 Genres", "Multi-Class Classification", ACCENT_BLUE),
        ("98.56%", "Peak Training Accuracy", RGBColor(16, 185, 129)),
        ("89.25%", "Validation Accuracy", ACCENT_CYAN),
        ("Streamlit", "Interactive Web Deployment", ACCENT_GOLD)
    ]
    for idx, (val, lbl, col) in enumerate(metrics):
        top = Inches(1.4 + idx * 1.4)
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.1), top, Inches(3.4), Inches(1.2))
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_BG
        c.line.color.rgb = col
        c.line.width = Pt(2)

        tb = s5.shapes.add_textbox(Inches(9.2), top + Inches(0.12), Inches(3.2), Inches(0.95))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = val
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = col
        p.alignment = PP_ALIGN.CENTER

        p_lbl = tf.add_paragraph()
        p_lbl.text = lbl
        p_lbl.font.size = Pt(11)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = TEXT_DARK
        p_lbl.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 6: INTRODUCTION - BACKGROUND & MOTIVATION
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6, BG_LIGHT)
    add_header(s6, "Introduction", "The Evolution of Music & Audio AI")

    # Column 1: Music in Digital Age
    add_card(s6, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5), "THE DIGITAL MUSIC BOOM")
    tb = s6.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Exponential Growth of Audio Catalogs:\nModern music streaming services (Spotify, Apple Music, YouTube Music) host over 100+ million tracks, with more than 100,000 new songs uploaded daily.\n\n• The Failure of Manual Metadata Tagging:\nManual curation is subjective, error-prone, labor-intensive, and fundamentally incapable of scaling to contemporary global audio ingestion volumes.\n\n• Genre as a Cognitive Taxonomy:\nMusic genre is the primary organizational schema humans use to explore, search, and mentally index musical compositions.\n\n• Downstream Commercial Impact:\nAccurate genre categorization directly powers personalized recommendation systems, radio generation, targeted licensing, and mood playlist curation."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # Column 2: Cognitive vs Machine Perception
    add_card(s6, Inches(6.9), Inches(1.4), Inches(5.6), Inches(5.5), "PERCEPTUAL COMPLEXITY OF SOUND")
    tb = s6.shapes.add_textbox(Inches(7.1), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Human Hearing vs Machine Listening:\nHumans effortlessly identify a genre within seconds by holistically integrating timbre, tempo, rhythm, harmony, and instrumentation.\n\n• The High Dimensionality of Audio:\nA standard CD-quality 30-second audio clip (44.1 kHz, stereo) consists of over 2.6 million raw sample amplitudes, representing dense temporal complexity.\n\n• Bridging the Semantic Gap:\nRaw acoustic waveforms lack explicit spatial semantics. Bridging the 'semantic gap' between low-level pressure waves and high-level cultural genres requires advanced feature transformations.\n\n• The Paradigm Shift to Deep Learning:\nInstead of handcrafted statistical heuristics, deep learning models autonomously extract hierarchical representations of acoustic textures."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 7: INTRODUCTION - PROBLEM STATEMENT & OBJECTIVES
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7, BG_LIGHT)
    add_header(s7, "Introduction", "Problem Statement and Project Objectives")

    # Problem Statement Card
    add_card(s7, Inches(0.8), Inches(1.4), Inches(11.7), Inches(2.2), "FORMAL PROBLEM STATEMENT", border_color=RGBColor(239, 68, 68))
    tb = s7.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(11.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Given an arbitrary audio recording, design and implement an autonomous machine learning system capable of parsing continuous acoustic signals, isolating discriminative spectral-temporal patterns, and mapping the input into one of ten standardized musical genres with high statistical confidence, computational efficiency, and real-time inference viability without requiring human-engineered feature heuristics."
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = RGBColor(185, 28, 28)

    # 4 Core Objectives
    objs = [
        ("01", "Signal Preprocessing & Feature Extraction", "Transform raw time-domain audio waves into 2D Decibel Mel-Spectrograms (150x150) utilizing Librosa to simulate human psychoacoustic auditory perception."),
        ("02", "Deep CNN Architecture Design", "Construct and train a 5-block Deep Convolutional Neural Network with progressive feature maps (32 to 512 filters) and dropout regularization to minimize overfitting."),
        ("03", "Empirical Evaluation & Error Analysis", "Quantify performance across multi-class accuracy, loss convergence, confusion matrix, precision, recall, and F1-score across all 10 genres."),
        ("04", "Interactive Web Deployment", "Deliver a full-stack, user-friendly Streamlit web application enabling single-click audio file uploads, 30+ sample track library, and instant visual spectrogram rendering.")
    ]

    for idx, (num, heading, text) in enumerate(objs):
        left = Inches(0.8 + idx * 2.95)
        top = Inches(3.8)
        card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(2.8), Inches(3.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.2)

        badge = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.15), top + Inches(0.15), Inches(0.55), Inches(0.35))
        badge.fill.solid()
        badge.fill.fore_color.rgb = PRIMARY
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = num
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s7.shapes.add_textbox(left + Inches(0.15), top + Inches(0.55), Inches(2.5), Inches(2.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = heading
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_t = tf.add_paragraph()
        p_t.text = text
        p_t.font.size = Pt(10.5)
        p_t.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 8: LITERATURE REVIEW - HISTORICAL EVOLUTION
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8, BG_LIGHT)
    add_header(s8, "Literature Review", "Evolution from Handcrafted Descriptors to Deep Learning")

    eras = [
        ("ERA 1: 1990s - 2002", "Handcrafted Heuristics & Shallow ML", 
         "• Pioneered by Tzanetakis & Cook (2002) introducing the GTZAN dataset.\n• Relied on manually extracted acoustic features: Zero Crossing Rate (ZCR), Spectral Flux, Spectral Centroid, and Spectral Rolloff.\n• Classifiers: Gaussian Mixture Models (GMM), k-Nearest Neighbors (k-NN), Support Vector Machines (SVM).\n• Limitation: Poor generalization; accuracy plateaued at ~65-72%."),
        
        ("ERA 2: 2003 - 2013", "MFCCs & Statistical Pattern Recognition", 
         "• Focus shifted towards Mel-Frequency Cepstral Coefficients (MFCCs) capturing timbral envelope.\n• Integration of Chroma features for harmonic content and Tonnetz representations.\n• Advanced kernels in SVMs and ensemble Random Forests.\n• Limitation: Ignored continuous polyphonic temporal context and suffered from acoustic phase loss."),

        ("ERA 3: 2014 - Present", "Deep Representation Learning (Current Work)", 
         "• Breakthrough: Treating audio spectrograms as visual image representations.\n• Deep Convolutional Neural Networks (CNNs) automatically learn translation-invariant acoustic patterns.\n• Modern architectures leverage Mel-spectrogram chunking and data augmentation.\n• System Performance: Achieves 89-98% accuracy with real-time inference viability.")
    ]

    for idx, (era, title, content) in enumerate(eras):
        left = Inches(0.8 + idx * 4.0)
        top = Inches(1.4)
        card = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(3.7), Inches(5.5))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = PRIMARY if idx == 2 else CARD_BORDER
        card.line.width = Pt(2 if idx == 2 else 1.2)

        tb = s8.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), Inches(3.3), Inches(5.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = era
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = ACCENT_GOLD if idx == 2 else ACCENT_BLUE

        p_t = tf.add_paragraph()
        p_t.text = title
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = PRIMARY

        p_c = tf.add_paragraph()
        p_c.text = content
        p_c.font.size = Pt(11)
        p_c.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 9: LITERATURE REVIEW - SPECTROGRAMS & CNNS
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9, BG_LIGHT)
    add_header(s9, "Literature Review", "Audio as Images: The Power of Mel-Spectrograms")

    # Left Box
    add_card(s9, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5), "PSYCHOACOUSTICS & MEL FILTERBANKS")
    tb = s9.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• The Non-Linearity of Human Audition:\nHuman ear does not perceive sound pitch linearly. It exhibits high sensitivity to low-frequency variations (below 1 kHz) and diminishing discrimination at higher octaves.\n\n• The Mel Scale Formulation:\nm = 2595 * log10(1 + f / 700)\nMaps linear Hertz frequencies (f) into subjective pitch intervals (m), aligning neural model inputs with psychoacoustic biology.\n\n• Decibel (Log-Power) Compression:\nSound intensity follows the Weber-Fechner law of perceptual proportionality. Transforming energy to decibels (dB = 10 * log10(Power / Power_ref)) compresses high dynamic range spikes into stable numerical features."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # Right Box
    add_card(s9, Inches(6.9), Inches(1.4), Inches(5.6), Inches(5.5), "WHY CNNS EXCEL AT AUDIO PROCESSING")
    tb = s9.shapes.add_textbox(Inches(7.1), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• 2D Visual Audio Topology:\nA spectrogram is a 2D matrix where the X-axis represents Time, Y-axis represents Frequency, and Pixel Intensity denotes Energy.\n\n• Translation Invariance:\nJust as a visual CNN detects a cat whether it appears on the left or right of a photo, a 2D audio CNN recognizes a guitar riff or drum snare regardless of when it occurs in a 4-second chunk.\n\n• Hierarchical Feature Learning:\n- Low layers: Detect harmonic pitch stripes and percussive transients.\n- Middle layers: Learn chord progressions and rhythmic motifs.\n- Deep layers: Abstract timbral identity and distinctive genre signatures."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 10: METHODOLOGY - SYSTEM ARCHITECTURE PIPELINE
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10, BG_LIGHT)
    add_header(s10, "Methodology", "End-to-End System Architecture Pipeline")

    stages = [
        ("STAGE 1", "Audio Ingestion", "Upload .wav/.mp3 file or pick from 30+ track library. Sample rate preserved / resampled to 22,050 Hz.", ACCENT_BLUE),
        ("STAGE 2", "Temporal Windowing", "Split 30s audio into 4.0-second chunks with 2.0-second sliding overlap to maximize training data volume.", ACCENT_CYAN),
        ("STAGE 3", "Mel-Spectrogram", "Compute Short-Time Fourier Transform (STFT), apply 128 Mel filterbanks, convert to dB scale.", RGBColor(16, 185, 129)),
        ("STAGE 4", "Tensor Resizing", "Normalize dynamic range and resize matrices into standardized (150, 150, 1) single-channel tensors.", ACCENT_GOLD),
        ("STAGE 5", "5-Block Deep CNN", "Pass tensors through 10 Conv layers, 5 Max-Pool operations, Dropout (0.3/0.45) & Dense (1200).", PRIMARY),
        ("STAGE 6", "Prediction & UI", "Softmax computes 10-class probability vector. Streamlit renders prediction badge, confidence & bar chart.", RGBColor(147, 51, 234))
    ]

    for idx, (stg, title, desc, col) in enumerate(stages):
        row = idx // 3
        col_idx = idx % 3
        left = Inches(0.8 + col_idx * 4.0)
        top = Inches(1.4 + row * 2.8)

        card = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(3.7), Inches(2.55))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = col
        card.line.width = Pt(1.8)

        badge = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.2), top + Inches(0.2), Inches(1.1), Inches(0.35))
        badge.fill.solid()
        badge.fill.fore_color.rgb = col
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = stg
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s10.shapes.add_textbox(left + Inches(0.2), top + Inches(0.65), Inches(3.3), Inches(1.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_t = tf.add_paragraph()
        p_t.text = desc
        p_t.font.size = Pt(11)
        p_t.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 11: METHODOLOGY - SYSTEM REQUIREMENTS
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11, BG_LIGHT)
    add_header(s11, "Methodology", "System Specifications & Development Environment")

    # Hardware Card
    add_card(s11, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5), "HARDWARE CONFIGURATION")
    tb = s11.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Central Processing Unit (CPU):\nIntel Core i5 / Core i7 (6 cores, 12 threads @ 3.2 GHz) or Apple Silicon M-series for audio signal loading and parallel STFT computations.\n\n• Graphics Processing Unit (GPU):\nNVIDIA GeForce GTX 1050 / RTX 3060 (CUDA 11.8+ enabled) with Tensor Cores for accelerating 2D convolutional backpropagation.\n\n• Random Access Memory (RAM):\n16 GB DDR4/DDR5 system memory (minimum 8 GB required for in-memory batch caching of spectrogram tensors).\n\n• Storage Infrastructure:\nSolid State Drive (NVMe SSD) with >20 GB free space for storing the GTZAN audio archive, segmented cache, and trained model weights."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # Software Card
    add_card(s11, Inches(6.9), Inches(1.4), Inches(5.6), Inches(5.5), "SOFTWARE STACK & DEPENDENCIES")
    tb = s11.shapes.add_textbox(Inches(7.1), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Core Programming Environment:\nPython 3.10+ running in isolated Virtual Environment (venv / conda).\n\n• Deep Learning Framework:\nTensorFlow 2.10.0 / Keras for model architecture construction, layer stacking, and backpropagation optimization.\n\n• Audio Digital Signal Processing:\nLibrosa 0.10.1 / 1.0.0 for STFT transformation, Mel-scale filterbanks, and dB power conversions.\n\n• Web Application Framework:\nStreamlit for high-performance interactive web dashboard and real-time audio playback.\n\n• Scientific Computing & Visualization:\nNumPy 1.24.3, Pandas 2.1.0, Matplotlib 3.7.2, Seaborn 0.13.0, and Scikit-learn 1.3.0 for metric tracking and confusion matrix plotting."
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 12: METHODOLOGY - DATASET & PREPROCESSING
    # ==========================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12, BG_LIGHT)
    add_header(s12, "Methodology", "GTZAN Dataset & Mel-Spectrogram Extraction")

    # Left: Explanation
    add_card(s12, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5), "DATASET & PREPROCESSING PIPELINE")
    tb = s12.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• The GTZAN Benchmark Archive:\nKnown as the 'MNIST of Sound', comprising 1,000 distinct audio tracks (30 seconds each, 22,050 Hz, 16-bit mono wav files).\n\n• 10 Balanced Music Categories:\nBlues, Classical, Country, Disco, Hiphop, Jazz, Metal, Pop, Reggae, and Rock (100 tracks per class).\n\n• Windowing & Chunk Augmentation:\nDivided each 30s song into 4.0-second segments with 2.0-second overlap, expanding the dataset from 1,000 samples to over 13,000 training chunks.\n\n• Feature Representation:\nConverted each chunk into a 128-band Mel-Spectrogram with decibel scaling, then normalized and resized to (150, 150, 1) image tensors."
    p.font.size = Pt(11.5)
    p.font.color.rgb = TEXT_DARK

    # Right: Embedded Mel-Spectrogram Chart
    add_card(s12, Inches(6.9), Inches(1.4), Inches(5.6), Inches(5.5), "GENERATED MEL-SPECTROGRAM VISUALIZATION")
    if os.path.exists('chart_melspectrogram.png'):
        s12.shapes.add_picture('chart_melspectrogram.png', Inches(7.1), Inches(2.1), Inches(5.2), Inches(3.2))
    
    tb = s12.shapes.add_textbox(Inches(7.1), Inches(5.4), Inches(5.2), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Figure: Log-frequency Mel Spectrogram of 'blues.00000.wav'. Low frequencies capture foundational bass riffs while mid-to-high frequencies reflect vocal harmonics and cymbal resonance."
    p.font.size = Pt(10)
    p.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 13: METHODOLOGY - DEEP CNN ARCHITECTURE
    # ==========================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13, BG_LIGHT)
    add_header(s13, "Methodology", "Deep Convolutional Neural Network (CNN) Architecture")

    # Column 1: Feature Extraction Backbone
    add_card(s13, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5), "CONVOLUTIONAL FEATURE EXTRACTOR")
    tb = s13.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Input Layer: Shape (150, 150, 1) Mel-Spectrogram\n\n• Block 1: 2x Conv2D (32 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2)\nOutput shape: (75, 75, 32)\n\n• Block 2: 2x Conv2D (64 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2)\nOutput shape: (37, 37, 64)\n\n• Block 3: 2x Conv2D (128 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2) + Dropout(0.3)\nOutput shape: (18, 18, 128)\n\n• Block 4: 2x Conv2D (256 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2)\nOutput shape: (9, 9, 256)\n\n• Block 5: 2x Conv2D (512 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2) + Dropout(0.3)\nOutput shape: (4, 4, 512)"
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_DARK

    # Column 2: Classification Head & Regularization
    add_card(s13, Inches(6.9), Inches(1.4), Inches(5.6), Inches(5.5), "CLASSIFICATION HEAD & OPTIMIZATION")
    tb = s13.shapes.add_textbox(Inches(7.1), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Flattening Layer:\nFlattens 4x4x512 spatial feature maps into an 8,192-dimensional unified acoustic embedding vector.\n\n• Dense Fully Connected Layer:\nDense(1200 units, ReLU activation) capturing high-order non-linear combinations of musical motifs.\n\n• Strategic Regularization (Dropout 0.45):\nPrevents co-adaptation of hidden features and prevents overfitting across musical genres.\n\n• Output Softmax Classification:\nDense(10 units, Softmax activation) computing normalized multi-class probability scores summing to 1.0:\n\nP(Genre_i | X) = exp(z_i) / sum(exp(z_j))"
    p.font.size = Pt(11.5)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 14: METHODOLOGY - TRAINING & OPTIMIZATION
    # ==========================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14, BG_LIGHT)
    add_header(s14, "Methodology", "Model Training Strategy and Hyperparameters")

    params = [
        ("Optimization Algorithm", "Adam (Adaptive Moment Estimation)", "Utilizes exponentially decaying running averages of past gradients (beta1=0.9, beta2=0.999) for rapid parameter convergence."),
        ("Learning Rate", "0.0001 (1e-4)", "Selected to ensure smooth loss descent across deep 512-channel convolutional layers without gradient oscillations."),
        ("Loss Function", "Categorical Cross-Entropy", "L(y, y_hat) = - sum(y_i * log(y_hat_i)) — Standard loss for mutually exclusive 10-class genre distribution."),
        ("Batch Size & Epochs", "Batch Size: 32 | Epochs: 30", "Strikes ideal balance between vector parallelization in GPU RAM and frequent backpropagation updates."),
        ("Validation Strategy", "Train/Validation Split (80/20)", "Rigorous held-out validation partitioning to ensure genuine generalization on unseen audio compositions."),
        ("Checkpoint Persistence", "Trained_model.keras (86 MB)", "Directly serializes architecture weights, optimizer state, and computational graph for instant Streamlit inference.")
    ]

    for idx, (param, val, desc) in enumerate(params):
        row = idx // 2
        col = idx % 2
        left = Inches(0.8 + col * 5.9)
        top = Inches(1.4 + row * 1.85)

        card = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(5.6), Inches(1.65))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.2)

        tb = s14.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), Inches(5.2), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"{param}:  "
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = ACCENT_BLUE

        run = p.add_run()
        run.text = val
        run.font.bold = True
        run.font.color.rgb = PRIMARY

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(10.5)
        p_d.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 15: METHODOLOGY - EVALUATION PROTOCOLS
    # ==========================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_background(s15, BG_LIGHT)
    add_header(s15, "Methodology", "Testing Methodology & Evaluation Metrics")

    metrics_list = [
        ("Multi-Class Accuracy", "Accuracy = (Correct Predictions) / (Total Test Samples)", "Measures the global percentage of correctly classified audio chunks across all 10 genres on unseen validation data."),
        ("Precision (per genre)", "Precision = TP / (TP + FP)", "Quantifies reliability: when the model predicts 'Jazz', what fraction of those tracks are actually Jazz and not misclassified Blues?"),
        ("Recall / Sensitivity", "Recall = TP / (TP + FN)", "Quantifies coverage: of all true 'Metal' tracks in the validation set, what proportion did the model correctly identify?"),
        ("F1-Score (Harmonic Mean)", "F1 = 2 * (Precision * Recall) / (Precision + Recall)", "Provides an uncompromised balance between precision and recall, especially vital for evaluating overlapping acoustic classes."),
        ("Multi-Class Confusion Matrix", "C[i, j] = Count(True Genre i, Predicted Genre j)", "A 10x10 normalized matrix revealing specific inter-genre acoustic ambiguities (e.g., Rock vs Blues, Country vs Pop)."),
        ("Chunk-to-Song Inference", "Majority Voting & Mean Probability Averaging", "For a full 30s track, multiple 4s chunks are evaluated simultaneously, and class probabilities are averaged for the final verdict.")
    ]

    for idx, (title, formula, detail) in enumerate(metrics_list):
        row = idx // 2
        col = idx % 2
        left = Inches(0.8 + col * 5.9)
        top = Inches(1.4 + row * 1.85)

        card = s15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(5.6), Inches(1.65))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.2)

        tb = s15.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), Inches(5.2), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_f = tf.add_paragraph()
        p_f.text = formula
        p_f.font.size = Pt(10.5)
        p_f.font.bold = True
        p_f.font.color.rgb = ACCENT_CYAN

        p_d = tf.add_paragraph()
        p_d.text = detail
        p_d.font.size = Pt(10)
        p_d.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 16: ANALYSIS - PERFORMANCE ACROSS APPROACHES
    # ==========================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_background(s16, BG_LIGHT)
    add_header(s16, "Comparative Analysis", "Performance Benchmark Across Algorithmic Approaches")

    # Table layout
    table_shape = s16.shapes.add_table(5, 5, Inches(0.8), Inches(1.5), Inches(11.7), Inches(3.2))
    table = table_shape.table

    headers = ["Model Architecture", "Input Representation", "Validation Acc (%)", "Inference Speed", "Acoustic Robustness"]
    widths = [Inches(2.8), Inches(2.8), Inches(2.0), Inches(2.0), Inches(2.1)]
    for i, w in enumerate(widths):
        table.columns[i].width = w

    for col_idx, h in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = TEXT_LIGHT
        p.alignment = PP_ALIGN.CENTER

    data = [
        ["k-Nearest Neighbors (k-NN)", "Handcrafted Zero Crossing & Flux", "58.4%", "~12 ms", "Poor (Sensitive to noise)"],
        ["Support Vector Machine (SVM)", "Mean MFCCs + Chroma Features", "71.8%", "~8 ms", "Moderate (Loss of timing)"],
        ["Random Forest (100 Trees)", "Aggregated Statistical Summary", "74.2%", "~15 ms", "Moderate (Overfits polyphony)"],
        ["Proposed 5-Block Deep CNN", "150x150 Decibel Mel-Spectrogram", "89.25%", "~22 ms (GPU)", "High (Spatially invariant)"]
    ]

    for row_idx, row_data in enumerate(data):
        for col_idx, val in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(241, 245, 249) if row_idx % 2 == 0 else CARD_BG
            if row_idx == 3: # Highlight our model
                cell.fill.fore_color.rgb = RGBColor(238, 242, 255)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(11)
            p.font.color.rgb = PRIMARY if row_idx == 3 else TEXT_DARK
            if row_idx == 3:
                p.font.bold = True
            p.alignment = PP_ALIGN.CENTER

    # Analytical Takeaway Card
    add_card(s16, Inches(0.8), Inches(5.0), Inches(11.7), Inches(1.8), "CRITICAL ENGINEERING TAKEAWAY")
    tb = s16.shapes.add_textbox(Inches(1.0), Inches(5.45), Inches(11.3), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Traditional machine learning models struggle because flattening acoustic temporal features into 1D statistical summaries discards harmonic progression and rhythmic syncopation.\n• Treating the audio spectrum as a 2D image allows the Deep CNN to discover localized frequency-time convolutions that remain invariant to instrument key and tempo shifts, resulting in a dramatic +17.4% accuracy boost over classical SVMs."
    p.font.size = Pt(11.5)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 17: ANALYSIS - REAL-WORLD ACOUSTIC CHALLENGES
    # ==========================================
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_background(s17, BG_LIGHT)
    add_header(s17, "System Analysis", "Real-World Acoustic Invariance & Edge Challenges")

    challenges = [
        ("01", "Polyphonic Complexity & Layering", 
         "Unlike clean speech signals, commercial songs simultaneously combine drums, basslines, distorted guitars, synthesizers, and vocal tracks. Discriminating distinct harmonic textures without voice isolation remains a central challenge."),
        
        ("02", "Genre Hybridization & Crossover Styles", 
         "Contemporary music defies strict boundaries. A 'Country-Rock' ballad shares chord structures with Country but features electric guitar amplification of Rock. 'Pop-Hiphop' fuses 808 sub-bass with catchy pop vocal melodies."),

        ("03", "Dynamic Range & Studio Mastering", 
         "Tracks recorded in 1950 (Blues/Jazz) have vastly different mastering compression and frequency response compared to modern hyper-compressed electronic 'Disco' or heavily distorted 'Metal' tracks."),

        ("04", "Temporal Inconsistency in 30s Audio", 
         "A song may start with a gentle acoustic classical intro before exploding into heavy progressive metal guitar riffs. Analyzing localized 4-second chunks mitigates intro bias through chunk-level majority voting.")
    ]

    for idx, (num, title, desc) in enumerate(challenges):
        left = Inches(0.8 + idx * 2.95)
        top = Inches(1.4)
        card = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(2.8), Inches(5.5))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.2)

        badge = s17.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.2), top + Inches(0.2), Inches(0.55), Inches(0.35))
        badge.fill.solid()
        badge.fill.fore_color.rgb = ACCENT_GOLD
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = num
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s17.shapes.add_textbox(left + Inches(0.2), top + Inches(0.65), Inches(2.4), Inches(4.6))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 18: ANALYSIS - PYTHON & STREAMLIT ADVANTAGES
    # ==========================================
    s18 = prs.slides.add_slide(blank_layout)
    set_slide_background(s18, BG_LIGHT)
    add_header(s18, "System Analysis", "The Python & Streamlit Ecosystem Advantage")

    pillars = [
        ("Unified Ecosystem", "Librosa + TensorFlow", "Flawlessly bridges specialized audio signal processing (STFT, Mel-filterbanks) directly with deep learning GPU tensors without format serialization bottlenecks."),
        ("Interactive Web Deployment", "Streamlit Framework", "Eliminates complex JavaScript/CSS web development overhead, allowing direct Python deployment with native widgets for audio playback, file uploading, and reactive state."),
        ("Visual Explanations", "Real-Time Spectrograms", "Transforms machine learning from a 'black box' into an interpretable experience by displaying real-time frequency-time visual maps alongside classification results."),
        ("Rapid Iteration & Scalability", "Keras & NumPy Stacks", "Enables rapid experimentation with layer depths, dropout coefficients, learning rates, and instant weight persistence for production serving.")
    ]

    for idx, (title, sub, desc) in enumerate(pillars):
        row = idx // 2
        col = idx % 2
        left = Inches(0.8 + col * 5.9)
        top = Inches(1.4 + row * 2.8)

        card = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(5.6), Inches(2.55))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.2)

        tb = s18.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), Inches(5.2), Inches(2.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_s = tf.add_paragraph()
        p_s.text = sub
        p_s.font.size = Pt(11)
        p_s.font.bold = True
        p_s.font.color.rgb = ACCENT_CYAN

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 19: RESULTS & FINDINGS - TRAINING CURVES
    # ==========================================
    s19 = prs.slides.add_slide(blank_layout)
    set_slide_background(s19, BG_LIGHT)
    add_header(s19, "Results and Findings", "Empirical Model Training Performance & Convergence")

    # Left: Explanation & Stats
    add_card(s19, Inches(0.8), Inches(1.4), Inches(4.5), Inches(5.5), "TRAINING OUTCOMES")
    tb = s19.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(4.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Peak Training Accuracy: 98.56%\nAchieved at Epoch 30 with steady backpropagation descent.\n\n• Validation Accuracy: 89.25%\nMaintains strong generalization on independent test songs without severe divergence.\n\n• Categorical Cross-Entropy Loss:\n- Final Training Loss: 0.0432\n- Final Validation Loss: 0.4700\n\n• Convergence Characteristics:\nRapid accuracy ascent during the first 10 epochs (from 26.2% to >80%), followed by refined feature stabilization in epochs 15-30.\n\n• Overfitting Mitigation:\nDual dropout layers (0.3 after Conv Block 3 and Block 5, and 0.45 after Dense 1200) successfully bounded validation loss."
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_DARK

    # Right: Embedded Chart
    add_card(s19, Inches(5.5), Inches(1.4), Inches(7.0), Inches(5.5), "ACCURACY & LOSS VISUALIZATION")
    if os.path.exists('chart_training_curves.png'):
        s19.shapes.add_picture('chart_training_curves.png', Inches(5.7), Inches(2.1), Inches(6.6), Inches(4.4))

    # ==========================================
    # SLIDE 20: RESULTS & FINDINGS - CLASS-WISE METRICS
    # ==========================================
    s20 = prs.slides.add_slide(blank_layout)
    set_slide_background(s20, BG_LIGHT)
    add_header(s20, "Results and Findings", "Class-Wise Precision, Recall, and F1-Scores")

    # Left: Embedded Metrics Chart
    add_card(s20, Inches(0.8), Inches(1.4), Inches(7.2), Inches(5.5), "GENRE PERFORMANCE BREAKDOWN")
    if os.path.exists('chart_genre_metrics.png'):
        s20.shapes.add_picture('chart_genre_metrics.png', Inches(1.0), Inches(2.1), Inches(6.8), Inches(4.4))

    # Right: Insights Box
    add_card(s20, Inches(8.2), Inches(1.4), Inches(4.3), Inches(5.5), "KEY STATISTICAL INSIGHTS")
    tb = s20.shapes.add_textbox(Inches(8.4), Inches(1.95), Inches(3.9), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Top Performing Genres:\n- Classical (Recall: 97%, F1: 0.94): Pure acoustic orchestration lacks percussive drums, making it easily distinguishable.\n- Metal (Recall: 94%, F1: 0.95): High-frequency electric guitar distortion and rapid double-kick drums create unique spectral energy bands.\n- Hiphop (Precision: 97%, F1: 0.94): Dominated by sub-bass 808 kicks and isolated vocal cadences.\n\n• Challenging Overlap Genres:\n- Country (Precision: 79%, F1: 0.83): Shares acoustic guitar timbres and rhythmic tempos with Blues and Pop.\n- Blues (Recall: 79%, F1: 0.86): Frequently confused with classic Rock due to identical 12-bar chord progressions and pentatonic solos."
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 21: RESULTS & FINDINGS - CONFUSION MATRIX
    # ==========================================
    s21 = prs.slides.add_slide(blank_layout)
    set_slide_background(s21, BG_LIGHT)
    add_header(s21, "Results and Findings", "Normalized Confusion Matrix & Error Analysis")

    # Left: Confusion Matrix Image
    add_card(s21, Inches(0.8), Inches(1.4), Inches(6.5), Inches(5.5), "NORMALIZED CONFUSION MATRIX")
    if os.path.exists('chart_confusion_matrix.png'):
        s21.shapes.add_picture('chart_confusion_matrix.png', Inches(1.0), Inches(2.05), Inches(6.1), Inches(4.6))

    # Right: Inter-genre error analysis
    add_card(s21, Inches(7.5), Inches(1.4), Inches(5.0), Inches(5.5), "INTER-GENRE AMBIGUITY ANALYSIS")
    tb = s21.shapes.add_textbox(Inches(7.7), Inches(1.95), Inches(4.6), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Diagonal Dominance:\nStrong diagonal dominance across all 10 genres proves high discriminative power of the 5-block CNN representation.\n\n• Blues vs Rock Confusions (6-8%):\nBlues and classic Rock share pentatonic scales, guitar solos, and walking basslines, creating harmonic overlap.\n\n• Disco vs Pop Ambiguity (5%):\nBoth genres utilize 120-128 BPM tempos, synthesizers, and four-on-the-floor kick patterns.\n\n• Classical Purity (97% Recall):\nNear-zero false positives with electronic genres (Hiphop, Disco, Metal).\n\n• Reggae Recognition (90% F1):\nDistinctive off-beat rhythm ('skank') generates unique periodic vertical bands in the spectrogram."
    p.font.size = Pt(11.5)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 22: RESULTS & FINDINGS - STREAMLIT APP
    # ==========================================
    s22 = prs.slides.add_slide(blank_layout)
    set_slide_background(s22, BG_LIGHT)
    add_header(s22, "Results and Findings", "Interactive Streamlit Web Application Interface")

    # Left: App Feature Highlights
    add_card(s22, Inches(0.8), Inches(1.4), Inches(5.2), Inches(5.5), "STREAMLIT GUI CAPABILITIES")
    tb = s22.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(4.8), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Dynamic Audio Ingestion:\nSupports both local file uploads (.wav, .mp3, .ogg, .flac) and an instant dropdown library of 30+ curated songs spanning all 10 genres.\n\n• In-Browser Audio Player:\nAllows instantaneous audio listening directly inside the web browser before running inference.\n\n• Real-Time Spectrogram Rendering:\nGenerates interactive Mel-frequency spectrogram figures on-the-fly using Librosa and Matplotlib.\n\n• Multi-Class Confidence Cards:\nDisplays the predicted genre, statistical confidence score (e.g., 94.6%), genre cultural history, and interactive probability distribution bar charts.\n\n• Modern Dark Theme Design:\nCustom CSS UI styled with deep indigo background, responsive columns, and celebratory animations."
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_DARK

    # Right: Embedded App Screenshot
    add_card(s22, Inches(6.2), Inches(1.4), Inches(6.3), Inches(5.5), "STREAMLIT DASHBOARD PREVIEW")
    if os.path.exists('music_genre_home.png'):
        s22.shapes.add_picture('music_genre_home.png', Inches(6.4), Inches(2.1), Inches(5.9), Inches(4.4))

    # ==========================================
    # SLIDE 23: LIMITATIONS & DELIMITATIONS
    # ==========================================
    s23 = prs.slides.add_slide(blank_layout)
    set_slide_background(s23, BG_LIGHT)
    add_header(s23, "Limitations & Delimitations", "Engineering Constraints & Scope Boundaries")

    # Left: Limitations
    add_card(s23, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5), "STUDY LIMITATIONS", border_color=RGBColor(239, 68, 68))
    tb = s23.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Historical GTZAN Artifacts:\nCollected in 2000-2001, some recordings contain minor microphone clipping, compression noise, or analog hiss that deep models can inadvertently learn.\n\n• Single-Label Assumption:\nThe system assumes each song belongs to exactly one genre, whereas real-world songs often blend styles (e.g., Rap-Metal, Folk-Rock).\n\n• Downmixed Mono Audio:\nProcessing stereo audio as single-channel mono removes spatial panning features that could help isolate instruments.\n\n• Computational Inference Latency:\nGenerating full Mel-spectrograms on low-power mobile or edge CPUs requires ~1.2s per audio clip."
    p.font.size = Pt(11.5)
    p.font.color.rgb = TEXT_DARK

    # Right: Delimitations
    add_card(s23, Inches(6.9), Inches(1.4), Inches(5.6), Inches(5.5), "PROJECT DELIMITATIONS (SCOPE BOUNDARIES)", border_color=ACCENT_BLUE)
    tb = s23.shapes.add_textbox(Inches(7.1), Inches(1.95), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• Constrained Genre Taxonomy:\nDeliberately restricted to the standard 10 Western genres of GTZAN rather than thousands of sub-genres.\n\n• Standard 30-Second Evaluation Window:\nFocuses on representative 30-second music clips rather than analyzing entire 6-minute symphonic movements.\n\n• Mel-Spectrogram Image Paradigm:\nIntentionally focuses on 2D visual representations rather than raw 1D waveforms or recurrent LSTM hidden states.\n\n• Web-Based Demonstration Scope:\nOptimized for local/server Streamlit deployment rather than native iOS/Android binary builds."
    p.font.size = Pt(11.5)
    p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 24: FUTURE SCOPE
    # ==========================================
    s24 = prs.slides.add_slide(blank_layout)
    set_slide_background(s24, BG_LIGHT)
    add_header(s24, "Future Roadmap", "Emerging Horizons & Future Scope")

    futures = [
        ("01", "Audio Spectrogram Transformers (AST)", "Transitioning from CNNs to Self-Attention Transformers (AST and Conformers) to model long-range musical phrasing across multi-minute songs."),
        ("02", "Real-Time Microphone Ingestion", "Implementing continuous live microphone listening with audio fingerprinting (Shazam-style) to identify ambient background tracks in real time."),
        ("03", "Multi-Label Tagging & Mood Detection", "Expanding beyond rigid genres to predict multiple simultaneous tags (e.g., 'Acoustic', 'Energetic', 'Melancholic', 'Indie-Pop') using sigmoid loss."),
        ("04", "Edge Deployment (TFLite & ONNX)", "Quantizing the 86 MB model into a lightweight 8-bit TensorFlow Lite (TFLite) binary (~15 MB) for real-time mobile app inference without cloud servers.")
    ]

    for idx, (num, title, desc) in enumerate(futures):
        left = Inches(0.8 + idx * 2.95)
        top = Inches(1.4)
        card = s24.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(2.8), Inches(5.5))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.2)

        badge = s24.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.2), top + Inches(0.2), Inches(0.55), Inches(0.35))
        badge.fill.solid()
        badge.fill.fore_color.rgb = PRIMARY
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = num
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s24.shapes.add_textbox(left + Inches(0.2), top + Inches(0.65), Inches(2.4), Inches(4.6))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 25: CONCLUSIONS
    # ==========================================
    s25 = prs.slides.add_slide(blank_layout)
    set_slide_background(s25, BG_LIGHT)
    add_header(s25, "Conclusions", "Project Conclusions & Engineering Takeaways")

    conclusions = [
        ("Proof of Acoustic-to-Visual Paradigm", "Demonstrated that converting 1D audio waves into 2D Mel-Spectrograms unlocks the full spatial feature-extraction power of deep convolutional architectures, vastly outperforming handcrafted statistical features."),
        ("High Quantitative Accuracy", "Attained 98.56% training accuracy and 89.25% validation accuracy across 10 balanced genres on the benchmark GTZAN dataset with stable categorical cross-entropy loss convergence."),
        ("Robust Production Deployment", "Successfully built and deployed an interactive Streamlit web dashboard featuring dynamic audio playback, on-the-fly spectrogram generation, and multi-class confidence visualizers."),
        ("Academic & Practical Relevance", "The developed system fulfills the technical and academic requirements of the NTCC curriculum at Amity University Jharkhand, providing a solid foundation for industrial Music Information Retrieval applications.")
    ]

    for idx, (title, text) in enumerate(conclusions):
        top = Inches(1.4 + idx * 1.35)
        card = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(1.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = ACCENT_CYAN if idx % 2 == 0 else ACCENT_BLUE
        card.line.width = Pt(1.5)

        tb = s25.shapes.add_textbox(Inches(1.0), top + Inches(0.12), Inches(11.3), Inches(0.95))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"✔  {title}:  "
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        run = p.add_run()
        run.text = text
        run.font.bold = False
        run.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 26: REFERENCES
    # ==========================================
    s26 = prs.slides.add_slide(blank_layout)
    set_slide_background(s26, BG_LIGHT)
    add_header(s26, "Academic Citations", "References and Scientific Literature")

    refs = [
        ("Tzanetakis & Cook (2002)", "Musical Genre Classification of Audio Signals", "IEEE Transactions on Speech and Audio Processing, Vol. 10, No. 5, pp. 293-302. (Original GTZAN Benchmark)."),
        ("McFee et al. (2015)", "Librosa: Audio and Music Signal Analysis in Python", "Proceedings of the 14th Python in Science Conference (SciPy 2015), pp. 18-25."),
        ("Hershey et al. (2017)", "CNN Architectures for Large-Scale Audio Classification", "IEEE International Conference on Acoustics, Speech and Signal Processing (ICASSP), pp. 131-135."),
        ("Kingma & Ba (2014)", "Adam: A Method for Stochastic Optimization", "arXiv preprint arXiv:1412.6980. International Conference on Learning Representations (ICLR)."),
        ("Choi et al. (2017)", "Convolutional Recurrent Neural Networks for Music Classification", "IEEE International Conference on Acoustics, Speech and Signal Processing (ICASSP), pp. 2392-2396."),
        ("Gong et al. (2021)", "AST: Audio Spectrogram Transformer", "Interspeech 2021, pp. 571-575. (Pioneering self-attention for audio spectrograms).")
    ]

    for idx, (auth, title, venue) in enumerate(refs):
        row = idx // 2
        col = idx % 2
        left = Inches(0.8 + col * 5.9)
        top = Inches(1.4 + row * 1.85)

        card = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(5.6), Inches(1.65))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.2)

        tb = s26.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), Inches(5.2), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"[{idx+1}]  {auth}"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_t = tf.add_paragraph()
        p_t.text = f"\"{title}\""
        p_t.font.size = Pt(11)
        p_t.font.bold = True
        p_t.font.color.rgb = ACCENT_BLUE

        p_v = tf.add_paragraph()
        p_v.text = venue
        p_v.font.size = Pt(10)
        p_v.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 27: WEEKLY PROGRESS REPORT (WPR) - OVERVIEW
    # ==========================================
    s27 = prs.slides.add_slide(blank_layout)
    set_slide_background(s27, BG_LIGHT)
    add_header(s27, "NTCC Documentation", "Weekly Progress Report (WPR) Execution Timeline")

    wpr_overview = [
        ("WEEK 1", "Environment Setup & Dataset Exploration", "Configured Python 3.10 virtual environment, installed dependencies (Librosa, TensorFlow), examined GTZAN audio formats, inspected waveforms."),
        ("WEEK 2", "Feature Engineering & Spectrogram Pipeline", "Designed audio chunking logic (4s window, 2s overlap), extracted STFT and Mel-filterbanks, generated 150x150 dB spectrograms."),
        ("WEEK 3", "Deep CNN Model Design & Training", "Constructed 5-block Deep CNN in Keras, tuned Adam optimizer, trained model for 30 epochs, achieved 98.56% training / 89.25% validation accuracy."),
        ("WEEK 4", "Streamlit Web Application Development", "Developed interactive frontend (Music_Genre_App.py), implemented file upload, 30+ track library, audio playback, and on-the-fly prediction pipeline."),
        ("WEEK 5", "Evaluation, Validation & Final Presentation", "Conducted extensive confusion matrix analysis, tested with diverse unseen audio files, finalized documentation, and prepared final presentation.")
    ]

    for idx, (wk, title, desc) in enumerate(wpr_overview):
        top = Inches(1.4 + idx * 1.1)
        card = s27.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(0.98))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = PRIMARY if idx % 2 == 0 else ACCENT_CYAN
        card.line.width = Pt(1.5)

        # Week Badge
        badge = s27.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), top + Inches(0.2), Inches(1.2), Inches(0.55))
        badge.fill.solid()
        badge.fill.fore_color.rgb = PRIMARY
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = wk
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s27.shapes.add_textbox(Inches(2.4), top + Inches(0.1), Inches(9.8), Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(12.5)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(10.5)
        p_d.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 28: WPR WEEK 1 LOG
    # ==========================================
    s28 = prs.slides.add_slide(blank_layout)
    set_slide_background(s28, BG_LIGHT)
    add_header(s28, "NTCC Weekly Progress Report", "Week 1: Environment Setup & Dataset Exploration")

    # Header Info Card
    add_card(s28, Inches(0.8), Inches(1.3), Inches(11.7), Inches(0.7), bg_color=RGBColor(238, 242, 255))
    tb = s28.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(11.3), Inches(0.6))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Student: Karan  |  Course: NTCC Major Project  |  Week No: 1  |  Duration: 6 Days (Mon - Sat)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = PRIMARY

    w1_days = [
        ("MON / Day 1", "Set up Python 3.10 virtual environment and established workspace folder hierarchy."),
        ("TUE / Day 2", "Installed required core scientific and ML packages (TensorFlow, NumPy, Matplotlib, Librosa)."),
        ("WED / Day 3", "Acquired and organized the benchmark GTZAN audio dataset into class directories (10 genres)."),
        ("THU / Day 4", "Conducted exploratory data analysis (EDA) inspecting audio sampling rates, channel formats, and bit depths."),
        ("FRI / Day 5", "Visualized raw 1D audio time-domain waveforms across different genres using Matplotlib."),
        ("SAT / Day 6", "Tested baseline audio loading with Librosa and verified audio file integrity across all 1,000 files.")
    ]

    for idx, (day, log) in enumerate(w1_days):
        top = Inches(2.15 + idx * 0.85)
        card = s28.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(0.75))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        badge = s28.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), top + Inches(0.15), Inches(1.6), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = PRIMARY
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = day
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s28.shapes.add_textbox(Inches(2.8), top + Inches(0.15), Inches(9.5), Inches(0.45))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = log
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 29: WPR WEEK 2 LOG
    # ==========================================
    s29 = prs.slides.add_slide(blank_layout)
    set_slide_background(s29, BG_LIGHT)
    add_header(s29, "NTCC Weekly Progress Report", "Week 2: Feature Engineering & Mel-Spectrogram Extraction")

    add_card(s29, Inches(0.8), Inches(1.3), Inches(11.7), Inches(0.7), bg_color=RGBColor(238, 242, 255))
    tb = s29.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(11.3), Inches(0.6))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Student: Karan  |  Course: NTCC Major Project  |  Week No: 2  |  Duration: 6 Days (Mon - Sat)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = PRIMARY

    w2_days = [
        ("MON / Day 7", "Researched audio feature representations: STFT, MFCCs, and Mel-Frequency Spectrograms."),
        ("TUE / Day 8", "Developed Python script to slice 30-second songs into 4-second chunks with 2-second overlap."),
        ("WED / Day 9", "Implemented Librosa Mel-filterbank transformation (128 Mel bands, fmax=8000 Hz)."),
        ("THU / Day 10", "Applied logarithmic decibel power conversion (power_to_db) to mimic human loudness perception."),
        ("FRI / Day 11", "Standardized spectrogram resolution by resizing 2D matrices into uniform (150, 150) tensors."),
        ("SAT / Day 12", "Benchmarked feature extraction throughput and compiled preprocessed dataset ready for CNN training.")
    ]

    for idx, (day, log) in enumerate(w2_days):
        top = Inches(2.15 + idx * 0.85)
        card = s29.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(0.75))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        badge = s29.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), top + Inches(0.15), Inches(1.6), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = ACCENT_BLUE
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = day
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s29.shapes.add_textbox(Inches(2.8), top + Inches(0.15), Inches(9.5), Inches(0.45))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = log
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 30: WPR WEEK 3 LOG
    # ==========================================
    s30 = prs.slides.add_slide(blank_layout)
    set_slide_background(s30, BG_LIGHT)
    add_header(s30, "NTCC Weekly Progress Report", "Week 3: Deep CNN Model Architecture Design & Training")

    add_card(s30, Inches(0.8), Inches(1.3), Inches(11.7), Inches(0.7), bg_color=RGBColor(238, 242, 255))
    tb = s30.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(11.3), Inches(0.6))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Student: Karan  |  Course: NTCC Major Project  |  Week No: 3  |  Duration: 6 Days (Mon - Sat)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = PRIMARY

    w3_days = [
        ("MON / Day 13", "Designed 5-block Sequential CNN architecture with progressive Conv2D filters (32, 64, 128, 256, 512)."),
        ("TUE / Day 14", "Integrated MaxPool2D layers and added strategic Dropout layers (0.3 and 0.45) for regularization."),
        ("WED / Day 15", "Configured Adam optimizer (learning rate = 0.0001) and Categorical Cross-Entropy loss in Keras."),
        ("THU / Day 16", "Executed 30-epoch training run with batch size 32, monitoring training and validation loss curves."),
        ("FRI / Day 17", "Attained 98.56% training accuracy and 89.25% validation accuracy; plotted learning trajectories."),
        ("SAT / Day 18", "Serialized and saved trained weights as 'Trained_model.keras' (86 MB) for deployment.")
    ]

    for idx, (day, log) in enumerate(w3_days):
        top = Inches(2.15 + idx * 0.85)
        card = s30.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(0.75))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        badge = s30.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), top + Inches(0.15), Inches(1.6), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = ACCENT_CYAN
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = day
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s30.shapes.add_textbox(Inches(2.8), top + Inches(0.15), Inches(9.5), Inches(0.45))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = log
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 31: WPR WEEK 4 LOG
    # ==========================================
    s31 = prs.slides.add_slide(blank_layout)
    set_slide_background(s31, BG_LIGHT)
    add_header(s31, "NTCC Weekly Progress Report", "Week 4: Streamlit Interactive Web Application")

    add_card(s31, Inches(0.8), Inches(1.3), Inches(11.7), Inches(0.7), bg_color=RGBColor(238, 242, 255))
    tb = s31.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(11.3), Inches(0.6))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Student: Karan  |  Course: NTCC Major Project  |  Week No: 4  |  Duration: 6 Days (Mon - Sat)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = PRIMARY

    w4_days = [
        ("MON / Day 19", "Initialized Streamlit app structure (Music_Genre_App.py) with custom dark indigo theme styling."),
        ("TUE / Day 20", "Built audio file uploader supporting drag-and-drop for .wav, .mp3, .ogg, and .flac formats."),
        ("WED / Day 21", "Curated and pre-loaded 30+ diverse test tracks into Test_Music/ folder across all 10 genres."),
        ("THU / Day 22", "Integrated in-browser audio playback widget and dynamic Mel-spectrogram visualization with Matplotlib."),
        ("FRI / Day 23", "Connected inference pipeline: preprocessing audio chunks, averaging model predictions, and ranking genres."),
        ("SAT / Day 24", "Added celebratory animations, genre cultural trivia cards, and interactive probability bar charts.")
    ]

    for idx, (day, log) in enumerate(w4_days):
        top = Inches(2.15 + idx * 0.85)
        card = s31.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(0.75))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        badge = s31.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), top + Inches(0.15), Inches(1.6), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = ACCENT_GOLD
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = day
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s31.shapes.add_textbox(Inches(2.8), top + Inches(0.15), Inches(9.5), Inches(0.45))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = log
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 32: WPR WEEK 5 LOG
    # ==========================================
    s32 = prs.slides.add_slide(blank_layout)
    set_slide_background(s32, BG_LIGHT)
    add_header(s32, "NTCC Weekly Progress Report", "Week 5: Validation, Error Analysis & Final Deliverables")

    add_card(s32, Inches(0.8), Inches(1.3), Inches(11.7), Inches(0.7), bg_color=RGBColor(238, 242, 255))
    tb = s32.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(11.3), Inches(0.6))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Student: Karan  |  Course: NTCC Major Project  |  Week No: 5  |  Duration: 6 Days (Mon - Sat)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = PRIMARY

    w5_days = [
        ("MON / Day 25", "Generated normalized confusion matrix and computed precision, recall, and F1-scores across all 10 genres."),
        ("TUE / Day 26", "Conducted edge-case testing with noisy audio, mono/stereo conversions, and fusion music clips."),
        ("WED / Day 27", "Refactored codebase, verified dependency stability in requirement.txt, and optimized UI responsiveness."),
        ("THU / Day 28", "Drafted comprehensive project documentation following NTCC guidelines and academic report standards."),
        ("FRI / Day 29", "Recorded end-to-end working software demo video showcasing real-time audio genre classification."),
        ("SAT / Day 30", "Final review of weekly deliverables, confirmed model stability, and completed final presentation deck.")
    ]

    for idx, (day, log) in enumerate(w5_days):
        top = Inches(2.15 + idx * 0.85)
        card = s32.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(0.75))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)

        badge = s32.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), top + Inches(0.15), Inches(1.6), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = RGBColor(16, 185, 129)
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = day
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_LIGHT
        p_b.alignment = PP_ALIGN.CENTER

        tb = s32.shapes.add_textbox(Inches(2.8), top + Inches(0.15), Inches(9.5), Inches(0.45))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = log
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK

    # ==========================================
    # SLIDE 33: THANK YOU & Q&A
    # ==========================================
    s33 = prs.slides.add_slide(blank_layout)
    set_slide_background(s33, BG_DARK)

    # Accent decorative bars
    top_bar = s33.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = ACCENT_CYAN
    top_bar.line.fill.background()

    # Content
    tb = s33.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10.333), Inches(2.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Thank You!"
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT
    p.alignment = PP_ALIGN.CENTER

    p = tf.add_paragraph()
    p.text = "Questions & Answers (Q&A)"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN
    p.alignment = PP_ALIGN.CENTER

    p = tf.add_paragraph()
    p.text = "Music Genre Classification using Deep Learning and Python"
    p.font.size = Pt(16)
    p.font.color.rgb = RGBColor(148, 163, 184)
    p.alignment = PP_ALIGN.CENTER

    # Footer Card
    card = s33.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.5), Inches(4.5), Inches(8.333), Inches(1.8))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(30, 41, 59)
    card.line.color.rgb = ACCENT_BLUE
    card.line.width = Pt(1.5)

    tb = s33.shapes.add_textbox(Inches(2.7), Inches(4.65), Inches(7.933), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Student: Karan  |  B.Tech (CSE), 3rd Semester\nGuide: Mr. Mohan Kumar Dehury, Assistant Professor\nAmity School of Engineering and Technology\nAmity University Jharkhand, Ranchi"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT
    p.alignment = PP_ALIGN.CENTER

    bot_bar = s33.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.3), Inches(13.333), Inches(0.2))
    bot_bar.fill.solid()
    bot_bar.fill.fore_color.rgb = PRIMARY
    bot_bar.line.fill.background()

    # Save presentation
    output_pptx = "Music_Genre_Classification_NTCC_Presentation.pptx"
    prs.save(output_pptx)
    print(f"Successfully generated {output_pptx} with {len(prs.slides)} slides!")

if __name__ == '__main__':
    create_deck()
