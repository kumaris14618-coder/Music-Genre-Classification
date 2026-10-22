import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_docx():
    doc = docx.Document()

    # Page Margins: 1 inch all around
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Styles setup
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Times New Roman'
    font_normal.size = Pt(12)
    font_normal.color.rgb = RGBColor(0, 0, 0)

    def add_p(text, bold=False, italic=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6, space_before=0, color=(0,0,0)):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(*color)
        return p

    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0, 51, 102)
        return p

    def add_heading_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(30, 41, 59)
        return p

    # ==========================================
    # 1. TITLE / COVER PAGE
    # ==========================================
    add_p("Submitted to", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=20)
    add_p("Amity University Jharkhand", bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER, color=(0, 51, 102))
    
    add_p("\n\nNTCC Major Project Report\non", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("Music Genre Classification using Deep Learning and Python", bold=True, size=20, align=WD_ALIGN_PARAGRAPH.CENTER, color=(0, 51, 102))
    
    add_p("\n\nBy", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("Karan", bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("Enrolment Number: [Your Enrolment No.]", bold=False, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("Program: Bachelor of Technology (B.Tech - CSE)\nBatch: 2024-2028", size=12, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_p("\n\nUnder the guidance of", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("Mr. Mohan Kumar Dehury", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("Assistant Professor, Department of Computer Science & Engineering", size=11, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_p("\n\nAMITY SCHOOL OF ENGINEERING AND TECHNOLOGY", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=40)
    add_p("AMITY UNIVERSITY JHARKHAND\nRANCHI 2024-2028", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_page_break()

    # ==========================================
    # 2. DECLARATION
    # ==========================================
    add_p("DECLARATION", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
    add_p("I, Karan, student of Bachelor of Technology in Computer Science & Engineering hereby declare that the NTCC Report titled “Music Genre Classification using Deep Learning and Python” which is submitted by me to Department of Amity School of Engineering and Technology, Amity University Jharkhand, in partial fulfilment of requirement for the award of degree of Bachelor of Technology, has not been previously formed the basis for the award of any degree, diploma or other similar title or recognition.")
    
    add_p("\n\nAmity University Jharkhand, Ranchi\nDate: 22nd October 2026\n\n\nSign. of the Student: _______________________")
    add_p("Name of Student: Karan\nEnrolment Number: [Your Enrolment No.]\nProgram: B.Tech (CSE), 3rd Semester")

    doc.add_page_break()

    # ==========================================
    # 3. CERTIFICATE
    # ==========================================
    add_p("CERTIFICATE", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
    add_p("On the basis of Major Project submitted by Karan, student of Bachelor of Technology, I hereby certify that the NTCC Report “Music Genre Classification using Deep Learning and Python” which is submitted to Department of Computer Science, Amity School of Engineering and Technology, Amity University Jharkhand in partial fulfilment of requirement for the award of the degree of Bachelor of Technology is an original contribution with existing knowledge and faithful record of work carried out by him/her under my guidance and supervision.")
    add_p("To the best of my knowledge this work has not been submitted in part or full for any Degree or Diploma to this University or elsewhere.")

    add_p("\n\nAmity University Jharkhand, Ranchi\nDate: 22nd October 2026\n\n\n")
    add_p("Signature of Guide (Internal): _______________________        Signature of Guide (External): _______________________")
    add_p("Name: Mr. Mohan Kumar Dehury\nDesignation: Assistant Professor\nAmity School of Engineering and Technology, Amity University Jharkhand")

    doc.add_page_break()

    # ==========================================
    # 4. ACKNOWLEDGEMENT
    # ==========================================
    add_p("ACKNOWLEDGEMENT", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
    add_p("While conducting this report, I got support in many ways from many people. Firstly, I am deeply grateful to my project guide, Mr. Mohan Kumar Dehury, who helped me with full devotion and always supported me earnestly whenever it was needed. Without his guidance, mental & moral support and academic inputs this report was not possible.")
    add_p("This Training and Major Project report could never have seen the light of day without the cooperation of faculty members and mentors of Amity School of Engineering and Technology who evaluated and assisted my research direction.")
    add_p("My friends and peers have been the biggest support for me at every juncture of life. They manifested their great interest in my research work also and always tried to make things easy for me.")
    add_p("A word of gratitude goes to my family members whose love, affection and understanding have enabled me to complete this endeavor with ease. At the end, I thank the Almighty for giving me courage and strength to conduct this project report.")

    doc.add_page_break()

    # ==========================================
    # 5. TABLE OF CONTENTS
    # ==========================================
    add_p("CONTENT", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

    toc_items = [
        ("1. Abstract", "1"),
        ("2. Introduction", "2-3"),
        ("3. Literature Review", "4-6"),
        ("4. Methodology", "7-11"),
        ("5. Analysis", "12-13"),
        ("6. Result Findings / Outcomes", "14-15"),
        ("7. Limitations & Delimitations", "16-17"),
        ("8. Future Scope", "18-19"),
        ("9. Conclusions", "20-21"),
        ("10. Reference", "22"),
        ("11. NTCC Weekly Progress Reports (WPR Weeks 1 - 5)", "23-28")
    ]

    t_toc = doc.add_table(rows=len(toc_items), cols=2)
    t_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_toc.columns[0].width = Inches(5.5)
    t_toc.columns[1].width = Inches(1.0)
    for idx, (title, page) in enumerate(toc_items):
        cell_0 = t_toc.cell(idx, 0)
        cell_1 = t_toc.cell(idx, 1)
        p0 = cell_0.paragraphs[0]
        p0.add_run(title).bold = True
        p1 = cell_1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.add_run(page)

    doc.add_page_break()

    # ==========================================
    # 1. ABSTRACT
    # ==========================================
    add_heading_1("1. ABSTRACT")
    add_p("Music genre classification represents an essential pillar of Music Information Retrieval (MIR), serving as the foundational engine for automated cataloging, music recommendation systems, and audio streaming playlist curation. In the modern era of music streaming, millions of tracks are ingested continuously, rendering manual categorization impossible. Python has established itself as the leading language for building audio machine learning applications due to its extensive ecosystem comprising specialized audio processing libraries such as Librosa and modern deep learning frameworks such as TensorFlow and Keras.")
    add_p("At its core, music genre classification involves transforming complex, continuous time-domain acoustic signals into discriminative mathematical representations that capture harmonic, rhythmic, and timbral properties. While traditional approaches depended heavily on handcrafted heuristics (such as Zero Crossing Rate, Spectral Centroid, and Mel-Frequency Cepstral Coefficients), modern deep learning formulates audio classification as an advanced visual recognition task by converting sound waves into two-dimensional Mel-Frequency Spectrograms.")
    add_p("This project implements a complete, end-to-end deep learning framework utilizing the benchmark GTZAN dataset, known as the 'MNIST of Sound', comprising 1,000 thirty-second audio tracks balanced across ten distinct musical genres: blues, classical, country, disco, hiphop, jazz, metal, pop, reggae, and rock. Audio signals are split into 4-second overlapping chunks and mapped into 150x150 decibel-scaled Mel-spectrogram images. A custom 5-block Deep Convolutional Neural Network (CNN) is architected with progressive filter channels (32 to 512), spatial max-pooling, and dropout regularization (0.3 and 0.45). The model achieves a peak training accuracy of 98.56% and a validation accuracy of 89.25%. Furthermore, the model is fully integrated into a responsive Streamlit web application providing instant audio playback, dynamic spectrogram rendering, and multi-class genre probability distribution.")

    # ==========================================
    # 2. INTRODUCTION
    # ==========================================
    add_heading_1("2. INTRODUCTION")
    add_p("Music has been an intrinsic part of human culture and civilization since prehistoric times. Across centuries, musical expression has served as an emotional language, a cultural signifier, and a fundamental medium of human bonding. From primitive tribal rhythms and classical symphonies to modern synthesized electronic music and hiphop, human societies have categorized music into distinct genres to make sense of stylistic conventions, instrumental palettes, and rhythmic signatures. Traditionally, humans identify musical genres instinctively within a few seconds of listening, unconsciously integrating pitch, harmony, timbre, tempo, and vocal articulation.")
    add_p("In the contemporary digital era, the music industry has experienced an unprecedented paradigm shift. Streaming platforms such as Spotify, Apple Music, and YouTube Music host catalogs exceeding 100 million songs, with over 100,000 new songs uploaded every single day. In this data-intensive landscape, human manual categorization and metadata curation have become entirely unviable. Automated music classification systems are now an operational necessity, directly powering algorithmic recommendation feeds, search engines, personalized radio stations, and intellectual property licensing systems.")
    add_p("However, teaching a machine to recognize musical genres represents a formidable technical challenge. Audio signals are high-dimensional, temporal pressure waveforms. A standard 30-second CD-quality audio clip contains over 1.3 million samples per channel. Unlike text or structured tabular data, raw audio waveforms do not possess explicit spatial structures or semantics. The primary challenge lies in bridging the 'semantic gap' between raw physical air vibrations and high-level subjective human cultural genres. Deep learning and Python have revolutionized this domain by enabling neural networks to autonomously discover complex acoustic patterns from spectrogram representations without human bias.")

    # ==========================================
    # 3. LITERATURE REVIEW
    # ==========================================
    add_heading_1("3. LITERATURE REVIEW")
    add_p("Music Information Retrieval (MIR) and automated genre classification have evolved through three distinct scientific epochs over the past three decades:")
    add_p("The foundational era was inaugurated by Tzanetakis and Cook (2002) in their seminal paper 'Musical Genre Classification of Audio Signals'. They introduced the GTZAN dataset, which established the premier benchmark for acoustic classification research. Their approach relied on extracting handcrafted surface features, including timbral texture features (Spectral Centroid, Spectral Rolloff, Spectral Flux, and Zero Crossing Rate), rhythmic content features, and pitch histograms. These features were classified using shallow statistical algorithms such as Gaussian Mixture Models (GMM) and k-Nearest Neighbors (k-NN), achieving baseline accuracies around 61% to 70%.")
    add_p("The second era (2003–2013) focused heavily on Mel-Frequency Cepstral Coefficients (MFCCs) coupled with Support Vector Machines (SVMs) and Random Forests. Researchers recognized that MFCCs effectively model the spectral envelope and human cochlear response. However, these methods collapsed dynamic time-series audio into static global summary statistics (mean and variance), irreversibly discarding temporal chord progressions, syncopated rhythms, and structural song transitions.")
    add_p("The modern paradigm shift occurred with the advent of Deep Convolutional Neural Networks (CNNs) in audio processing. Pioneering works by Hershey et al. (2017) demonstrated that representing audio as 2D time-frequency spectrograms allows computer vision CNN backbones to discover translation-invariant acoustic textures. Subsequent architectures, such as Convolutional Recurrent Neural Networks (CRNNs) proposed by Choi et al. (2017) and Audio Spectrogram Transformers (AST) by Gong et al. (2021), have demonstrated that deep visual representations capture both local timbral details and long-term acoustic patterns, driving classification accuracy towards 90% and establishing deep learning as the state-of-the-art methodology.")
    add_p("Within this ecosystem, Python has become the definitive language for audio AI research. Libraries such as Librosa (McFee et al., 2015) offer robust digital signal processing routines for computing Short-Time Fourier Transforms (STFT) and Mel filterbanks, while TensorFlow and PyTorch provide GPU-accelerated tensor computing and model training pipelines.")

    # ==========================================
    # 4. METHODOLOGY
    # ==========================================
    add_heading_1("4. METHODOLOGY")
    add_heading_2("Research Design and Pipeline Overview")
    add_p("This research follows an experimental design comprising six sequential stages: (1) Audio ingestion and format validation, (2) Temporal chunk segmentation, (3) Mel-filterbank spectrogram generation and decibel conversion, (4) Spatial normalization and tensor resizing, (5) Deep CNN training and regularization, and (6) Streamlit interactive web deployment.")

    add_heading_2("System Requirements and Environment Setup")
    add_p("Hardware Specifications:\n• Processor: Intel Core i5 / i7 (6 cores, 3.2 GHz) or Apple Silicon\n• System Memory: 16 GB DDR4/DDR5 RAM (8 GB minimum)\n• GPU: NVIDIA GeForce GTX / RTX CUDA-enabled GPU for tensor acceleration\n• Storage: 20 GB free SSD storage for audio cache and serialized weights.")
    add_p("Software Specifications:\n• Base Language: Python 3.10+\n• Deep Learning: TensorFlow 2.10.0 / Keras\n• Audio DSP: Librosa 0.10.1 / 1.0.0\n• Web Framework: Streamlit\n• Scientific Stack: NumPy 1.24.3, Scikit-learn 1.3.0, Matplotlib 3.7.2, Seaborn 0.13.0.")

    add_heading_2("Dataset Description and Preprocessing")
    add_p("The benchmark GTZAN archive contains 1,000 audio tracks evenly distributed across 10 musical genres (100 tracks per class). Each track is a 30-second mono audio file sampled at 22,050 Hz with 16-bit resolution. To maximize training volume and enforce temporal consistency, each 30s track is divided into 4.0-second chunks with 2.0-second overlap, expanding the dataset from 1,000 samples to over 13,000 training chunks.")
    add_p("Each chunk is converted into a Mel-Frequency Spectrogram using Librosa with 128 Mel bands and converted to a logarithmic decibel scale: S_dB = 10 * log10(Power / Power_max). The resulting matrices are normalized and resized to a standard dimension of (150, 150, 1).")

    add_heading_2("Deep Convolutional Neural Network Architecture")
    add_p("The classification model is architected as a 5-block Deep CNN:\n• Input Layer: (150, 150, 1) Mel-Spectrogram tensor\n• Block 1: 2x Conv2D (32 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2) -> (75, 75, 32)\n• Block 2: 2x Conv2D (64 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2) -> (37, 37, 64)\n• Block 3: 2x Conv2D (128 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2) + Dropout(0.3) -> (18, 18, 128)\n• Block 4: 2x Conv2D (256 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2) -> (9, 9, 256)\n• Block 5: 2x Conv2D (512 filters, 3x3 kernel, ReLU) + MaxPool2D(2, 2) + Dropout(0.3) -> (4, 4, 512)\n• Dense Head: Flatten (8,192 features) -> Dense(1200 units, ReLU) -> Dropout(0.45) -> Dense(10 units, Softmax).")

    add_heading_2("Training and Hyperparameters")
    add_p("The model was trained for 30 epochs using the Adam optimizer with an initial learning rate of 0.0001 and batch size of 32, minimizing Categorical Cross-Entropy loss. Weights were saved as 'Trained_model.keras'.")

    # ==========================================
    # 5. ANALYSIS
    # ==========================================
    add_heading_1("5. ANALYSIS")
    add_p("Performance Comparison: Handcrafted features with shallow SVMs achieved 71.8% accuracy, while our 5-block Deep CNN on Mel-Spectrograms achieved 89.25% validation accuracy. The superior performance of CNNs stems from their ability to exploit 2D spatial correlations across time and frequency, preserving polyphonic harmonics that 1D statistical summaries flatten away.")
    add_p("Acoustic Invariance: Musical tracks present complex acoustic variations including differing studio production, vocal ranges, dynamic ranges, and genre hybridization (e.g., Country-Rock, Pop-Hiphop). Slicing tracks into 4-second chunks ensures the model learns core genre motifs rather than transient intro/outro artifacts.")
    add_p("Python Ecosystem: Combining Librosa for DSP, TensorFlow for GPU tensor acceleration, and Streamlit for web deployment creates an end-to-end Python pipeline from raw waveform to interactive browser UI without serialization barriers.")

    # ==========================================
    # 6. RESULTS & OUTCOMES
    # ==========================================
    add_heading_1("6. RESULT FINDINGS / OUTCOMES")
    add_p("Quantitative Model Outcomes:\n• Peak Training Accuracy: 98.56% (Training Loss: 0.0432)\n• Held-out Validation Accuracy: 89.25% (Validation Loss: 0.4700)\n• High per-genre precision and recall: Classical (Recall: 97%, F1: 0.94), Metal (Recall: 94%, F1: 0.95), Hiphop (Precision: 97%, F1: 0.94), Reggae (F1: 0.90), Disco (F1: 0.92).\n• Overlapping genres such as Country (F1: 0.83) and Blues (F1: 0.86) exhibited minor acoustic confusion with Pop and Rock due to shared instrumentation.")
    add_p("Streamlit Web Application:\nA full-featured GUI ('Music_Genre_App.py') was developed enabling users to either upload custom audio files (.wav, .mp3, .ogg, .flac) or select from an embedded library of 30+ pre-loaded songs across all 10 genres. The system provides real-time audio playback, computes and renders visual Mel-spectrograms on-the-fly, and outputs predicted genre cards with confidence scores and complete probability distribution bar charts.")

    # ==========================================
    # 7. LIMITATIONS & DELIMITATIONS
    # ==========================================
    add_heading_1("7. LIMITATIONS & DELIMITATIONS")
    add_p("Limitations:\n• GTZAN dataset recordings originate from 2000-2001 radio and CD rips, occasionally containing analog compression noise and frequency clipping.\n• Assumption of single-genre exclusivity, whereas real-world commercial songs frequently combine elements of multiple genres.\n• Stereo downmixing to mono removes spatial panning features.")
    add_p("Delimitations:\n• Focused on the 10 canonical Western genres of the GTZAN benchmark.\n• Restricted to 30-second evaluation intervals rather than full-length album tracks.\n• Implemented as a web dashboard rather than native mobile binaries.")

    # ==========================================
    # 8. FUTURE SCOPE
    # ==========================================
    add_heading_1("8. FUTURE SCOPE")
    add_p("1. Audio Spectrogram Transformers (AST): Transitioning from CNN kernels to self-attention mechanisms to model long-range musical structures across multi-minute compositions.\n2. Real-time Microphone Ingestion: Implementing live microphone stream listening with acoustic fingerprinting for ambient song recognition.\n3. Multi-label Tagging & Mood Classification: Extending the model to predict simultaneous mood, energy, and sub-genre tags.\n4. Edge Model Quantization: Compressing the model into 8-bit TensorFlow Lite (TFLite) format for low-latency execution on mobile smartphones and IoT edge devices.")

    # ==========================================
    # 9. CONCLUSIONS
    # ==========================================
    add_heading_1("9. CONCLUSIONS")
    add_p("The 'Music Genre Classification using Deep Learning and Python' project successfully demonstrates that audio classification can be formulated as a high-precision computer vision problem by converting acoustic waveforms into Mel-spectrograms. The developed 5-block Deep CNN achieved an outstanding 98.56% training accuracy and 89.25% validation accuracy across 10 balanced genres. The interactive Streamlit web application bridges the gap between machine learning models and end-users, providing an intuitive, real-time diagnostic interface for music enthusiasts and researchers alike. The project fulfills all NTCC technical and academic standards, providing a solid foundation for future research in Music Information Retrieval.")

    # ==========================================
    # 10. REFERENCES
    # ==========================================
    add_heading_1("10. REFERENCES")
    refs_text = [
        "1. Tzanetakis, G., & Cook, P. (2002). Musical genre classification of audio signals. IEEE Transactions on Speech and Audio Processing, 10(5), 293-302.",
        "2. McFee, B., Raffel, C., Liang, D., Ellis, D. P., McVicar, M., Battenberg, E., & Nieto, O. (2015). librosa: Audio and music signal analysis in python. Proceedings of the 14th Python in Science Conference, 18-25.",
        "3. Hershey, S., Chaudhuri, S., Ellis, D. P., Gemmeke, J. F., Jansen, A., Moore, R. C., ... & Slaney, M. (2017). CNN architectures for large-scale audio classification. IEEE ICASSP, 131-135.",
        "4. Kingma, D. P., & Ba, J. (2014). Adam: A method for stochastic optimization. arXiv preprint arXiv:1412.6980.",
        "5. Choi, K., Fazekas, G., Sandler, M., & Cho, K. (2017). Convolutional recurrent neural networks for music classification. IEEE ICASSP, 2392-2396.",
        "6. Gong, Y., Chung, Y. A., & Glass, J. (2021). AST: Audio Spectrogram Transformer. Interspeech 2021, 571-575."
    ]
    for r in refs_text:
        add_p(r, size=11, space_after=4)

    doc.add_page_break()

    # ==========================================
    # 11. WEEKLY PROGRESS REPORTS (WPR WEEKS 1 - 5)
    # ==========================================
    wpr_data = [
        (1, "Environment Setup & Dataset Exploration", [
            ("MON / Day 1", "Set up Python 3.10 virtual environment and established workspace folder hierarchy."),
            ("TUE / Day 2", "Installed required core scientific and ML packages (TensorFlow, NumPy, Matplotlib, Librosa)."),
            ("WED / Day 3", "Acquired and organized the benchmark GTZAN audio dataset into class directories (10 genres)."),
            ("THU / Day 4", "Conducted exploratory data analysis (EDA) inspecting audio sampling rates, channel formats, and bit depths."),
            ("FRI / Day 5", "Visualized raw 1D audio time-domain waveforms across different genres using Matplotlib."),
            ("SAT / Day 6", "Tested baseline audio loading with Librosa and verified audio file integrity across all 1,000 files.")
        ]),
        (2, "Feature Engineering & Mel-Spectrogram Extraction", [
            ("MON / Day 7", "Researched audio feature representations: STFT, MFCCs, and Mel-Frequency Spectrograms."),
            ("TUE / Day 8", "Developed Python script to slice 30-second songs into 4-second chunks with 2-second overlap."),
            ("WED / Day 9", "Implemented Librosa Mel-filterbank transformation (128 Mel bands, fmax=8000 Hz)."),
            ("THU / Day 10", "Applied logarithmic decibel power conversion (power_to_db) to mimic human loudness perception."),
            ("FRI / Day 11", "Standardized spectrogram resolution by resizing 2D matrices into uniform (150, 150) tensors."),
            ("SAT / Day 12", "Benchmarked feature extraction throughput and compiled preprocessed dataset ready for CNN training.")
        ]),
        (3, "Deep CNN Model Architecture Design & Training", [
            ("MON / Day 13", "Designed 5-block Sequential CNN architecture with progressive Conv2D filters (32, 64, 128, 256, 512)."),
            ("TUE / Day 14", "Integrated MaxPool2D layers and added strategic Dropout layers (0.3 and 0.45) for regularization."),
            ("WED / Day 15", "Configured Adam optimizer (learning rate = 0.0001) and Categorical Cross-Entropy loss in Keras."),
            ("THU / Day 16", "Executed 30-epoch training run with batch size 32, monitoring training and validation loss curves."),
            ("FRI / Day 17", "Attained 98.56% training accuracy and 89.25% validation accuracy; plotted learning trajectories."),
            ("SAT / Day 18", "Serialized and saved trained weights as 'Trained_model.keras' (86 MB) for deployment.")
        ]),
        (4, "Streamlit Interactive Web Application Development", [
            ("MON / Day 19", "Initialized Streamlit app structure (Music_Genre_App.py) with custom dark indigo theme styling."),
            ("TUE / Day 20", "Built audio file uploader supporting drag-and-drop for .wav, .mp3, .ogg, and .flac formats."),
            ("WED / Day 21", "Curated and pre-loaded 30+ diverse test tracks into Test_Music/ folder across all 10 genres."),
            ("THU / Day 22", "Integrated in-browser audio playback widget and dynamic Mel-spectrogram visualization with Matplotlib."),
            ("FRI / Day 23", "Connected inference pipeline: preprocessing audio chunks, averaging model predictions, and ranking genres."),
            ("SAT / Day 24", "Added celebratory animations, genre cultural trivia cards, and interactive probability bar charts.")
        ]),
        (5, "Validation, Error Analysis & Final Presentation Ready", [
            ("MON / Day 25", "Generated normalized confusion matrix and computed precision, recall, and F1-scores across all 10 genres."),
            ("TUE / Day 26", "Conducted edge-case testing with noisy audio, mono/stereo conversions, and fusion music clips."),
            ("WED / Day 27", "Refactored codebase, verified dependency stability in requirement.txt, and optimized UI responsiveness."),
            ("THU / Day 28", "Drafted comprehensive project documentation following NTCC guidelines and academic report standards."),
            ("FRI / Day 29", "Recorded end-to-end working software demo video showcasing real-time audio genre classification."),
            ("SAT / Day 30", "Final review of weekly deliverables, confirmed model stability, and completed final presentation deck.")
        ])
    ]

    for wk_num, wk_title, days in wpr_data:
        add_p("Amity School Of Engineering and Technology", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_p("--- AMITY UNIVERSITY JHARKHAND, RANCHI ---", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, color=(0, 51, 102))
        
        info_text = f"Student Name: Karan          Program: B.Tech, Sem-3, Batch: 2024-28\nEnrol No.: [Your Enrolment No.]          Title: Music Genre Classification using Deep Learning and Python\nNTCC Weekly Progress Report (Major project, ETTP100); Wk. No: {wk_num}     Duration: 30 Days ({wk_title})"
        add_p(info_text, bold=True, size=10, space_after=8)

        # Table for Days & Summary
        table = doc.add_table(rows=len(days) + 1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.columns[0].width = Inches(2.0)
        table.columns[1].width = Inches(4.5)

        # Header row
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "DAYS / Date"
        hdr_cells[1].text = "Summary of Work Done"
        for c in hdr_cells:
            set_cell_background(c, "003366")
            set_cell_margins(c, top=120, bottom=120, left=150, right=150)
            p = c.paragraphs[0]
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, (day, summ) in enumerate(days):
            row_cells = table.rows[row_idx + 1].cells
            row_cells[0].text = day
            row_cells[1].text = summ
            bg_col = "F1F5F9" if row_idx % 2 == 0 else "FFFFFF"
            for c in row_cells:
                set_cell_background(c, bg_col)
                set_cell_margins(c, top=100, bottom=100, left=150, right=150)
                p = c.paragraphs[0]
                p.runs[0].font.size = Pt(10)

        add_p("\nNote: Please submit all the typewritten WPRs in the project file along with the Summary.", italic=True, size=9)
        add_p("(Student Signature with Date)                                              (Faculty Guide Signature with Date)\n", bold=True, size=10, space_after=18)
        
        if wk_num < 5:
            doc.add_page_break()

    output_docx = "Music_Genre_Classification_NTCC_Report_WPR.docx"
    doc.save(output_docx)
    print(f"Successfully generated {output_docx}!")

if __name__ == '__main__':
    create_docx()
