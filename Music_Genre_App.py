import os
import streamlit as st
import tensorflow as tf
import numpy as np
import librosa
import matplotlib.pyplot as plt
from tensorflow.image import resize
import pandas as pd

# Page Configuration
st.set_page_config(
    page_title="Music Genre Classification System",
    page_icon="🎶",
    layout="wide"
)

# Function to load model
@st.cache_resource
def load_model():
    model = tf.keras.models.load_model("Trained_model.keras")
    return model

# Function to generate Mel-Spectrogram figure
def generate_spectrogram_plot(file_path):
    y, sr = librosa.load(file_path, sr=None)
    mel_spectrogram = librosa.feature.melspectrogram(y=y, sr=sr)
    mel_db = librosa.power_to_db(mel_spectrogram, ref=np.max)
    
    fig, ax = plt.subplots(figsize=(10, 3.5), facecolor='#181646')
    ax.set_facecolor('#0f0d2c')
    img = ax.imshow(mel_db, aspect='auto', origin='lower', cmap='magma')
    cbar = fig.colorbar(img, ax=ax, format="%+2.0f dB")
    cbar.ax.yaxis.set_tick_params(color='white')
    plt.setp(plt.getp(cbar.ax.axes, 'yticklabels'), color='white')
    
    ax.set_title("Audio Mel-Frequency Spectrogram Analysis", color='#a5f3fc', fontsize=13, fontweight='bold')
    ax.set_xlabel("Time Frames", color='white', fontsize=10)
    ax.set_ylabel("Mel Frequency Bands", color='white', fontsize=10)
    ax.tick_params(colors='white')
    for spine in ax.spines.values():
        spine.set_color('#4f46e5')
    plt.tight_layout()
    return fig

# Load and preprocess audio data for prediction
def load_and_preprocess_data(file_path, target_shape=(150, 150)):
    data = []
    audio_data, sample_rate = librosa.load(file_path, sr=None)
    chunk_duration = 4  # seconds
    overlap_duration = 2  # seconds
                
    chunk_samples = int(chunk_duration * sample_rate)
    overlap_samples = int(overlap_duration * sample_rate)
                
    num_chunks = int(np.ceil((len(audio_data) - chunk_samples) / (chunk_samples - overlap_samples))) + 1
                
    for i in range(num_chunks):
        start = i * (chunk_samples - overlap_samples)
        end = start + chunk_samples
                    
        chunk = audio_data[start:end]
        if len(chunk) < chunk_samples:
            chunk = np.pad(chunk, (0, chunk_samples - len(chunk)))
                    
        mel_spectrogram = librosa.feature.melspectrogram(y=chunk, sr=sample_rate)
        mel_spectrogram = resize(np.expand_dims(mel_spectrogram, axis=-1), target_shape)
        data.append(mel_spectrogram)
    
    return np.array(data)

# Tensorflow Model Prediction
def model_prediction(X_test):
    model = load_model()
    y_pred = model.predict(X_test)
    avg_probs = np.mean(y_pred, axis=0)
    predicted_categories = np.argmax(y_pred, axis=1)
    unique_elements, counts = np.unique(predicted_categories, return_counts=True)
    max_count = np.max(counts)
    max_elements = unique_elements[counts == max_count]
    return max_elements[0], avg_probs

# Sidebar Dashboard
st.sidebar.title("Dashboard")
app_mode = st.sidebar.selectbox("Select Page", ["Home", "About Project", "Prediction"])

# Genre descriptions for educational insight
GENRE_DESCRIPTIONS = {
    'blues': 'Deeply emotional roots music characterized by call-and-response vocals, 12-bar blues progressions, and expressive guitar bends.',
    'classical': 'Sophisticated orchestral compositions featuring intricate harmonies, dynamic symphonies, and acoustic instrumentation.',
    'country': 'Storytelling folk tradition often utilizing acoustic guitars, banjos, fiddles, and vocal twangs.',
    'disco': 'Uptempo four-on-the-floor dance music with synchronized brass, electric rhythm guitars, and synthesizer grooves.',
    'hiphop': 'Rhythmic lyrical delivery over syncopated beats, heavy bass lines, sampling, and drum breaks.',
    'jazz': 'Harmonically complex improvisational style with swing rhythms, brass instruments, and extended chord voicings.',
    'metal': 'Energetic and aggressive music driven by heavily distorted electric guitars, fast double-kick drums, and powerful vocals.',
    'pop': 'Catchy melodic hooks, verse-chorus structures, polished modern production, and radio-friendly arrangements.',
    'reggae': 'Jamaican musical style with offbeat rhythm guitar chops (skank), deep prominent bass lines, and relaxed tempo.',
    'rock': 'High-energy guitar-driven anthems with driving drums, bass grooves, and powerful vocal hooks.'
}

# ----------------- HOME PAGE -----------------
if app_mode == "Home":
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #181646;
            color: white;
        }
        h1, h2, h3, p, span, label {
            color: white !important;
        }
        .stats-badge {
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 10px;
            padding: 10px;
            text-align: center;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(''' ## Welcome to the,\n## Music Genre Classification System! 🎶🎧''')
    image_path = "music_genre_home.png"
    if os.path.exists(image_path):
        st.image(image_path, use_container_width=True)

    colA, colB, colC = st.columns(3)
    with colA:
        st.markdown('<div class="stats-badge"><h3>10</h3><p>Supported Genres</p></div>', unsafe_allow_html=True)
    with colB:
        st.markdown('<div class="stats-badge"><h3>30+</h3><p>Pre-loaded Sample Songs</p></div>', unsafe_allow_html=True)
    with colC:
        st.markdown('<div class="stats-badge"><h3>Deep CNN</h3><p>Spectrogram Feature Model</p></div>', unsafe_allow_html=True)

    st.markdown("""
**Our goal is to help in identifying music genres from audio tracks efficiently. Upload an audio file, and our system will analyze it to detect its genre. Discover the power of AI in music analysis!**

### How It Works
1. **Upload or Select Audio:** Go to the **Prediction** page and pick from 30+ pre-loaded tracks or upload your own audio.
2. **Analysis:** Our system converts audio into Mel-spectrogram frequency chunks and classifies them using a Deep Convolutional Neural Network.
3. **Results:** View the predicted genre, confidence metric, visual spectrogram, and the full genre probability breakdown.

### Why Choose Us?
- **Accuracy:** Leverages deep learning models trained on audio spectrogram representations.
- **Visual Analytics:** View real-time audio Mel-spectrogram frequency maps.
- **User-Friendly & Fast:** Simple interface with pre-loaded library tracks and custom file uploads.

### Get Started
Click on the **Prediction** page in the sidebar to test songs and explore the magic of our AI system!
""")

# ----------------- ABOUT PROJECT PAGE -----------------
elif app_mode == "About Project":
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #181646;
            color: white;
        }
        h1, h2, h3, h4, p, span, label, li {
            color: white !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    st.markdown("""
    ### About Project
    Music. Experts have been trying for a long time to understand sound and what differentiates one song from another. How to visualize sound. What makes a tone different from another.

    This project leverages deep learning techniques to classify audio files automatically based on Mel Spectrogram representations.

    ### About Dataset
    #### Content
    1. **genres original** - A collection of 10 genres with 100 audio files each, all having a length of 30 seconds (the famous GTZAN dataset, the MNIST of sounds).
    2. **List of Genres** - blues, classical, country, disco, hiphop, jazz, metal, pop, reggae, rock.
    3. **images original** - A visual representation for each audio file. One way to classify data is through neural networks. Because Convolutional Neural Networks (CNNs) take image representations, the audio files are converted to Mel Spectrograms to make this classification possible.
    4. **2 CSV files** - Containing extracted features of the audio files over 30-second and 3-second intervals.
    """)

# ----------------- PREDICTION PAGE -----------------
elif app_mode == "Prediction":
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #181646;
            color: white;
        }
        h1, h2, h3, h4, p, span, label {
            color: white !important;
        }
        .prediction-box {
            background: linear-gradient(135deg, #1e1b4b, #4338ca);
            border-radius: 14px;
            padding: 1.5rem;
            margin-top: 1rem;
            margin-bottom: 1.5rem;
            border: 1px solid rgba(255, 255, 255, 0.2);
            text-align: center;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    st.header("🎵 Music Genre Prediction")

    os.makedirs("Test_Music", exist_ok=True)
    all_sample_files = sorted([f for f in os.listdir("Test_Music") if f.endswith(('.wav', '.mp3', '.ogg', '.flac'))])

    audio_source = st.radio("Choose Audio Input Method:", ["Select From Library (30+ Songs Across All 10 Genres)", "Upload Your Own Audio File"], horizontal=True)

    filepath = None
    audio_bytes = None

    if audio_source == "Select From Library (30+ Songs Across All 10 Genres)":
        genre_filter = st.selectbox(
            "Filter Songs by Genre:",
            ["All Genres", "Blues", "Classical", "Country", "Disco", "Hiphop", "Jazz", "Metal", "Pop", "Reggae", "Rock"]
        )
        
        if genre_filter == "All Genres":
            filtered_files = all_sample_files
        else:
            filtered_files = [f for f in all_sample_files if f.lower().startswith(genre_filter.lower())]

        if filtered_files:
            selected_sample = st.selectbox("Select Audio Track:", filtered_files)
            filepath = os.path.join("Test_Music", selected_sample)
            try:
                with open(filepath, "rb") as f:
                    audio_bytes = f.read()
            except Exception as e:
                st.error(f"Error loading file: {e}")
        else:
            st.info(f"No tracks found for genre '{genre_filter}'.")

    else:
        test_audio = st.file_uploader("Upload an audio file (.mp3, .wav, .ogg, .flac)", type=["mp3", "wav", "ogg", "flac"])
        if test_audio is not None:
            filepath = os.path.join("Test_Music", test_audio.name)
            with open(filepath, "wb") as f:
                f.write(test_audio.getbuffer())
            audio_bytes = test_audio.getvalue()

    col1, col2 = st.columns([1, 1])

    with col1:
        play_clicked = st.button("🎧 Play Audio & View Spectrogram", use_container_width=True)

    with col2:
        predict_clicked = st.button("✨ Predict Genre", use_container_width=True)

    if play_clicked:
        if audio_bytes is not None and filepath is not None:
            st.audio(audio_bytes)
            with st.spinner("Generating audio spectrogram..."):
                fig = generate_spectrogram_plot(filepath)
                st.pyplot(fig)
        else:
            st.warning("Please upload or select an audio file first.")

    if predict_clicked:
        if filepath is not None and os.path.exists(filepath):
            with st.spinner("Analyzing audio frequencies and generating spectrogram chunks..."):
                try:
                    X_test = load_and_preprocess_data(filepath)
                    result_index, avg_probs = model_prediction(X_test)
                    st.balloons()
                    
                    label = ['blues', 'classical', 'country', 'disco', 'hiphop', 'jazz', 'metal', 'pop', 'reggae', 'rock']
                    predicted_genre = label[result_index]
                    confidence = float(avg_probs[result_index]) * 100

                    st.markdown(
                        f"""
                        <div class="prediction-box">
                            <h3 style="margin: 0; color: #a5f3fc !important;">Prediction Result</h3>
                            <h1 style="font-size: 2.8rem; margin: 0.5rem 0; color: #facc15 !important;">🎶 {predicted_genre.upper()}</h1>
                            <p style="font-size: 1.1rem; color: #ffffff !important; margin-bottom: 0.5rem;">Overall Model Confidence: <strong>{confidence:.1f}%</strong></p>
                            <p style="font-size: 0.95rem; color: #e2e8f0 !important; max-width: 700px; margin: 0 auto;"><em>{GENRE_DESCRIPTIONS.get(predicted_genre, '')}</em></p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    st.audio(audio_bytes)

                    col_chart, col_spec = st.columns([1, 1])
                    with col_chart:
                        st.subheader("📊 Genre Probability Distribution")
                        prob_df = pd.DataFrame({
                            "Genre": [g.capitalize() for g in label],
                            "Probability (%)": [round(float(p) * 100, 2) for p in avg_probs]
                        }).set_index("Genre")
                        st.bar_chart(prob_df)

                    with col_spec:
                        st.subheader("📈 Frequency Spectrogram")
                        fig = generate_spectrogram_plot(filepath)
                        st.pyplot(fig)

                except Exception as e:
                    st.error(f"Error during audio processing: {e}")
        else:
            st.warning("Please upload an audio file or select a sample track first.")
