import os
import numpy as np
import librosa
from tensorflow.image import resize
import tensorflow as tf
from tensorflow.keras.layers import Conv2D, MaxPool2D, Dropout, Flatten, Dense
from tensorflow.keras.utils import to_categorical

data_dir = r"C:\Users\karan\Downloads\archive (3)\Data\genres_original"
classes = ['blues', 'classical', 'country', 'disco', 'hiphop', 'jazz', 'metal', 'pop', 'reggae', 'rock']

print("Loading audio samples from dataset...")
data = []
labels = []
target_shape = (150, 150)
chunk_duration = 4
overlap_duration = 2

# We take 2-3 songs per genre to quickly train and produce a functioning Trained_model.keras
for i_class, class_name in enumerate(classes):
    class_dir = os.path.join(data_dir, class_name)
    if not os.path.exists(class_dir):
        print(f"Warning: {class_dir} not found")
        continue
    wav_files = [f for f in os.listdir(class_dir) if f.endswith('.wav')][:3]
    for filename in wav_files:
        file_path = os.path.join(class_dir, filename)
        try:
            audio_data, sample_rate = librosa.load(file_path, sr=None)
            chunk_samples = chunk_duration * sample_rate
            overlap_samples = overlap_duration * sample_rate
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
                labels.append(i_class)
        except Exception as e:
            print(f"Error processing {file_path}: {e}")

X = np.array(data)
Y = to_categorical(np.array(labels), num_classes=len(classes))
print(f"Loaded dataset: X shape = {X.shape}, Y shape = {Y.shape}")

model = tf.keras.models.Sequential([
    Conv2D(32, 3, padding='same', activation='relu', input_shape=(150, 150, 1)),
    Conv2D(32, 3, activation='relu'),
    MaxPool2D(2, 2),
    Conv2D(64, 3, padding='same', activation='relu'),
    Conv2D(64, 3, activation='relu'),
    MaxPool2D(2, 2),
    Conv2D(128, 3, padding='same', activation='relu'),
    Conv2D(128, 3, activation='relu'),
    MaxPool2D(2, 2),
    Dropout(0.3),
    Conv2D(256, 3, padding='same', activation='relu'),
    Conv2D(256, 3, activation='relu'),
    MaxPool2D(2, 2),
    Conv2D(512, 3, padding='same', activation='relu'),
    Conv2D(512, 3, activation='relu'),
    MaxPool2D(2, 2),
    Dropout(0.3),
    Flatten(),
    Dense(1200, activation='relu'),
    Dropout(0.45),
    Dense(len(classes), activation='softmax')
])

model.compile(optimizer=tf.keras.optimizers.Adam(learning_rate=0.0001), loss='categorical_crossentropy', metrics=['accuracy'])
print("Training model for 4 epochs...")
model.fit(X, Y, epochs=4, batch_size=32, verbose=1)

output_path = r"c:\Users\karan\OneDrive\Desktop\ntcc_project\Trained_model.keras"
model.save(output_path)
print(f"Model saved successfully to: {output_path}")
