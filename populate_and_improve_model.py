import os
import shutil
import numpy as np
import librosa
from tensorflow.image import resize
import tensorflow as tf
from tensorflow.keras.layers import Conv2D, MaxPool2D, Dropout, Flatten, Dense
from tensorflow.keras.utils import to_categorical

data_dir = r"C:\Users\karan\Downloads\archive (3)\Data\genres_original"
test_dir = r"c:\Users\karan\OneDrive\Desktop\ntcc_project\Test_Music"
os.makedirs(test_dir, exist_ok=True)

classes = ['blues', 'classical', 'country', 'disco', 'hiphop', 'jazz', 'metal', 'pop', 'reggae', 'rock']

print("1. Copying diverse sample songs for all 10 genres to Test_Music/...")
# Copy 3 representative songs per genre into Test_Music for immediate user testing
test_indices = [10, 25, 40]
copied_count = 0
for class_name in classes:
    class_folder = os.path.join(data_dir, class_name)
    if not os.path.exists(class_folder):
        continue
    wav_files = sorted([f for f in os.listdir(class_folder) if f.endswith('.wav')])
    for idx in test_indices:
        if idx < len(wav_files):
            src = os.path.join(class_folder, wav_files[idx])
            dst = os.path.join(test_dir, wav_files[idx])
            shutil.copy2(src, dst)
            copied_count += 1

print(f"Successfully copied {copied_count} test songs across all 10 genres to {test_dir}!")

print("\n2. Extracting training features for model refinement...")
# Train on 10-12 different songs per genre (disjoint from test set)
train_data = []
train_labels = []
target_shape = (150, 150)
chunk_duration = 4
overlap_duration = 2

for i_class, class_name in enumerate(classes):
    class_folder = os.path.join(data_dir, class_name)
    if not os.path.exists(class_folder):
        continue
    wav_files = sorted([f for f in os.listdir(class_folder) if f.endswith('.wav')])
    # Take songs between index 50 and 65 (15 songs per genre)
    selected_train_files = [f for i, f in enumerate(wav_files) if 50 <= i < 62]
    print(f"Extracting features for genre: {class_name} ({len(selected_train_files)} tracks)...")
    for filename in selected_train_files:
        filepath = os.path.join(class_folder, filename)
        try:
            audio_data, sample_rate = librosa.load(filepath, sr=None)
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
                train_data.append(mel_spectrogram)
                train_labels.append(i_class)
        except Exception as e:
            print(f"Error reading {filename}: {e}")

X_train = np.array(train_data)
Y_train = to_categorical(np.array(train_labels), num_classes=len(classes))
print(f"\nTraining set prepared: {X_train.shape} samples across {len(classes)} genres.")

print("\n3. Training CNN model...")
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
model.fit(X_train, Y_train, epochs=6, batch_size=32, verbose=1, shuffle=True)

model_path = r"c:\Users\karan\OneDrive\Desktop\ntcc_project\Trained_model.keras"
model.save(model_path)
print(f"\nModel successfully saved to {model_path}!")
