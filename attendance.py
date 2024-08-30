import cv2
import face_recognition
import time
import os
import pandas as pd

# Inisialisasi
video_capture = cv2.VideoCapture(0)
known_face_encodings = []
known_face_names = []
face_folder = 'face'

# Muat gambar wajah dari folder 'face'
for filename in os.listdir(face_folder):
    if filename.endswith('.jpeg'):
        image_path = os.path.join(face_folder, filename)
        face_image = face_recognition.load_image_file(image_path)
        face_encoding = face_recognition.face_encodings(face_image)[0]
        known_face_encodings.append(face_encoding)
        known_face_names.append(os.path.splitext(filename)[0])

# Inisialisasi waktu dan durasi
start_times = {name: None for name in known_face_names}
total_times = {name: 0 for name in known_face_names}

# Muat data dari Excel jika ada
if os.path.exists('attendance.xlsx'):
    df = pd.read_excel('attendance.xlsx')
    for index, row in df.iterrows():
        name = row['Name']
        if name in total_times:
            total_times[name] = row['Duration (s)']

# Fungsi untuk menyimpan data ke Excel
def save_to_excel(data):
    df = pd.DataFrame(data, columns=['Name', 'Duration (s)'])
    df.to_excel('attendance.xlsx', index=False)

while True:
    ret, frame = video_capture.read()

    # Temukan semua wajah dalam frame
    face_locations = face_recognition.face_locations(frame)
    face_encodings = face_recognition.face_encodings(frame, face_locations)

    detected_faces = set()

    for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        name = "Unknown"

        if True in matches:
            first_match_index = matches.index(True)
            name = known_face_names[first_match_index]
            detected_faces.add(name)

            # Update waktu kehadiran
            current_time = time.time()
            if start_times[name] is None:
                start_times[name] = current_time
            else:
                total_times[name] += current_time - start_times[name]
                start_times[name] = current_time

            # Gambar kotak dan nama di sekitar wajah
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
            cv2.putText(frame, name, (left + 6, bottom - 6), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)

            # Tampilkan waktu kehadiran di setiap kotak wajah
            if name != "Unknown":
                time_text = f"Time: {total_times[name]:.2f}s"
                cv2.putText(frame, time_text, (left + 6, top - 6), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)

    # Cek wajah yang tidak terdeteksi lagi
    for name in known_face_names:
        if name not in detected_faces and start_times[name] is not None:
            # Simpan data ke Excel
            save_to_excel([[name, total_times[name]] for name in known_face_names])
            # Reset waktu mulai
            start_times[name] = None

    # Tampilkan hasil
    cv2.imshow('Video', frame)

    # Tekan 'q' untuk keluar
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Bersihkan
video_capture.release()
cv2.destroyAllWindows()

# Simpan semua data terakhir ke Excel
final_data = [[name, time] for name, time in total_times.items()]
save_to_excel(final_data)

print("Data kehadiran telah disimpan ke attendance.xlsx")