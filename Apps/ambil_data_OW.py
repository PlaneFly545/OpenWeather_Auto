import requests
import pandas as pd
from datetime import datetime, timedelta
import os
import time

# Konfigurasi API dan lokasi
API_KEY = 'e6b915aa2e053071d457d47bbfa031b1'
LAT = -6.1781
LON = 106.63
URL = f'https://api.openweathermap.org/data/2.5/weather?lat={LAT}&lon={LON}&appid={API_KEY}&units=metric'

# Path file Excel
FILE_PATH = r"./Data/data_harian_OP.xlsx"

# Fungsi untuk mengambil data dari API
def fetch_weather_data():
    try:
        response = requests.get(URL)
        if response.status_code == 200:
            data = response.json()
            suhu = data['main']['temp']
            kelembapan = data['main']['humidity']
            kecepatan_angin = data['wind']['speed']
            
            # Waktu diubah ke UTC+7
            waktu_utc = datetime.utcnow() + timedelta(hours=7)
            waktu_str = waktu_utc.strftime('%Y-%m-%d %H:%M:%S')
            
            return {
                'tanggal': waktu_str,
                'suhu': suhu,
                'kelembapan': kelembapan,
                'kec_angin': kecepatan_angin
            }
        else:
            print(f"Error: Gagal mengambil data (HTTP {response.status_code})")
            return None
    except Exception as e:
        print(f"Error: {e}")
        return None

# Fungsi untuk menyimpan data ke Excel
def save_to_excel(data, file_path=FILE_PATH):
    if not os.path.exists(file_path):
        # Jika file belum ada, buat dataframe baru dengan header
        df = pd.DataFrame([data])
        df.to_excel(file_path, index=False, sheet_name='Data Cuaca')
        print(f"File {file_path} berhasil dibuat dengan data pertama.")
    else:
        # Jika file sudah ada, tambahkan data baru di bawahnya
        existing_data = pd.read_excel(file_path, sheet_name='Data Cuaca')
        df = pd.DataFrame([data])
        updated_data = pd.concat([existing_data, df], ignore_index=True)
        updated_data.to_excel(file_path, index=False, sheet_name='Data Cuaca')
        print(f"Data berhasil ditambahkan ke file {file_path}.")

# Main Program dengan interval 2 menit
if __name__ == "__main__":
    print("Memulai pengambilan data cuaca setiap 2 menit...")
    while True:
        weather_data = fetch_weather_data()
        if weather_data:
            print(f"Data berhasil diambil pada {weather_data['tanggal']}")
            print(weather_data)
            save_to_excel(weather_data)
        else:
            print("Gagal mengambil data cuaca.")
        
        # Tunggu selama 2 menit sebelum iterasi berikutnya
        time.sleep(120)
