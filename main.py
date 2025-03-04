##########################################################################
#Modul Unzip File                                                        #
##########################################################################
import os
import zipfile
import pandas as pd
import cv2
import json
import qrcode
import shutil
import logging
#from tqdm import tqdm
from pyzbar.pyzbar import decode
from PIL import Image, ImageDraw, UnidentifiedImageError


# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Folder input dan output static
INPUT_FOLDER = "zip"
OUTPUT_FOLDER = "unzipped_files"


# Fungsi Validasi Gambar
def validate_image(image_path):
    try:
        with Image.open(image_path):
            return True
    except UnidentifiedImageError:
        return False

def extract_nested_zip(file_path, output_folder):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        temp_folder = os.path.join(output_folder, os.path.splitext(os.path.basename(file_path))[0])
        os.makedirs(temp_folder, exist_ok=True)
        zip_ref.extractall(temp_folder)

        for root, _, files in os.walk(temp_folder):
            for file in files:
                nested_file_path = os.path.join(root, file)
                if zipfile.is_zipfile(nested_file_path):
                    extract_nested_zip(nested_file_path, output_folder)

# Fungsi Proses Ekstraksi ZIP
# Fungsi Proses Ekstraksi ZIP
def process_zip_file(zip_file_name):
    zip_file_path = os.path.join(INPUT_FOLDER, zip_file_name)
    output_directory = os.path.join(OUTPUT_FOLDER, os.path.splitext(zip_file_name)[0])

    if not zipfile.is_zipfile(zip_file_path):
        logging.error(f"File '{zip_file_path}' bukan file ZIP yang valid.")
        return

    # Buat folder output
    os.makedirs(output_directory, exist_ok=True)

    try:
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(output_directory)
            logging.info(f"Berhasil mengekstrak file ZIP ke '{output_directory}'.")
    except zipfile.BadZipFile:
        logging.error(f"File '{zip_file_path}' rusak.")
        return

    # Membaca QR Code dari gambar
    qr_results = []
    for root, _, files in os.walk(output_directory):
        for file in files:
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                file_path = os.path.join(root, file)
                if validate_image(file_path):
                    qr_data = read_qr_code(file_path)
                    if qr_data:
                        qr_results.append({'file': file, 'qr_data': qr_data})
        

    # Menampilkan hasil
    if qr_results:
        logging.info("Hasil pembacaan QR Code:")
        for result in qr_results:
            logging.info(f"File: {result['file']} - QR Code: {result['qr_data']}")
    else:
        logging.warning("Tidak ditemukan QR Code yang valid.")

def batch_unzip(folder_path):
    # Periksa apakah folder ada
    if not os.path.isdir(folder_path):
        print(f"Folder '{folder_path}' tidak ditemukan.")
        return

    # Buat folder untuk menyimpan hasil ekstraksi
    output_folder = os.path.join(folder_path, "../unzipped_files")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Ambil semua file dalam folder
    files = os.listdir(folder_path)

    # Filter hanya file zip
    zip_files = [f for f in files if f.endswith('.zip')]

    if not zip_files:
        print("Tidak ada file zip di folder ini.")
        return

    # List untuk menyimpan data file, QR string, dan tarif
    data = []

    for zip_file in zip_files:
        zip_path = os.path.join(folder_path, zip_file)

        # Ekstrak file zip
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(output_folder)
                print(f"Berhasil mengekstrak '{zip_file}' ke folder '{output_folder}'.")
        except zipfile.BadZipFile:
            print(f"Gagal mengekstrak '{zip_file}': File zip rusak.")

    

    # Baca file hasil ekstraksi dan cari QR code
    for root, _, files in os.walk(output_folder):
        for file in files:
            file_path = os.path.join(root, file)
            qr_string = read_qr_code(file_path)
            tarif = determine_tarif(qr_string)
            data.append({"filename": file, "qrstring": qr_string, "tarif": tarif})

    # Simpan data ke file Excel
    excel_path = os.path.join(folder_path, "../listQr.xlsx")
    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)

    # Gunakan tqdm untuk menampilkan progress bar
    #with tqdm(total=len(df), desc="Menulis ke Excel", unit="baris") as pbar:
    #    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    #        for i, (idx, row) in enumerate(df.iterrows()):
    #            df.iloc[[idx]].to_excel(writer, index=False, header=(i == 0))
    #            pbar.update(1)

    print(f"Data berhasil disimpan ke file Excel: {excel_path}")

# Fungsi Membaca QR Code
def read_qr_code(image_path):
    try:
        image = cv2.imread(image_path)
        if image is None:
            logging.warning(f"File '{image_path}' bukan gambar yang valid.")
            return None

        decoded_objects = decode(image)
        if decoded_objects:
            return decoded_objects[0].data.decode('utf-8')
        else:
            return None
    except Exception as e:
        logging.error(f"Error membaca QR code dari file '{image_path}': {e}")
        return None
    
# Fungsi Proses ZIP File secara Batch
def process_all_zip_files():
    if not os.path.exists(INPUT_FOLDER):
        os.makedirs(INPUT_FOLDER)
        logging.info(f"Folder input '{INPUT_FOLDER}' dibuat. Silakan masukkan file ZIP ke folder ini.")
        return

    zip_files = [f for f in os.listdir(INPUT_FOLDER) if f.endswith('.zip')]

    if not zip_files:
        logging.warning(f"Tidak ada file ZIP di folder '{INPUT_FOLDER}'.")
        return

    # Buat folder output utama jika belum ada
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Proses setiap file ZIP
    for zip_file_name in zip_files:
        zip_file_path = os.path.join(INPUT_FOLDER, zip_file_name)

        if not zipfile.is_zipfile(zip_file_path):
            logging.error(f"File '{zip_file_path}' bukan file ZIP yang valid.")
            continue

        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(OUTPUT_FOLDER)
                logging.info(f"Berhasil mengekstrak file ZIP '{zip_file_name}' ke folder '{OUTPUT_FOLDER}'.")
        except zipfile.BadZipFile:
            logging.error(f"File '{zip_file_path}' rusak.")
            continue

    # Membaca QR Code dari semua file yang diekstrak
    qr_results = []
    for root, _, files in os.walk(OUTPUT_FOLDER):
        for file in files:
            file_path = os.path.join(root, file)
            if validate_image(file_path):
                qr_data = read_qr_code(file_path)
                if qr_data:
                    qr_results.append({'file': file, 'qr_data': qr_data})

    # Menampilkan hasil
    if qr_results:
        logging.info("Hasil pembacaan QR Code:")
        for result in qr_results:
            logging.info(f"File: {result['file']} - QR Code: {result['qr_data']}")
    else:
        logging.warning("Tidak ditemukan QR Code yang valid di file hasil ekstraksi.")



def determine_tarif(qr_string):
    if not qr_string:
        return None
    
    tarif_mapping = {
        "SBY REGULER": 6200,
        "SBY KHUSUS": 2000,
        "BMS REGULER": 3900,
        "BMS KHUSUS": 2000,
        "PLG REGULER": 4000,
        "PLG KHUSUS": 2000,
        "BPN REGULER": 4500,
        "BPN KHUSUS": 2000,
        "SKT REGULER": 3700,
        "SKT KHUSUS": 2000,
        "MKS REGULER": 4600,
        "MKS KHUSUS": 2000
    }

    for key, value in tarif_mapping.items():
        if key in qr_string:
            return value
    return None

##########################################################################
#hapus qr

def load_config(config_path):
    with open(config_path, 'r') as config_file:
        config = json.load(config_file)
    return config

def overlay_images(base_image_path, overlay_image_path, output_path, position):
    base_image = Image.open(base_image_path).convert("RGBA")
    overlay_image = Image.open(overlay_image_path).convert("RGBA")

    overlay_resized = overlay_image.resize((position['width'], position['height']))
    position_tuple = (position['x'], position['y'])

    combined = Image.new("RGBA", base_image.size)
    combined.paste(base_image, (0, 0))
    combined.paste(overlay_resized, position_tuple, mask=overlay_resized)

    combined.convert("RGB").save(output_path)

def process_images_hapusimages(excel_path, folder_path, overlay_image_path, output_folder, config_path):
    # Load configuration
    config = load_config(config_path)
    position = config.get('position', {'x': 0, 'y': 0, 'width': 100, 'height': 100})

    # Load Excel file
    df = pd.read_excel(excel_path)
    filenames = df['filename'].tolist()

    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Process each file
    for filename in filenames:
        base_image_path = os.path.join(folder_path, filename)
        output_path = os.path.join(output_folder, filename)

        if os.path.exists(base_image_path):
            try:
                overlay_images(base_image_path, overlay_image_path, output_path, position)
                print(f"Processed: {filename}")
            except Exception as e:
                print(f"Error processing {filename}: {e}")
        else:
            print(f"File not found: {filename}")

#############################################################################
# 3 Modify QR
def calculate_crc(data: bytes, polynomial: int = 0x1021, initial_value: int = 0xFFFF) -> str:
    """
    Menghitung nilai CRC-16 dengan polinomial tertentu. https://crccalc.com/ CRC-16/IBM-3740

    :param data: Data input dalam bentuk bytes.
    :param polynomial: Polinomial CRC yang digunakan (default: 0x1021).
    :param initial_value: Nilai awal register CRC (default: 0xFFFF).
    :return: Nilai CRC 4 digit dalam format hexadecimal.
    """
    crc = initial_value

    for byte in data:
        crc ^= (byte << 8)  # Masukkan byte ke register CRC
        for _ in range(8):
            if crc & 0x8000:  # Jika bit tertinggi adalah 1
                crc = (crc << 1) ^ polynomial
            else:
                crc <<= 1
            crc &= 0xFFFF  # Pastikan CRC tetap dalam 16-bit

    return f"{crc:04X}"  # Mengembalikan CRC dalam 4 digit hexadecimal

def edit_data_after_148th_char_tarif_and_crc(df: pd.DataFrame, data_column: str, tarif_column: str):
    """
    Mengedit data dengan menambahkan string "5404" setelah karakter ke-148, menambahkan nilai kolom tarif setelah karakter ke-148,
    menghapus 4 karakter terakhir, dan menambahkan nilai CRC setelah karakter terakhir pada setiap baris.

    :param df: DataFrame yang akan diedit.
    :param data_column: Nama kolom data yang akan diedit.
    :param tarif_column: Nama kolom tarif yang nilainya akan disisipkan.
    """
    if data_column in df.columns and tarif_column in df.columns:
        df[data_column] = df.apply(
            lambda row: (
                # row[data_column][:148] + "5404" 
                row[data_column][:10] + "12" + row[data_column][12:148] + "5404" + str(row[tarif_column]) + row[data_column][148:]
            )[:-4] if len(row[data_column]) > 4 else row[data_column] + "5404" + str(row[tarif_column]),
            axis=1
        )
        df[data_column] = df[data_column].astype(str).apply(
            lambda x: x + calculate_crc(x.encode('utf-8'))
        )
        print(f"Setiap baris defpada kolom '{data_column}' telah diperbarui dengan menambahkan '5404' setelah karakter ke-148, nilai tarif setelah karakter ke-148, menghapus 4 karakter terakhir, dan menambahkan nilai CRC di akhir baris.")
    else:
        print(f"Kolom '{data_column}' atau '{tarif_column}' tidak ditemukan dalam file Excel.")



#############################################################################

#############################################################################
#attach qr to aspi
def generate_qr_code(data):
    """Generate a QR code from a given string and return it as a PIL Image."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white")

def overlay_qr_on_image(image_path, qr_image, output_path, config):
    """Overlay a QR code on a specific position in an image and save it."""
    with Image.open(image_path) as base_image:
        base_image = base_image.convert("RGBA")

        # Resize QR code based on config
        qr_width = config['position']['width']
        qr_height = config['position']['height']
        qr_image = qr_image.resize((qr_width, qr_height))

        # Position QR code based on config
        qr_position = (
            config['position']['x'],
            config['position']['y']
        )

        base_image.paste(qr_image, qr_position, qr_image if qr_image.mode == 'RGBA' else None)
        base_image.save(output_path)

def process_images(excel_file, image_folder, output_folder, config):
    """Read Excel file, generate QR codes, and overlay them on images."""
    # Load data from Excel
    df = pd.read_excel(excel_file)

    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    for _, row in df.iterrows():
        filename = row['filename']
        qrstring = row['qrstring']

        image_path = os.path.join(image_folder, filename)
        output_path = os.path.join(output_folder, filename)

        if os.path.exists(image_path):
            print(f"Processing {filename}...")
            qr_image = generate_qr_code(qrstring)
            overlay_qr_on_image(image_path, qr_image, output_path, config)
        else:
            print(f"Image {filename} not found in {image_folder}. Skipping.")

#############################################################################
#zip
# Fungsi untuk melakukan batch zip pada file PNG
# Fungsi untuk melakukan batch zip pada file PNG
def batch_zip_files():
    folder_path = "qrModifiedOutput"
    final_folder = "final"
    
    # Cek apakah folder ada
    if not os.path.exists(folder_path):
        print(f"Folder '{folder_path}' tidak ditemukan!")
        return
    
    # Cek apakah folder final ada, jika tidak buat folder tersebut
    if not os.path.exists(final_folder):
        os.makedirs(final_folder)

    # Dapatkan daftar file PNG di dalam folder
    png_files = [f for f in os.listdir(folder_path) if f.endswith('.png')]

    if not png_files:
        print("Tidak ada file PNG di folder tersebut.")
        return

    # Proses batch zip untuk setiap file PNG
    for png_file in png_files:
        # Tentukan path file PNG dan path file ZIP
        png_file_path = os.path.join(folder_path, png_file)
        zip_file_name = f"{os.path.splitext(png_file)[0]}.zip"
        zip_file_path = os.path.join(folder_path, zip_file_name)
        
        # Buat file ZIP dan masukkan file PNG
        with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(png_file_path, os.path.basename(png_file))

        print(f"{png_file} telah di-zip menjadi {zip_file_path}")

        # Pindahkan file ZIP ke folder final
        shutil.move(zip_file_path, os.path.join(final_folder, zip_file_name))
        print(f"{zip_file_name} dipindahkan ke folder 'final'.")

    # Setelah semua file ZIP dipindahkan, lakukan zip seluruh file ZIP di folder final
    zip_final_file(final_folder)


# Fungsi untuk meng-zip seluruh file di dalam folder final
def zip_final_file(final_folder):
    final_zip_path = "final_output.zip"

    # Dapatkan daftar semua file ZIP di dalam folder final
    zip_files = [f for f in os.listdir(final_folder) if f.endswith('.zip')]

    if not zip_files:
        print("Tidak ada file ZIP di folder 'final' untuk di-zip.")
        return

    # Buat file ZIP final dan masukkan seluruh file ZIP
    with zipfile.ZipFile(final_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for zip_file in zip_files:
            zipf.write(os.path.join(final_folder, zip_file), zip_file)

    print(f"Seluruh file ZIP telah digabungkan menjadi {final_zip_path}")
#############################################################################
#Readme
def show_about():
    """Display information about the application."""
    about_text = """
    QR Code Attachment Tool
    ------------------------
    This application attaches QR codes to images based on data provided in an Excel file.

    Usage:
    1. Letakkan semua file gambar dalam format zip dalam folder zip, pilih menu 1, aplikasi akan melakukan unzip file dan membuat file excel dengan nama listQr.xlsx yang berisi field filename,qrstring dan tarif
       Tarif di di isi dengan cara maping conten string code dengan excel dari list Aino , contoh KEMENHUB SBY KHUSUS tarif 2000
    2. hapus QR existing dengan cara overlay gambar QR dengan kotak putih
    3. Modify string QR pada Excel dengan menambahkan tag 54 dan set amount sesuai maping
    4. Tempel QRcode yang baru pada template image qr yang lama (qr kosong), selain gambar qr conten lain tidak diubah 
    5. Lakukan kembali proses zip dan ambil hasilnya di folder final_output.zip

    Developed by: masCha
    """
    print(about_text)


#############################################################################

def parse_tlv(data):
    parsed_data = []
    index = 0
    while index < len(data):
        tag = data[index:index+2]  # Tag terdiri dari 2 digit
        length = int(data[index+2:index+4])  # Panjang terdiri dari 2 digit
        value = data[index+4:index+4+length]  # Ambil nilai berdasarkan panjang
        parsed_data.append({"tag": tag, "length": length, "value": value})
        index += 4 + length
    return parsed_data

# Membaca file Excel dan mengambil kolom "qr string"
def read_excel_file(file_path):
    df = pd.read_excel(file_path)
    if "qrstring" not in df.columns:
        raise ValueError("Kolom 'qrstring' tidak ditemukan dalam file Excel.")
    return df["qrstring"].dropna().tolist()

def parse_tlv(data):
    parsed_data = []
    index = 0
    while index < len(data):
        tag = data[index:index+2]  # Tag terdiri dari 2 digit
        length = int(data[index+2:index+4])  # Panjang terdiri dari 2 digit
        value = data[index+4:index+4+length]  # Ambil nilai berdasarkan panjang
        parsed_data.append({"tag": tag, "length": length, "value": value})
        index += 4 + length
    return parsed_data

# Membaca file Excel dan mengambil kolom "qr string"
def read_excel_file(file_path):
    df = pd.read_excel(file_path)
    if "qrstring" not in df.columns:
        raise ValueError("Kolom 'qrstring' tidak ditemukan dalam file Excel.")
    return df["qrstring"].dropna().tolist()

# Membaca konfigurasi dari file config.txt
def read_config_file(file_path):
    modifications = []
    with open(file_path, "r") as file:
        for line in file:
            parts = line.strip().split("|")
            if parts[0] == "+":
                modifications.append({"action": "+", "tag": parts[1], "length": parts[2], "value": parts[3]})
            elif parts[0] == "-":
                modifications.append({"action": "-", "tag": parts[1]})
    return modifications

# Memodifikasi QR string berdasarkan konfigurasi
def modify_qr_string(qr_string, modifications):
    parsed = parse_tlv(qr_string)
    parsed_dict = {item["tag"]: item for item in parsed}
    
    for mod in modifications:
        if mod["action"] == "-":
            parsed_dict.pop(mod["tag"], None)
        elif mod["action"] == "+":
            parsed_dict[mod["tag"]] = {"tag": mod["tag"], "length": int(mod["length"]), "value": mod["value"]}
    
    sorted_tags = sorted(parsed_dict.keys())
    modified_qr = "".join(f"{parsed_dict[tag]['tag']}{parsed_dict[tag]['length']:02}{parsed_dict[tag]['value']}" for tag in sorted_tags)
    return modified_qr

########################

def menu_utama():
    while True:  # Looping agar menu terus muncul setelah pilihan dieksekusi
        print("\nMenu Utama:")
        print("1. Unzip File dan Simpan String QR ke Excel")
        print("2. Hapus QR")
        print("3. Modify QR")
        print("4. Attach QR to Image")
        print("5. Zip")
        print("6. Parsing")
        print("7. Modifikasi QR String")
        print("8. Readme")
        print("9. Exit")

        # Ambil input dari pengguna
        pilihan = input("Pilih opsi (1-6): ")

        # Panggil fungsi sesuai dengan pilihan pengguna
        if pilihan == '1':
            #process_all_zip_files()
            print("Unzip File")
            folder_path = input("Masukkan path folder: ").strip()
            batch_unzip(folder_path)
        elif pilihan == '2':
            print("Hapus QR.")
            excel_path = "listQr.xlsx"  # Path to the Excel file
            folder_path = "unzipped_files"  # Folder containing base images
            overlay_image_path = "overlay.png"  # Path to the overlay image
            output_folder = "qrModified"  # Folder to save output images
            config_path = "config/config.json"  # Path to the configuration file

            #process_images(excel_path, folder_path, overlay_image_path, output_folder, config_path)
            #process_images(excel_file, image_folder, output_folder, config)
            #process_images(excel_path, folder_path, output_folder, config_path)
            process_images_hapusimages(excel_path, folder_path, overlay_image_path, output_folder, config_path)

            
        elif pilihan == '3':
            print("Modify QR")
            #file_path = input("Masukkan path file Excel: ")
            file_path ="listQr.xlsx"
            sheet_name ="Sheet1"
            #sheet_name = input("Masukkan nama sheet (kosongkan untuk default): ")
            
            try:
                if sheet_name:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                else:
                    df = pd.read_excel(file_path)
                
                # Asumsi kolom pertama berisi data untuk dihitung CRC
                # data_column = input("Masukkan nama kolom yang berisi data: ")
                data_column = "qrstring"
                tarif_column = "tarif"
                #tarif_column = input("Masukkan nama kolom yang berisi tarif: ")
                if data_column not in df.columns or tarif_column not in df.columns:
                    print(f"Kolom '{data_column}' atau '{tarif_column}' tidak ditemukan dalam file Excel.")
                else:
                    # Edit data pada setiap baris dan tambahkan CRC
                    edit_data_after_148th_char_tarif_and_crc(df, data_column, tarif_column)
                    
                    # Simpan hasilnya
                    output_file = "output_crc.xlsx"
                    df.to_excel(output_file, index=False)
                    print(f"Hasil CRC telah disimpan ke {output_file}")
            except Exception as e:
                print(f"Terjadi kesalahan: {e}")
        elif pilihan == '4':
            print("Attach QR to ASPI Format.")
            # Static paths
            excel_file = "output_crc.xlsx"
            image_folder = "qrModified"
            output_folder = "qrModifiedOutput"

            # Load configuration
            config_file = "config/config.json"
            if not os.path.exists(config_file):
                print(f"Error: Configuration file {config_file} not found.")
                continue

            with open(config_file, 'r') as f:
                config = json.load(f)

            if not os.path.exists(excel_file):
                print("Error: Excel file not found.")
                continue

            if not os.path.exists(image_folder):
                print("Error: Image folder not found.")
                continue

            process_images(excel_file, image_folder, output_folder, config)
            print("Processing complete. Check the output folder for results.")
        elif pilihan == '5':
             batch_zip_files()
        elif pilihan == '6':
             file_path = "listQr.xlsx"
             qr_strings = read_excel_file(file_path)
            
             for qr in qr_strings:
                parsed_data = parse_tlv(qr)  # Menggunakan langsung string sebagai input
                
                print(f"QR String: {qr}")
                for item in parsed_data:
                    print(f"Tag: {item['tag']}, Length: {item['length']}, Value: {item['value']}")
                print("-")
        elif pilihan == '7':
             file_path = "listQr.xlsx"
             qr_strings = read_excel_file(file_path)
             config_path = "config.txt"
             modifications = read_config_file(config_path)
            
             for qr in qr_strings:
                modified_qr = modify_qr_string(qr, modifications)
                print(f"Original QR: {qr}")
                print(f"Modified QR: {modified_qr}")
                parsed_data = parse_tlv(modified_qr)
                for item in parsed_data:
                    print(f"Tag: {item['tag']}, Length: {item['length']}, Value: {item['value']}")
                print("-")
                print("-")
        elif pilihan == '8':
             show_about()
        elif pilihan == '9':
            print("Keluar dari program.")
            break  # Keluar dari loop, program selesai

# Panggil menu utama
if __name__ == "__main__":
    menu_utama()