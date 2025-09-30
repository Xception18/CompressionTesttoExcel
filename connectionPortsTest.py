import serial
import serial.tools.list_ports
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os
import time

# Konfigurasi default untuk koneksi serial
BAUD_RATE = 9600
TIMEOUT = 1
DATA = 8
EXCEL_FILE = "pressure_data.xlsx"

def init_excel_file():
    """
    Inisialisasi file Excel dengan header jika belum ada
    """
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        if ws is not None:
            ws.title = "Pressure Data"
            ws.append(["No", "Timestamp", "Port", "Nilai Tekanan (KN)"])
        wb.save(EXCEL_FILE)
        print(f"File Excel baru dibuat: {EXCEL_FILE}")
    else:
        print(f"Menggunakan file Excel: {EXCEL_FILE}")

def save_to_excel(port, nilai_kn):
    """
    Menyimpan data tekanan ke file Excel
    
    Args:
        port (str): Nama port yang digunakan
        nilai_kn (str): Nilai tekanan yang dibaca
    """
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        
        if ws is None:
            print("✗ Error: Worksheet tidak ditemukan")
            return False
        
        # Hitung nomor urut (jumlah baris - 1 untuk header)
        row_number = ws.max_row
        
        # Tambahkan data baru
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.append([row_number, timestamp, port, nilai_kn])
        
        wb.save(EXCEL_FILE)
        print(f"✓ Data tersimpan ke Excel (Baris {row_number + 1}): {nilai_kn} KN")
        return True
        
    except Exception as e:
        print(f"✗ Error menyimpan ke Excel: {e}")
        return False

def list_com_ports():
    """
    Menampilkan dan mengembalikan daftar COM ports yang tersedia
    """
    ports = serial.tools.list_ports.comports()
    
    if not ports:
        print("No COM ports found")
        return []
    
    print("Available COM ports:")
    available_ports = []
    
    for port in ports:
        print(f"Port: {port.device}")
        print(f"Description: {port.description}")
        print(f"Hardware ID: {port.hwid}")
        print("-" * 40)
        available_ports.append(port.device)
    
    return available_ports

def koneksi(port, baud=BAUD_RATE, data=DATA):
    """
    Menguji koneksi ke port serial yang ditentukan
    
    Args:
        port (str): Nama port (misal: 'COM3' atau '/dev/ttyUSB0')
        baud (int): Baud rate (default: 9600)
        data (int): Data bits (default: 8)
    
    Returns:
        str: Status koneksi
    """
    try:
        ser = serial.Serial(port, baud, data, timeout=TIMEOUT)
        if ser.is_open:
            print(f"Berhasil terhubung ke {port}")
            ser.close()
            return "Tersambung !"
        else:
            return "Gagal membuka port"
    except Exception as e:
        pesanError = str(e)
        print(f"Error connecting to {port}: {pesanError}")
        return f"Error: {pesanError}"

def nilaiTekan(port, baud=BAUD_RATE, data=DATA):
    """
    Membaca nilai tekanan dari perangkat serial
    
    Args:
        port (str): Nama port (misal: 'COM3' atau '/dev/ttyUSB0')
        baud (int): Baud rate (default: 9600)
        data (int): Data bits (default: 8)
    
    Returns:
        str: Nilai tekanan yang dibaca dari perangkat, atau None jika gagal
    """
    try:
        ser = serial.Serial(port, baud, data, timeout=TIMEOUT)
        if ser.is_open:
            print(f"Membaca data dari {port}...")
            print(ser)
            nilaiKN = ""
            
            while True:
                if datanya := ser.readline():
                    data_str = datanya.decode("utf-8", errors="ignore").strip()
                    print(f"Data diterima: {data_str}")  # Debug print
                    
                    if "ovalue" in data_str.lower():
                        try:
                            nilaiKN = data_str.split()[1]
                            ser.close()
                            print(f"Nilai Tekan: {nilaiKN}")
                            return nilaiKN
                        except IndexError:
                            print("Format data tidak sesuai, mencoba lagi...")
                            continue
                else:
                    print("Tidak ada data yang diterima, timeout...")
                    break
            
            ser.close()
            return None
            
    except Exception as e:
        pesanError = str(e)
        print(f"Error reading pressure value: {pesanError}")
        return None

def continuous_reading_mode():
    """
    Mode pembacaan kontinyu dengan auto-save ke Excel
    """
    ports = list_com_ports()
    
    if not ports:
        return
    
    print(f"\nFound {len(ports)} COM ports")
    print("Select a port for continuous reading:")
    
    for i, port in enumerate(ports, 1):
        print(f"{i}. {port}")
    
    try:
        choice = int(input("Enter port number (or 0 to cancel): "))
        
        if choice == 0:
            print("Cancelled...")
            return
        
        if 1 <= choice <= len(ports):
            selected_port = ports[choice - 1]
            
            # Konfigurasi interval pembacaan
            try:
                interval = input("Enter reading interval in seconds (default 2): ").strip()
                interval = float(interval) if interval else 2.0
            except ValueError:
                print("Invalid interval, using default 2 seconds")
                interval = 2.0
            
            print(f"\n{'='*60}")
            print(f"Mode Pembacaan Kontinyu Aktif")
            print(f"Port: {selected_port}")
            print(f"Interval: {interval} detik")
            print(f"Data akan disimpan ke: {EXCEL_FILE}")
            print(f"Ketik 'exit' dan tekan Enter untuk berhenti")
            print(f"{'='*60}\n")
            
            # Inisialisasi file Excel
            init_excel_file()
            
            reading_count = 0
            
            try:
                while True:
                    # Cek input user (non-blocking)
                    import sys
                    import select
                    
                    # Untuk Windows
                    if sys.platform == 'win32':
                        import msvcrt
                        if msvcrt.kbhit():
                            user_input = input().strip().lower()
                            if user_input == 'exit':
                                print("\n✓ Pembacaan dihentikan oleh user")
                                print(f"Total pembacaan: {reading_count}")
                                break
                    else:
                        # Untuk Linux/Mac
                        i, o, e = select.select([sys.stdin], [], [], 0.1)
                        if i:
                            user_input = sys.stdin.readline().strip().lower()
                            if user_input == 'exit':
                                print("\n✓ Pembacaan dihentikan oleh user")
                                print(f"Total pembacaan: {reading_count}")
                                break
                    
                    # Baca nilai tekanan
                    reading_count += 1
                    print(f"\n[Pembacaan #{reading_count}] {datetime.now().strftime('%H:%M:%S')}")
                    
                    nilai = nilaiTekan(selected_port)
                    
                    if nilai:
                        save_to_excel(selected_port, nilai)
                    else:
                        print("✗ Gagal membaca nilai tekanan, mencoba lagi...")
                    
                    # Tunggu sebelum pembacaan berikutnya
                    print(f"Menunggu {interval} detik... (ketik 'exit' untuk berhenti)")
                    time.sleep(interval)
                    
            except KeyboardInterrupt:
                print("\n\n✓ Pembacaan dihentikan oleh user (Ctrl+C)")
                print(f"Total pembacaan: {reading_count}")
            except Exception as e:
                print(f"\nError: {e}")
                print(f"Total pembacaan: {reading_count}")
                
        else:
            print("Invalid selection!")
            
    except ValueError:
        print("Please enter a valid number!")
    except KeyboardInterrupt:
        print("\n\n✓ Pembacaan dihentikan oleh user (Ctrl+C)")

def test_all_ports():
    """
    Menguji koneksi ke semua port yang tersedia
    """
    print("Testing all available ports...\n")
    ports = list_com_ports()
    
    if not ports:
        return
    
    print("\nTesting connections:")
    print("=" * 50)
    
    for port in ports:
        print(f"Testing {port}...")
        status = koneksi(port)
        print(f"Status: {status}")
        print("-" * 30)

def connect_to_specific_port():
    """
    Memungkinkan pengguna memilih port tertentu untuk koneksi
    """
    ports = list_com_ports()
    
    if not ports:
        return
    
    print(f"\nFound {len(ports)} COM ports")
    print("Select a port to connect:")
    
    for i, port in enumerate(ports, 1):
        print(f"{i}. {port}")
    
    try:
        choice = int(input("Enter port number (or 0 to exit): "))
        
        if choice == 0:
            print("Exiting...")
            return
        
        if 1 <= choice <= len(ports):
            selected_port = ports[choice - 1]
            print(f"\nConnecting to {selected_port}...")
            status = koneksi(selected_port)
            print(f"Connection status: {status}")
        else:
            print("Invalid selection!")
            
    except ValueError:
        print("Please enter a valid number!")
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")

def read_pressure_from_port():
    """
    Memungkinkan pengguna memilih port untuk membaca nilai tekanan sekali
    """
    ports = list_com_ports()
    
    if not ports:
        return
    
    print(f"\nFound {len(ports)} COM ports")
    print("Select a port to read pressure value:")
    
    for i, port in enumerate(ports, 1):
        print(f"{i}. {port}")
    
    try:
        choice = int(input("Enter port number (or 0 to exit): "))
        
        if choice == 0:
            print("Exiting...")
            return
        
        if 1 <= choice <= len(ports):
            selected_port = ports[choice - 1]
            print(f"\nReading pressure from {selected_port}...")
            
            nilai = nilaiTekan(selected_port)
            
            if nilai:
                print(f"✓ Berhasil membaca nilai tekanan: {nilai}")
                
                # Tanya apakah ingin menyimpan ke Excel
                save_choice = input("Simpan ke Excel? (y/n): ").strip().lower()
                if save_choice == 'y':
                    init_excel_file()
                    save_to_excel(selected_port, nilai)
            else:
                print("✗ Gagal membaca nilai tekanan")
        else:
            print("Invalid selection!")
            
    except ValueError:
        print("Please enter a valid number!")
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")

# Main program
if __name__ == "__main__":
    print("=" * 60)
    print("Serial Port Manager dengan Excel Data Logging")
    print("=" * 60)
    
    while True:
        print("\n" + "=" * 60)
        print("Menu Utama:")
        print("=" * 60)
        print("1. List all COM ports")
        print("2. Test all ports")
        print("3. Connect to specific port")
        print("4. Read pressure value (single)")
        print("5. Continuous reading mode (AUTO-SAVE to Excel)")
        print("6. Exit")
        print("=" * 60)
        
        try:
            choice = input("Pilih opsi (1-6) atau ketik 'exit': ").strip().lower()
            
            if choice == 'exit' or choice == '6':
                print("\n" + "=" * 60)
                print("Terima kasih! Program selesai.")
                print("=" * 60)
                break
                
            elif choice == '1':
                print()
                ports = list_com_ports()
                print(f"\nFound {len(ports)} COM ports: {ports}")
                
            elif choice == '2':
                print()
                test_all_ports()
                
            elif choice == '3':
                connect_to_specific_port()
                
            elif choice == '4':
                read_pressure_from_port()
                
            elif choice == '5':
                continuous_reading_mode()
                
            else:
                print("Opsi tidak valid! Pilih 1-6 atau ketik 'exit'.")
                
        except KeyboardInterrupt:
            print("\n\n" + "=" * 60)
            print("Program dihentikan oleh user (Ctrl+C)")
            print("=" * 60)
            break
        except Exception as e:
            print(f"Terjadi error: {e}")