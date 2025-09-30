# Test script untuk memverifikasi PySerial installation

def test_pyserial_installation():
    """
    Test apakah PySerial sudah terinstall dengan benar
    """
    try:
        import serial
        print("✓ Module 'serial' berhasil diimport")
        
        # Test apakah Serial class ada
        if hasattr(serial, 'Serial'):
            print("✓ Class 'Serial' ditemukan")
            print(f"✓ PySerial version: {serial.__version__}")
        else:
            print("✗ Class 'Serial' tidak ditemukan")
            print("Ini menunjukkan bahwa module yang diimport bukan PySerial")
            return False
            
        # Test apakah list_ports tersedia
        try:
            import serial.tools.list_ports
            print("✓ Module 'serial.tools.list_ports' tersedia")
        except ImportError as e:
            print(f"✗ Error importing list_ports: {e}")
            return False
            
        # Coba buat instance Serial (tanpa membuka port)
        try:
            ser = serial.Serial()
            print("✓ Berhasil membuat instance Serial")
            ser.close()
        except Exception as e:
            print(f"✗ Error membuat instance Serial: {e}")
            return False
            
        print("\n🎉 PySerial terinstall dengan benar!")
        return True
        
    except ImportError as e:
        print(f"✗ Error importing serial: {e}")
        print("PySerial belum terinstall atau tidak ditemukan")
        return False
    except Exception as e:
        print(f"✗ Unexpected error: {e}")
        return False

def check_conflicting_files():
    """
    Check apakah ada file yang mungkin konflik dengan PySerial
    """
    import os
    import sys
    
    print("\nMemeriksa kemungkinan konflik file:")
    print("-" * 40)
    
    # Check file serial.py di direktori saat ini
    if os.path.exists('serial.py'):
        print("⚠️  DITEMUKAN: file 'serial.py' di direktori saat ini")
        print("   File ini mungkin menyebabkan konflik dengan PySerial")
        print("   Pertimbangkan untuk mengubah nama file ini")
    else:
        print("✓ Tidak ada file 'serial.py' di direktori saat ini")
    
    # Check folder serial di direktori saat ini
    if os.path.exists('serial') and os.path.isdir('serial'):
        print("⚠️  DITEMUKAN: folder 'serial' di direktori saat ini")
        print("   Folder ini mungkin menyebabkan konflik dengan PySerial")
    else:
        print("✓ Tidak ada folder 'serial' di direktori saat ini")
    
    # Show Python path
    print(f"\nPython path saat ini:")
    for i, path in enumerate(sys.path, 1):
        print(f"  {i}. {path}")

def installation_guide():
    """
    Panduan instalasi PySerial
    """
    print("\n" + "="*50)
    print("PANDUAN INSTALASI PYSERIAL")
    print("="*50)
    
    print("\n1. Install menggunakan pip:")
    print("   pip install pyserial")
    
    print("\n2. Atau install menggunakan conda (jika menggunakan Anaconda):")
    print("   conda install pyserial")
    
    print("\n3. Verifikasi instalasi:")
    print("   python -c \"import serial; print(serial.__version__)\"")
    
    print("\n4. Jika masih bermasalah, coba:")
    print("   pip uninstall serial")
    print("   pip uninstall pyserial") 
    print("   pip install pyserial")
    
    print("\n5. Untuk sistem Linux/Mac, mungkin perlu:")
    print("   sudo pip install pyserial")
    print("   atau")
    print("   python3 -m pip install pyserial")

if __name__ == "__main__":
    print("DIAGNOSIS MASALAH PYSERIAL")
    print("=" * 50)
    
    # Test instalasi
    success = test_pyserial_installation()
    
    # Check konflik file
    check_conflicting_files()
    
    # Jika gagal, tampilkan panduan
    if not success:
        installation_guide()
    
    print("\n" + "="*50)
    print("DIAGNOSIS SELESAI")
    print("="*50)