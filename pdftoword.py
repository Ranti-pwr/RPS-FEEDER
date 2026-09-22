import os
from pdf2docx import Converter
# folder_path = 'C:\Users\Sofiandi\Downloads\RPS_FOLDER'
folder_path = r'C:\Users\Sofiandi\Downloads\RPS_FOLDER'
 

if not os.path.exists(folder_path):
    print(f"Error: Folder '{folder_path}' tidak ditemukan.")
else:

    files = os.listdir(folder_path)

    pdf_files = [f for f in files if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("Tidak ada file PDF ditemukan di folder tersebut.")
    else:
        print(f"Menemukan {len(pdf_files)} file PDF. Memulai konversi...\n")
        
  
        for file_name in pdf_files:
            pdf_path = os.path.join(folder_path, file_name)
            
            name_without_ext = os.path.splitext(file_name)[0]
            docx_name = f"{name_without_ext}.docx"
            docx_path = os.path.join(folder_path, docx_name)
            
            print(f"Mengonversi: '{file_name}' -> '{docx_name}'...")
            
            try:
                cv = Converter(pdf_path)
                cv.convert(docx_path)
                cv.close()
                print(f"Sukses mengonversi '{file_name}'!\n")
            except Exception as e:
                print(f"Gagal mengonversi '{file_name}'. Error: {e}\n")
                
        print("Semua proses konversi selesai!")
