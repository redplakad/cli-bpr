import os
import sqlite3
import pyfiglet
from datetime import datetime
from colorama import init, Fore, Style
from termcolor import colored
from script.data_import import import_loan_data  # Import function from script/data_import.py
from script.data_export import export_nominatif  # Import function from script/data_export.py

# Clear the terminal
os.system('cls' if os.name == 'nt' else 'clear')

# Function to create connection to SQLite database
def create_connection(db_file):
    conn = sqlite3.connect(db_file)
    return conn

# Function to create loan_data table if not exists
def create_table(conn):
    cursor = conn.cursor()
    create_table_query = """
    CREATE TABLE IF NOT EXISTS loan_data (
        CAB TEXT,
        NOMOR_REKENING TEXT,
        NO_CIF TEXT,
        NAMA_NASABAH TEXT,
        ALAMAT TEXT,
        KODE_KOLEK TEXT,
        JML_HRI_PKK INTEGER,
        JML_HARI_BGA INTEGER,
        JML_HARI_TUNGGAKAN INTEGER,
        KD_PRD TEXT,
        KET_KD_PRD TEXT,
        NOMOR_PERJANJIAN TEXT,
        NO_AKSEP TEXT,
        TGL_PK TEXT,
        TGL_AWAL_FAS TEXT,
        TGL_AKHIR_FAS TEXT,
        TGL_AWAL_AKSEP TEXT,
        TGL_AKH_AKSEP TEXT,
        PLAFOND_AWAL REAL,
        BAKI_DEBET REAL,
        LONGGAR_TARIK REAL,
        BGA REAL,
        TUNGGAKAN_POKOK REAL,
        TUNGGAKAN_BUNGA REAL,
        BGA_JTH_TEMPO TEXT,
        SMP_TGL_CADANG TEXT,
        NILAI_CADANG REAL,
        ANGSURAN_TOTAL REAL,
        TGL_PROSES_DENDA TEXT,
        AKUM_DENDA_PKK REAL,
        AKUM_DENDA_BGA REAL,
        PRD_AMORT REAL,
        PRDK_AMORT REAL,
        FLAG TEXT,
        TGL_AMORT TEXT,
        NILAI_BIAYA_PROVISI REAL,
        AMORTISASI_PRD REAL,
        SISA_AMORT_PROV REAL,
        TAGIH_BIAYA_PROV REAL,
        NILAI_BIAYA_ADM REAL,
        AMORT_ADM_PRD REAL,
        SISA_AMORT_ADM REAL,
        BYA_ASURANSI REAL,
        BYA_NOTARIS REAL,
        PKK_JATEM TEXT,
        BGA_JATEM TEXT,
        REK_BYR_PKK_BGA TEXT,
        SLD_REK_DB REAL,
        KD_INSTANSI TEXT,
        NM_INSTANSI TEXT,
        REK_BENDAHARA TEXT,
        SFT_KRD TEXT,
        GOL_KRD TEXT,
        JNS_KRD TEXT,
        SKTR_EKNM TEXT,
        ORNTS TEXT,
        NO_HP TEXT,
        POKOK_PINJAMAN REAL,
        TITIPAN_EFEKTIF REAL,
        JANGKA_WAKTU INTEGER,
        REK_PENCAIRAN TEXT,
        NO_REKENING_LAMA TEXT,
        CIF_LAMA TEXT,
        KODE_GROUP TEXT,
        KET_GROUP TEXT,
        TGL_LAHIR TEXT,
        NIK TEXT,
        NIP TEXT,
        NILAI_BYA_TRANS REAL,
        AMORT_TRANS_PRD REAL,
        SISA_AMORT_TRANS REAL,
        AO TEXT,
        CAB_REK TEXT,
        KELURAHAN TEXT,
        KECAMATAN TEXT,
        CADANGAN_PPAP REAL,
        TEMPAT_BEKERJA TEXT,
        TGL_AKHIR_BAYAR TEXT,
        PIHAK_TERKAIT TEXT,
        JENIS_JAMINAN TEXT,
        NILAI_LEGALITAS REAL,
        RESTRUKTUR_KE TEXT,
        TGL_VALID_KOLEK TEXT,
        TGL_MACET TEXT
    );
    """
    cursor.execute(create_table_query)
    conn.commit()

# Function to display main menu and handle user input
def main_menu(conn):
    init(autoreset=True)
    current_datetime = datetime.now()
    
    while True:
        # ASCII art using pyfiglet
        ascii_art = pyfiglet.figlet_format("BPR SERANG")
        print(colored(ascii_art, 'cyan'))
        print(Fore.CYAN + f"Tanggal : {current_datetime.strftime('%d/%m/%Y')}                            Jam: {current_datetime.strftime('%H:%M:%S')}")
        print(Fore.YELLOW + "##############################################################")
        print(Fore.YELLOW + "#                                                            #")
        print(Fore.YELLOW + "#    ini adalah aplikasi untuk convert csv loan ke excel     #")
        print(Fore.YELLOW + "#                                                            #")
        print(Fore.YELLOW + "##############################################################")
        
        print(Fore.YELLOW + "\n====================== MENU UTAMA ============================")
        print(Fore.YELLOW + "||                                                          ||")
        print(Fore.YELLOW + "||      1. IMPORT DATA LOAN                                 ||")
        print(Fore.YELLOW + "||      2. EXPORT DATA NOMINATIF                            ||")
        print(Fore.YELLOW + "||      3. EXPORT NOMINATIF DENGAN FILTER                   ||")
        print(Fore.YELLOW + "||      4. KELUAR                                           ||")
        print(Fore.YELLOW + "||                                                          ||")
        print(Fore.YELLOW + "===================== version.0.1 ============================")
        choice = input(Fore.CYAN + "\nMasukkan pilihan Anda (1/2/3/4): ")
        
        if choice == '1':
            import_loan_data(conn)  # Call function from import.py
        elif choice == '2':
            export_nominatif(conn)
        elif choice == '3':
            nominatif_per_kolektibilitas(conn)
        elif choice == '4':
            print(Fore.MAGENTA + "Terima kasih telah menggunakan aplikasi ini.")
            break
        else:
            print(Fore.RED + "Pilihan tidak valid. Silakan pilih kembali.")

if __name__ == "__main__":
    # SQLite database file
    db_file = 'database.db'
    
    # Create connection to SQLite database
    conn = create_connection(db_file)
    
    # Create loan_data table if not exists
    create_table(conn)
    
    # Run the main menu
    main_menu(conn)
    
    # Close database connection
    conn.close()
