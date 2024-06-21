import os
import sqlite3
import pandas as pd
from colorama import init, Fore, Style

# Initialize colorama
init(autoreset=True)

# Function to import loan data from CSV into SQLite database
def import_loan_data(conn):
    try:
        # List all CSV files in the current directory
        csv_files = [file for file in os.listdir('.') if file.endswith('.csv')]
        
        if not csv_files:
            
            os.system('cls' if os.name == 'nt' else 'clear')
            print(Fore.RED + "Tidak ada file CSV yang ditemukan dalam direktori saat ini.")
            return
        
        # Print list of CSV files
        print(Fore.YELLOW + "Pilih file CSV yang ingin diimpor:")
        for idx, file in enumerate(csv_files, start=1):
            print(Fore.YELLOW + f"{idx}. {file}")
        
        # Prompt user to choose a file
        choice_idx = input(Fore.CYAN + "Masukkan nomor file (1, 2, ...): ")
        
        try:
            choice_idx = int(choice_idx)
            if choice_idx < 1 or choice_idx > len(csv_files):
                raise ValueError("Nomor file tidak valid.")
        except ValueError:
            print(Fore.RED + "Nomor file tidak valid.")
            return
        
        # Select the chosen file
        csv_file = csv_files[choice_idx - 1]
        
        # Read the selected CSV file
        df = pd.read_csv(csv_file, sep='|', on_bad_lines='skip')
        
    except pd.errors.ParserError as e:
        print(Fore.RED + f"Error parsing CSV file: {e}")
        return
    
    cursor = conn.cursor()
    try:
        for index, row in df.iterrows():
            columns = ', '.join([f'"{col}"' for col in row.index])
            placeholders = ', '.join(['?'] * len(row))
            insert_query = f"INSERT INTO loan_data ({columns}) VALUES ({placeholders})"
            cursor.execute(insert_query, tuple(row))
        conn.commit()
        
        # Clear the terminal
        os.system('cls' if os.name == 'nt' else 'clear')
        # Print success message with box frame
        print(Fore.GREEN + "=" * 40)
        print(Fore.GREEN + "+   Data imported successfully!   +")
        print(Fore.GREEN + "=" * 40)
        
    except sqlite3.Error as e:
        os.system('cls' if os.name == 'nt' else 'clear')
        conn.rollback()
        print(Fore.RED + f"Error inserting data into database: {e}")
    finally:
        cursor.close()

# Example usage:
if __name__ == "__main__":
    # SQLite database file
    db_file = 'database.db'
    
    # Create connection to SQLite database
    conn = sqlite3.connect(db_file)
    
    # Run the import function
    import_loan_data(conn)
    
    # Close database connection
    conn.close()
