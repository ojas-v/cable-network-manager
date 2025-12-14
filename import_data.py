import sqlite3
import pandas as pd
import os

# --- CONFIGURATION ---
EXCEL_FILE = "Sample_Customer_List.xlsx" 
DB_FILE = "cable_manager.db" # Updated to match main.py

def clean_text(val):
    if pd.isna(val) or val == "" or str(val).lower() == "nan": return ""
    return str(val).strip()

def clean_can(val):
    if pd.isna(val) or val == "": return ""
    try: return str(int(float(val)))
    except: return str(val).strip()

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    # Simplified table creation for import test
    cursor.execute('''CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT, can TEXT, name TEXT, address TEXT, 
        contact_no TEXT, stb_no TEXT, recovery_date TEXT, area TEXT, 
        smart_card_no TEXT, wifi_router_id TEXT, net_acc_no TEXT, install_date TEXT, 
        monthly_rental TEXT, total_connections TEXT, status TEXT DEFAULT 'Active',
        deposits TEXT, wifi_payment_details TEXT, last_payment_date TEXT, paid_amount TEXT
    )''')
    conn.commit()
    conn.close()

def import_data():
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: {EXCEL_FILE} not found.")
        return
    df = pd.read_excel(EXCEL_FILE)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    count = 0
    for index, row in df.iterrows():
        can = clean_can(row.get('CAN'))
        name = clean_text(row.get('Customer Name'))
        address = clean_text(row.get('Address'))
        contact = clean_can(row.get('Contact')) 
        stb = clean_text(row.get('STB No'))
        rec_date = clean_text(row.get('Payment Date')) 
        rental = clean_text(row.get('Paid'))
        
        cursor.execute('''
            INSERT INTO customers (
                can, name, address, contact_no, stb_no, recovery_date, monthly_rental, area, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (can, name, address, contact, stb, rec_date, rental, "", "Active"))
        count += 1
    conn.commit()
    conn.close()
    print(f"Successfully imported {count} customers into {DB_FILE}.")

if __name__ == "__main__":
    init_db()
    import_data()