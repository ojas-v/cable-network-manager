import sqlite3
import pandas as pd
import os

# --- CONFIGURATION ---
EXCEL_FILE = "Sample_Customer_List.xlsx" 
DB_FILE = "cable_manager.db" 

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
    # Create basic structure if DB doesn't exist
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            can TEXT, name TEXT, address TEXT, contact_no TEXT, stb_no TEXT, 
            recovery_date TEXT, area TEXT, smart_card_no TEXT, wifi_router_id TEXT, 
            net_acc_no TEXT, install_date TEXT, monthly_rental TEXT, 
            total_connections TEXT, status TEXT DEFAULT 'Active',
            deposits TEXT, wifi_payment_details TEXT, last_payment_date TEXT, paid_amount TEXT
        )
    ''')
    conn.commit()
    conn.close()

def check_and_migrate_schema():
    """ FIX: Adds missing columns to prevent 'no column named' errors """
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    # List of new columns to check and add
    columns_to_check = {
        "stb_type": "TEXT DEFAULT 'SD'",
        "outstanding_amount": "TEXT DEFAULT '0'"
    }

    print("Checking database schema...")
    for col, dtype in columns_to_check.items():
        try:
            c.execute(f"SELECT {col} FROM customers LIMIT 1")
        except sqlite3.OperationalError:
            print(f"  -> Adding missing column: {col}")
            try:
                c.execute(f"ALTER TABLE customers ADD COLUMN {col} {dtype}")
            except Exception as e:
                print(f"     Error adding {col}: {e}")
    
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
    print("Starting import...")
    
    for index, row in df.iterrows():
        can = clean_can(row.get('CAN'))
        name = clean_text(row.get('Customer Name'))
        address = clean_text(row.get('Address'))
        contact = clean_can(row.get('Contact')) 
        stb = clean_text(row.get('STB No'))
        rec_date = clean_text(row.get('Payment Date')) 
        rental = clean_text(row.get('Paid'))
        
        # Now safe to run, even if DB was old
        cursor.execute('''
            INSERT INTO customers (
                can, name, address, contact_no, stb_no, recovery_date, monthly_rental, area, status, stb_type, outstanding_amount
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (can, name, address, contact, stb, rec_date, rental, "", "Active", "SD", "0"))
        
        count += 1

    conn.commit()
    conn.close()
    print(f"Successfully imported {count} customers into {DB_FILE}.")

if __name__ == "__main__":
    init_db()
    check_and_migrate_schema() # Run migration BEFORE import
    import_data()