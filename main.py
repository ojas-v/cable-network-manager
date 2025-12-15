import customtkinter as ctk
from tkinter import messagebox, StringVar, filedialog, simpledialog
import sqlite3
import datetime
import shutil
import os
import webbrowser
import urllib.parse
import pandas as pd 

# --- CONFIGURATION ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

DB_FILE = "cable_manager.db" 
EXCEL_FILE = "Sample_Customer_List.xlsx" 
APP_ICON = "app_icon.ico"

# --- BUSINESS DETAILS ---
BUSINESS_NAME = "YOUR CABLE NETWORK NAME"
BUSINESS_ADDRESS = "123, Your Street Address, City, State - Zip"
SUPPORT_CONTACT = "9876543210" 
ADMIN_PASSWORD = "admin" 

class CableManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Cable Network Manager")
        self.geometry("1300x850")
        
        try:
            self.iconbitmap(APP_ICON)
        except:
            pass 

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- SYSTEM INITIALIZATION ---
        self.init_database()
        self.check_db_schema() 
        self.auto_import_data()

        # --- Variables ---
        self.current_customer_id = None
        
        self.var_can = StringVar()
        self.var_name = StringVar()
        self.var_address = StringVar()
        self.var_contact = StringVar()
        self.var_stb = StringVar()
        self.var_stb_type = StringVar(value="SD") 
        self.var_smartcard = StringVar()
        self.var_router = StringVar()
        self.var_net_acc = StringVar()
        self.var_rental = StringVar()
        self.var_connections = StringVar()
        self.var_recovery = StringVar()
        self.var_area = StringVar(value="Unassigned")
        self.var_status = StringVar(value="Active") 
        self.var_install_date = StringVar()
        self.var_outstanding = StringVar(value="0") 
        
        self.var_pay_search = StringVar()
        self.var_pay_amount = StringVar()
        self.var_pay_date = StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.var_pay_history_date = StringVar(value="Select Date")
        
        self.var_start_date = StringVar()
        self.var_end_date = StringVar()
        self.var_invoice_footer = StringVar(value="*Terms & Conditions Apply. Final Decision of the Proprietor.")
        self.var_inv_item = StringVar()
        self.var_inv_qty = StringVar()
        self.var_complaint_issue = StringVar()
        self.var_new_area = StringVar()

        self.setup_sidebar()
        self.setup_main_area()
        self.show_dashboard() 

    def get_db_connection(self):
        return sqlite3.connect(DB_FILE)

    def init_database(self):
        conn = self.get_db_connection()
        c = conn.cursor()
        
        c.execute('''
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                can TEXT, name TEXT, address TEXT, contact_no TEXT, stb_no TEXT, 
                stb_type TEXT, recovery_date TEXT, area TEXT, smart_card_no TEXT, 
                wifi_router_id TEXT, net_acc_no TEXT, install_date TEXT, monthly_rental TEXT, 
                total_connections TEXT, status TEXT DEFAULT 'Active',
                deposits TEXT, wifi_payment_details TEXT, last_payment_date TEXT, 
                paid_amount TEXT, outstanding_amount TEXT DEFAULT '0'
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS inventory (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_name TEXT UNIQUE,
                quantity INTEGER DEFAULT 0
            )
        ''')
        c.execute("SELECT count(*) FROM inventory")
        if c.fetchone()[0] == 0:
            defaults = ["Set Top Box", "Adapter", "Remote", "HDMI Cord", "AV Cord", "Wire (Bundle)", "WiFi Router"]
            for item in defaults:
                c.execute("INSERT OR IGNORE INTO inventory (item_name, quantity) VALUES (?, 0)", (item,))

        c.execute('''
            CREATE TABLE IF NOT EXISTS complaints (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_id INTEGER,
                customer_name TEXT,
                issue TEXT,
                status TEXT DEFAULT 'Open',
                date_logged TEXT,
                date_resolved TEXT
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS areas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                area_name TEXT UNIQUE
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS payment_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_id INTEGER,
                can TEXT,
                amount_paid TEXT,
                date_paid TEXT,
                remarks TEXT
            )
        ''')
        
        conn.commit()
        conn.close()

    def check_db_schema(self):
        conn = self.get_db_connection()
        c = conn.cursor()
        try: c.execute("SELECT stb_type FROM customers LIMIT 1")
        except sqlite3.OperationalError: 
            try: c.execute("ALTER TABLE customers ADD COLUMN stb_type TEXT DEFAULT 'SD'")
            except: pass
        try: c.execute("SELECT outstanding_amount FROM customers LIMIT 1")
        except sqlite3.OperationalError:
            try: c.execute("ALTER TABLE customers ADD COLUMN outstanding_amount TEXT DEFAULT '0'")
            except: pass
        try: c.execute("SELECT date_resolved FROM complaints LIMIT 1")
        except sqlite3.OperationalError:
            try: c.execute("ALTER TABLE complaints ADD COLUMN date_resolved TEXT")
            except: pass
        conn.commit()
        conn.close()

    def auto_import_data(self):
        if not os.path.exists(EXCEL_FILE): return 
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("SELECT count(*) FROM customers")
        if c.fetchone()[0] == 0:
            try:
                df = pd.read_excel(EXCEL_FILE)
                for index, row in df.iterrows():
                    can = str(row.get('CAN', '')).strip()
                    if can.lower() == 'nan': can = ""
                    name = str(row.get('Customer Name', '')).strip()
                    
                    c.execute('''
                        INSERT INTO customers (
                            can, name, address, contact_no, stb_no, recovery_date, monthly_rental, area, status, stb_type, outstanding_amount
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (can, name, str(row.get('Address','')), str(row.get('Contact','')), str(row.get('STB No','')), str(row.get('Payment Date','')), str(row.get('Paid','')), "", "Active", "SD", "0"))
                conn.commit()
            except Exception as e:
                print(f"Auto-import failed: {e}")
        conn.close()

    # --- UI SETUP ---
    def setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text=BUSINESS_NAME, font=ctk.CTkFont(size=20, weight="bold"), wraplength=200).grid(row=0, column=0, padx=20, pady=20)

        self.create_nav_btn("Dashboard", self.show_dashboard, 1)
        self.create_nav_btn("Customer Manager", self.show_customer_manager, 2)
        self.create_nav_btn("Customer Payments", self.show_payment_tab, 3) 
        self.create_nav_btn("Inventory", self.show_inventory, 4)
        self.create_nav_btn("Complaints", self.show_complaints, 5)
        self.create_nav_btn("Reports", self.show_reports, 6)
        self.create_nav_btn("Settings", self.show_settings, 7)
        
        ctk.CTkButton(self.sidebar, text="Exit", command=self.destroy, fg_color="#d9534f", hover_color="#c9302c").grid(row=9, column=0, padx=20, pady=40, sticky="s")

    def create_nav_btn(self, text, command, row):
        ctk.CTkButton(self.sidebar, text=text, command=command, fg_color="transparent", text_color=("gray10", "#DCE4EE"), hover_color=("gray70", "gray30"), anchor="w", height=40).grid(row=row, column=0, padx=10, pady=5, sticky="ew")

    def setup_main_area(self):
        self.main_view = ctk.CTkFrame(self, fg_color="transparent")
        self.main_view.grid(row=0, column=1, sticky="nsew")
        self.main_view.grid_columnconfigure(0, weight=1)
        self.main_view.grid_rowconfigure(1, weight=1)

        self.top_bar = ctk.CTkFrame(self.main_view, height=70, corner_radius=0)
        self.top_bar.grid(row=0, column=0, sticky="ew")
        
        self.search_entry = ctk.CTkEntry(self.top_bar, placeholder_text="Global Search: Name, CAN, STB...", width=500, font=("Arial", 16), height=40)
        self.search_entry.pack(side="left", padx=20, pady=15)
        self.search_entry.bind('<Return>', self.perform_search) 
        ctk.CTkButton(self.top_bar, text="Search", command=self.perform_search, width=120, height=40, font=("Arial", 14, "bold")).pack(side="left", padx=5)

    # --- DASHBOARD ---
    def show_dashboard(self):
        self.clear_content_frame()
        content = ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent")
        content.pack(fill="both", expand=True)

        ctk.CTkLabel(content, text="Dashboard Overview", font=("Arial", 24, "bold")).pack(anchor="w", padx=20, pady=10)
        
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM customers WHERE status='Active'")
        active_subs = c.fetchone()[0]
        c.execute("SELECT COUNT(DISTINCT area_name) FROM areas")
        coverage = c.fetchone()[0]
        conn.close()

        stats = ctk.CTkFrame(content, fg_color="transparent")
        stats.pack(fill="x", padx=10)
        self.card(stats, "Active Subscribers", str(active_subs), "#007bff").pack(side="left", fill="x", expand=True, padx=5)
        self.card(stats, "Network Coverage", str(coverage), "#6610f2").pack(side="left", fill="x", expand=True, padx=5)
        self.card(stats, "System Date", str(datetime.date.today()), "#28a745").pack(side="left", fill="x", expand=True, padx=5)

        area_frame = ctk.CTkFrame(content)
        area_frame.pack(fill="x", padx=15, pady=20)
        ctk.CTkLabel(area_frame, text="Manage Service Areas", font=("Arial", 16, "bold")).pack(anchor="w", padx=10, pady=10)
        ctk.CTkEntry(area_frame, textvariable=self.var_new_area, placeholder_text="Area Name").pack(side="left", padx=10, pady=10)
        ctk.CTkButton(area_frame, text="Add Area", command=self.add_area, fg_color="green").pack(side="left", padx=10)
        ctk.CTkButton(area_frame, text="Delete Selected Area", command=self.delete_area, fg_color="red").pack(side="left", padx=10)

        actions = ctk.CTkFrame(content)
        actions.pack(fill="x", padx=15, pady=10)
        ctk.CTkLabel(actions, text="Quick Actions for Selected Customer", font=("Arial", 16, "bold")).pack(anchor="w", padx=10, pady=10)
        
        if self.var_name.get():
            ctk.CTkLabel(actions, text=f"Selected: {self.var_name.get()} ({self.var_can.get()})").pack(pady=5)
            ctk.CTkButton(actions, text="Open WhatsApp Reminder", command=self.open_whatsapp_web, fg_color="#25D366").pack(side="left", padx=10, pady=10, expand=True)
            ctk.CTkButton(actions, text="Generate Receipt (PDF)", command=self.generate_receipt_pdf, fg_color="#17a2b8").pack(side="left", padx=10, pady=10, expand=True)
        else:
            ctk.CTkLabel(actions, text="No Customer Selected. Search above.").pack(pady=10)

    # --- CUSTOMER PAYMENT TAB ---
    def show_payment_tab(self):
        self.clear_content_frame()
        content = ctk.CTkScrollableFrame(self.content_frame)
        content.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(content, text="Customer Payments Portal", font=("Arial", 22, "bold")).pack(anchor="w", pady=(0, 20))

        search_frame = ctk.CTkFrame(content)
        search_frame.pack(fill="x", pady=10)
        ctk.CTkEntry(search_frame, textvariable=self.var_pay_search, placeholder_text="Type Name to Search...", width=300).pack(side="left", padx=10, pady=10)
        ctk.CTkButton(search_frame, text="Find Customer", command=self.search_for_payment).pack(side="left", padx=10)

        self.pay_results_frame = ctk.CTkScrollableFrame(content, height=100)
        self.pay_results_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(self.pay_results_frame, text="Search results will appear here...").pack()

        if self.current_customer_id:
            main_frame = ctk.CTkFrame(content)
            main_frame.pack(fill="both", expand=True, pady=10)
            
            left = ctk.CTkFrame(main_frame, fg_color="transparent")
            left.pack(side="left", fill="both", expand=True, padx=10, pady=10)
            
            ctk.CTkLabel(left, text=f"Customer: {self.var_name.get()}", font=("Arial", 18, "bold")).pack(anchor="w")
            ctk.CTkLabel(left, text=f"Voucher/CAN No: {self.var_can.get()}", font=("Arial", 14)).pack(anchor="w")
            ctk.CTkLabel(left, text=f"Contact: {self.var_contact.get()}", font=("Arial", 14)).pack(anchor="w")
            ctk.CTkLabel(left, text=f"Address: {self.var_address.get()[:40]}...", font=("Arial", 14)).pack(anchor="w")
            ctk.CTkLabel(left, text=f"STB No: {self.var_stb.get()}", font=("Arial", 14)).pack(anchor="w")
            ctk.CTkLabel(left, text=f"Rental: ₹{self.var_rental.get()}", font=("Arial", 14)).pack(anchor="w")
            ctk.CTkLabel(left, text=f"Total Connections: {self.var_connections.get()}", font=("Arial", 14)).pack(anchor="w")
            
            ctk.CTkLabel(left, text="---", font=("Arial", 10)).pack(anchor="w", pady=5)
            ctk.CTkLabel(left, text=f"Recovery Date: {self.var_recovery.get()}", text_color="orange", font=("Arial", 14, "bold")).pack(anchor="w")
            ctk.CTkLabel(left, text=f"Outstanding: ₹{self.var_outstanding.get()}", text_color="red", font=("Arial", 16, "bold")).pack(anchor="w")

            right = ctk.CTkFrame(main_frame)
            right.pack(side="right", fill="both", expand=True, padx=10, pady=10)
            
            ctk.CTkLabel(right, text="Update Payment", font=("Arial", 16, "bold")).pack(pady=10)
            ctk.CTkEntry(right, textvariable=self.var_pay_amount, placeholder_text="Amount Received (₹)").pack(pady=5)
            ctk.CTkEntry(right, textvariable=self.var_pay_date, placeholder_text="Date (YYYY-MM-DD)").pack(pady=5)
            ctk.CTkButton(right, text="Receive Payment", command=self.update_payment, fg_color="green").pack(pady=15)
            
            ctk.CTkLabel(right, text="Previous Payments", font=("Arial", 14, "bold")).pack(pady=(20,5))
            
            conn = self.get_db_connection()
            c = conn.cursor()
            c.execute("SELECT date_paid, amount_paid FROM payment_history WHERE customer_id=? ORDER BY date_paid DESC", (self.current_customer_id,))
            history = c.fetchall()
            conn.close()
            
            dates = [f"{h[0]} (₹{h[1]})" for h in history]
            if not dates: dates = ["No History"]
            
            self.history_menu = ctk.CTkOptionMenu(right, variable=self.var_pay_history_date, values=dates)
            self.history_menu.pack(pady=5)
            
        else:
            ctk.CTkLabel(content, text="Please search and select a customer to view payment details.", text_color="gray").pack(pady=20)

    def search_for_payment(self):
        query = self.var_pay_search.get().strip()
        if not query: return
        
        for widget in self.pay_results_frame.winfo_children():
            widget.destroy()

        conn = self.get_db_connection()
        c = conn.cursor()
        sql = "SELECT id, name, can, address FROM customers WHERE name LIKE ? OR can LIKE ?"
        param = f"%{query}%"
        c.execute(sql, (param, param))
        results = c.fetchall()
        conn.close()

        if not results:
            ctk.CTkLabel(self.pay_results_frame, text="No customers found.").pack()
        else:
            for res in results:
                btn_text = f"{res[1]} (CAN: {res[2]}) - {res[3]}"
                btn = ctk.CTkButton(self.pay_results_frame, text=btn_text, anchor="w", fg_color="transparent", border_width=1, border_color="gray",
                                    command=lambda r=res: self.select_payment_customer(r[0]))
                btn.pack(fill="x", pady=2)

    def select_payment_customer(self, cust_id):
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("SELECT * FROM customers WHERE id=?", (cust_id,))
        row = c.fetchone()
        conn.close()
        
        if row:
            self.load_customer(row)
            self.show_payment_tab()

    def update_payment(self):
        amt = self.var_pay_amount.get().strip()
        date = self.var_pay_date.get().strip()
        
        if not amt or not date:
            messagebox.showerror("Error", "Enter Amount and Date")
            return
            
        conn = self.get_db_connection()
        c = conn.cursor()
        
        c.execute("INSERT INTO payment_history (customer_id, can, amount_paid, date_paid) VALUES (?, ?, ?, ?)",
                  (self.current_customer_id, self.var_can.get(), amt, date))
        
        c.execute("UPDATE customers SET paid_amount=?, last_payment_date=?, outstanding_amount='0' WHERE id=?", 
                  (amt, date, self.current_customer_id))
        
        conn.commit()
        conn.close()
        
        # SYNC TO EXCEL WITH ERROR HANDLING
        self.sync_payment_to_excel(self.var_can.get(), amt, date)

        self.var_pay_amount.set("")
        self.show_payment_tab()

    def sync_payment_to_excel(self, can, amount, date):
        try:
            if os.path.exists(EXCEL_FILE):
                # Try to open file to check if locked
                try:
                    with open(EXCEL_FILE, "r+"): pass
                except IOError:
                    messagebox.showwarning("File Locked", "Please close the Excel file to sync changes.\nData saved to Database only.")
                    return

                df = pd.read_excel(EXCEL_FILE)
                df['CAN'] = df['CAN'].astype(str)
                can_str = str(can)
                
                if can_str in df['CAN'].values:
                    idx = df.index[df['CAN'] == can_str].tolist()[0]
                    df.at[idx, 'Paid'] = amount
                    df.at[idx, 'Payment Date'] = date
                    df.to_excel(EXCEL_FILE, index=False)
                    messagebox.showinfo("Success", "Payment Updated & Synced to Excel")
                else:
                    messagebox.showwarning("Excel Sync", "Payment saved to App, but Customer CAN not found in Excel.")
        except Exception as e:
            messagebox.showerror("Excel Error", f"Could not sync to Excel: {e}")

    # --- CUSTOMER MANAGER ---
    def show_customer_manager(self):
        self.clear_content_frame()
        form_scroll = ctk.CTkScrollableFrame(self.content_frame)
        form_scroll.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(form_scroll, text="Customer Details", font=("Arial", 20, "bold")).pack(pady=(0, 20))
        form_frame = ctk.CTkFrame(form_scroll, fg_color="transparent")
        form_frame.pack(fill="x")
        
        self.create_entry(form_frame, "CAN Number", self.var_can, 0, 0)
        self.create_entry(form_frame, "Customer Name", self.var_name, 0, 1)
        self.create_entry(form_frame, "Address", self.var_address, 1, 0, colspan=2)
        self.create_entry(form_frame, "Contact No", self.var_contact, 2, 0)
        
        areas = self.get_area_list()
        ctk.CTkLabel(form_frame, text="Area").grid(row=2, column=1, sticky="w", padx=10)
        ctk.CTkOptionMenu(form_frame, variable=self.var_area, values=areas).grid(row=3, column=1, sticky="ew", padx=10, pady=5)

        self.create_entry(form_frame, "STB Number", self.var_stb, 4, 0)
        
        ctk.CTkLabel(form_frame, text="STB Type").grid(row=4, column=1, sticky="w", padx=10)
        ctk.CTkOptionMenu(form_frame, variable=self.var_stb_type, values=["SD", "HD"]).grid(row=5, column=1, sticky="ew", padx=10, pady=5)
        
        self.create_entry(form_frame, "WiFi Router ID", self.var_router, 6, 0)
        self.create_entry(form_frame, "Net Account No", self.var_net_acc, 6, 1)
        self.create_entry(form_frame, "Install Date", self.var_install_date, 7, 0)
        self.create_entry(form_frame, "Recovery Date", self.var_recovery, 7, 1)
        self.create_entry(form_frame, "Monthly Rental", self.var_rental, 8, 0)
        self.create_entry(form_frame, "Total Connections", self.var_connections, 8, 1)
        self.create_entry(form_frame, "Smart Card No", self.var_smartcard, 9, 0)
        self.create_entry(form_frame, "Outstanding Amount", self.var_outstanding, 9, 1)
        
        ctk.CTkLabel(form_frame, text="Status").grid(row=10, column=0, sticky="w", padx=10)
        ctk.CTkOptionMenu(form_frame, variable=self.var_status, values=["Active", "Inactive"]).grid(row=11, column=0, sticky="ew", padx=10, pady=5)

        btn_frame = ctk.CTkFrame(form_scroll, fg_color="transparent")
        btn_frame.pack(fill="x", pady=20)
        ctk.CTkButton(btn_frame, text="Save / Update", command=self.save_customer, fg_color="green").pack(side="right", padx=10)
        
        if self.current_customer_id:
            ctk.CTkButton(btn_frame, text="Delete Customer", command=self.delete_customer, fg_color="#d9534f", hover_color="#c9302c").pack(side="right", padx=10)
        
        ctk.CTkButton(btn_frame, text="Clear", command=self.clear_form, fg_color="gray").pack(side="right", padx=10)

    def delete_customer(self):
        if not self.current_customer_id: return
        
        confirm = messagebox.askyesno("Delete Confirmation", f"Are you sure you want to delete {self.var_name.get()}?\nThis will remove them from the Database AND Excel.")
        if not confirm: return
        
        can_to_delete = str(self.var_can.get())

        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("DELETE FROM customers WHERE id=?", (self.current_customer_id,))
        conn.commit()
        conn.close()
        
        try:
            if os.path.exists(EXCEL_FILE):
                # Check for lock before trying
                try:
                    with open(EXCEL_FILE, "r+"): pass
                except IOError:
                    messagebox.showwarning("File Locked", "Deleted from App, but Excel file is open.\nClose Excel and try again to sync.")
                    self.clear_form()
                    self.show_dashboard()
                    return

                df = pd.read_excel(EXCEL_FILE)
                df['CAN'] = df['CAN'].astype(str)
                df = df[df['CAN'] != can_to_delete]
                df.to_excel(EXCEL_FILE, index=False)
                print("Deleted from Excel")
        except Exception as e:
            messagebox.showerror("Excel Error", f"Could not remove from Excel: {e}")

        messagebox.showinfo("Deleted", "Customer deleted successfully.")
        self.clear_form()
        self.show_dashboard()

    # --- INVENTORY ---
    def show_inventory(self):
        self.clear_content_frame()
        content = ctk.CTkScrollableFrame(self.content_frame)
        content.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(content, text="Inventory Management", font=("Arial", 20, "bold")).pack(anchor="w", pady=10)
        
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("SELECT item_name, quantity FROM inventory")
        items = c.fetchall()
        conn.close()
        
        for item in items:
            row = ctk.CTkFrame(content)
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=item[0], width=200, anchor="w").pack(side="left", padx=10)
            ctk.CTkLabel(row, text=f"Stock: {item[1]}", width=100).pack(side="left", padx=10)
        
        update_frame = ctk.CTkFrame(content)
        update_frame.pack(fill="x", pady=20)
        ctk.CTkLabel(update_frame, text="Update Stock").pack(anchor="w", padx=10, pady=5)
        
        item_names = [i[0] for i in items]
        if item_names:
            self.inv_menu = ctk.CTkOptionMenu(update_frame, values=item_names, variable=self.var_inv_item)
            self.inv_menu.pack(side="left", padx=10)
            ctk.CTkEntry(update_frame, textvariable=self.var_inv_qty, placeholder_text="Qty", width=60).pack(side="left", padx=10)
            ctk.CTkButton(update_frame, text="Add (+)", command=lambda: self.update_inventory(1), width=80).pack(side="left", padx=5)
            ctk.CTkButton(update_frame, text="Remove (-)", command=lambda: self.update_inventory(-1), width=80, fg_color="red").pack(side="left", padx=5)

    def update_inventory(self, multiplier):
        try:
            qty = int(self.var_inv_qty.get())
            item = self.var_inv_item.get()
            conn = self.get_db_connection()
            c = conn.cursor()
            c.execute("UPDATE inventory SET quantity = quantity + ? WHERE item_name = ?", (qty * multiplier, item))
            conn.commit()
            conn.close()
            self.show_inventory() 
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number")

    # --- COMPLAINTS ---
    def show_complaints(self):
        self.clear_content_frame()
        content = ctk.CTkScrollableFrame(self.content_frame)
        content.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(content, text="Complaints System", font=("Arial", 20, "bold")).pack(anchor="w", pady=10)

        if self.current_customer_id:
            f = ctk.CTkFrame(content)
            f.pack(fill="x", pady=10)
            ctk.CTkLabel(f, text=f"New Complaint for: {self.var_name.get()}").pack(anchor="w", padx=10, pady=5)
            ctk.CTkEntry(f, textvariable=self.var_complaint_issue, placeholder_text="Describe Issue", width=400).pack(side="left", padx=10, pady=10)
            ctk.CTkButton(f, text="Log Complaint", command=self.log_complaint).pack(side="left", padx=10)
        else:
            ctk.CTkLabel(content, text="Select a customer to log a new complaint.", text_color="gray").pack(anchor="w", padx=20)

        ctk.CTkLabel(content, text="All Active Complaints", font=("Arial", 16, "bold")).pack(anchor="w", pady=(20, 10))
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("SELECT id, customer_name, issue, date_logged, status FROM complaints WHERE status='Open' ORDER BY date_logged DESC")
        rows = c.fetchall()
        
        if not rows:
            ctk.CTkLabel(content, text="No open complaints.").pack(pady=10)
            
        for r in rows:
            card = ctk.CTkFrame(content)
            card.pack(fill="x", pady=5)
            ctk.CTkLabel(card, text=f"{r[3]}", width=100).pack(side="left", padx=5)
            ctk.CTkLabel(card, text=f"{r[1]}", width=150, font=("Arial", 12, "bold")).pack(side="left", padx=5)
            ctk.CTkLabel(card, text=f"{r[2]}", width=300, anchor="w").pack(side="left", padx=5)
            ctk.CTkButton(card, text="Mark Resolved", fg_color="green", width=100,
                          command=lambda cid=r[0]: self.resolve_complaint(cid)).pack(side="right", padx=10, pady=5)
        conn.close()

    def log_complaint(self):
        if not self.var_complaint_issue.get(): return
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("INSERT INTO complaints (customer_id, customer_name, issue, date_logged, status) VALUES (?, ?, ?, ?, 'Open')",
                  (self.current_customer_id, self.var_name.get(), self.var_complaint_issue.get(), datetime.date.today()))
        conn.commit()
        conn.close()
        self.var_complaint_issue.set("")
        self.show_complaints()

    def resolve_complaint(self, complaint_id):
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("UPDATE complaints SET status='Resolved', date_resolved=? WHERE id=?", 
                  (datetime.date.today(), complaint_id))
        conn.commit()
        conn.close()
        messagebox.showinfo("Success", "Complaint marked as Resolved.")
        self.show_complaints()

    # --- REPORTS ---
    def show_reports(self):
        self.clear_content_frame()
        content = ctk.CTkFrame(self.content_frame)
        content.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(content, text="Monthly Statement Report", font=("Arial", 20, "bold")).pack(pady=20)
        
        f = ctk.CTkFrame(content)
        f.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(f, text="Area:").pack(side="left", padx=5)
        self.var_filter_area = StringVar(value="All")
        areas = self.get_area_list()
        ctk.CTkOptionMenu(f, values=["All"] + areas, variable=self.var_filter_area).pack(side="left", padx=5)
        
        ctk.CTkLabel(f, text="Start Date (YYYY-MM-DD):").pack(side="left", padx=(20,5))
        ctk.CTkEntry(f, textvariable=self.var_start_date, width=120, placeholder_text="2025-01-01").pack(side="left", padx=5)
        
        ctk.CTkLabel(f, text="End Date:").pack(side="left", padx=5)
        ctk.CTkEntry(f, textvariable=self.var_end_date, width=120, placeholder_text="2025-01-31").pack(side="left", padx=5)
        
        ctk.CTkButton(f, text="Export Filtered Data", command=self.export_report).pack(side="right", padx=20)
        
        ctk.CTkLabel(content, text="Note: Dates filter based on 'Recovery/Payment Date'").pack(pady=10)

    # --- SETTINGS (PASSWORD PROTECTED) ---
    def show_settings(self):
        dialog = ctk.CTkInputDialog(text="Enter Developer Password:", title="Admin Access")
        password = dialog.get_input()
        
        if password != ADMIN_PASSWORD:
            messagebox.showerror("Access Denied", "Incorrect Password.")
            return

        self.clear_content_frame()
        content = ctk.CTkScrollableFrame(self.content_frame)
        content.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(content, text="Settings (Developer Mode)", font=("Arial", 22, "bold")).pack(anchor="w", pady=20)

        s1 = ctk.CTkFrame(content)
        s1.pack(fill="x", pady=10)
        ctk.CTkLabel(s1, text="Appearance Mode").pack(anchor="w", padx=10, pady=5)
        ctk.CTkSegmentedButton(s1, values=["Dark", "Light"], command=ctk.set_appearance_mode).pack(padx=10, pady=10, anchor="w")

        s2 = ctk.CTkFrame(content)
        s2.pack(fill="x", pady=10)
        ctk.CTkLabel(s2, text="Data Backup").pack(anchor="w", padx=10, pady=5)
        ctk.CTkButton(s2, text="One-Click Backup", command=self.backup_db, fg_color="#f0ad4e").pack(padx=10, pady=10, anchor="w")

    # --- LOGIC & HELPERS ---
    def get_area_list(self):
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("SELECT area_name FROM areas ORDER BY area_name ASC")
        areas = [row[0] for row in c.fetchall()]
        conn.close()
        return areas if areas else ["Unassigned"]

    def add_area(self):
        new_area = self.var_new_area.get().strip().upper()
        if new_area:
            conn = self.get_db_connection()
            try:
                conn.execute("INSERT INTO areas (area_name) VALUES (?)", (new_area,))
                conn.commit()
                messagebox.showinfo("Success", "Area Added")
                self.var_new_area.set("")
                self.show_dashboard() 
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Area already exists")
            conn.close()

    def delete_area(self):
        area = self.var_new_area.get().strip().upper()
        if not area:
            messagebox.showwarning("Warning", "Please enter area name to delete.")
            return
        conn = self.get_db_connection()
        c = conn.cursor()
        c.execute("SELECT 1 FROM areas WHERE area_name=?", (area,))
        exists = c.fetchone()
        if not exists:
            messagebox.showerror("Error", "The area doesn't exist.")
        else:
            c.execute("DELETE FROM areas WHERE area_name=?", (area,))
            conn.commit()
            messagebox.showinfo("Success", "Area Deleted")
            self.var_new_area.set("")
            self.show_dashboard()
        conn.close()

    def perform_search(self, event=None):
        query = self.search_entry.get().strip()
        if not query: return
        conn = self.get_db_connection()
        c = conn.cursor()
        sql = "SELECT * FROM customers WHERE name LIKE ? OR can LIKE ? OR stb_no LIKE ?"
        param = f"%{query}%"
        c.execute(sql, (param, param, param))
        results = c.fetchall()
        conn.close()
        if len(results) == 0: messagebox.showinfo("Data doesn't exist", "No customer found.")
        elif len(results) == 1: 
            self.load_customer(results[0])
            self.show_customer_manager()
        else: self.resolve_duplicates(results)

    def resolve_duplicates(self, results):
        top = ctk.CTkToplevel(self)
        top.title("Select Customer")
        scroll = ctk.CTkScrollableFrame(top, width=400, height=300)
        scroll.pack()
        for res in results:
            btn = ctk.CTkButton(scroll, text=f"{res[2]} | {res[1]}", command=lambda r=res: [self.load_customer(r), self.show_customer_manager(), top.destroy()])
            btn.pack(pady=2, padx=5, fill="x")

    def load_customer(self, row):
        # Maps database tuple to variables. CAUTION: Schema changed, index shifts are likely.
        # DB: id, can, name, address, contact, stb, stb_type, recovery, area, smart, router, net_acc, install, rental, conn, status, dep, wifi_pay, last_pay, paid, outstanding
        self.current_customer_id = row[0]
        self.var_can.set(row[1])
        self.var_name.set(row[2])
        self.var_address.set(row[3])
        self.var_contact.set(row[4])
        self.var_stb.set(row[5])
        self.var_stb_type.set(row[6] if row[6] else "SD")
        self.var_recovery.set(row[7])
        self.var_area.set(row[8])
        self.var_smartcard.set(row[9])
        self.var_router.set(row[10])
        self.var_net_acc.set(row[11])
        self.var_install_date.set(row[12])
        self.var_rental.set(row[13])
        self.var_connections.set(row[14])
        self.var_status.set(row[15])
        self.var_outstanding.set(row[19] if row[19] else "0")

    def save_customer(self):
        if not self.var_name.get(): return
        conn = self.get_db_connection()
        c = conn.cursor()
        data = (
            self.var_can.get(), self.var_name.get(), self.var_address.get(), self.var_contact.get(),
            self.var_stb.get(), self.var_stb_type.get(), self.var_recovery.get(), self.var_area.get(), 
            self.var_smartcard.get(), self.var_router.get(), self.var_net_acc.get(), self.var_install_date.get(), 
            self.var_rental.get(), self.var_connections.get(), self.var_status.get(), self.var_outstanding.get()
        )
        
        if self.current_customer_id:
            c.execute("UPDATE customers SET can=?, name=?, address=?, contact_no=?, stb_no=?, stb_type=?, recovery_date=?, area=?, smart_card_no=?, wifi_router_id=?, net_acc_no=?, install_date=?, monthly_rental=?, total_connections=?, status=?, outstanding_amount=? WHERE id=?", data + (self.current_customer_id,))
            messagebox.showinfo("Success", "Updated")
        else:
            # Note: inserts new customers with outstanding default 0
            c.execute("INSERT INTO customers (can, name, address, contact_no, stb_no, stb_type, recovery_date, area, smart_card_no, wifi_router_id, net_acc_no, install_date, monthly_rental, total_connections, status, outstanding_amount) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", data)
            messagebox.showinfo("Success", "Created")
        
        conn.commit()
        conn.close()
        
        # EXCEL SYNC (Runs for both ADD and UPDATE)
        self.sync_full_customer_to_excel(self.var_can.get(), self.var_name.get(), self.var_address.get(), self.var_contact.get(), self.var_stb.get(), self.var_recovery.get())

    def sync_full_customer_to_excel(self, can, name, address, contact, stb, date):
        try:
            if os.path.exists(EXCEL_FILE):
                # LOCK CHECK
                try:
                    with open(EXCEL_FILE, "r+"): pass
                except IOError:
                    messagebox.showwarning("File Locked", "Customer saved to Database, but NOT Excel.\nPlease close the Excel file.")
                    return

                df = pd.read_excel(EXCEL_FILE)
                df['CAN'] = df['CAN'].astype(str)
                can_str = str(can)
                
                if can_str in df['CAN'].values:
                    # Update existing
                    idx = df.index[df['CAN'] == can_str].tolist()[0]
                    df.at[idx, 'Customer Name'] = name
                    df.at[idx, 'Address'] = address
                    df.at[idx, 'Contact'] = contact
                    df.at[idx, 'STB No'] = stb
                    df.at[idx, 'Payment Date'] = date
                else:
                    # Append New
                    new_row = {
                        'CAN': can, 'Customer Name': name, 'Address': address, 
                        'Contact': contact, 'STB No': stb, 'Payment Date': date, 'Paid': '0'
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                
                df.to_excel(EXCEL_FILE, index=False)
        except Exception as e:
            messagebox.showerror("Excel Error", f"Sync failed: {e}")

    def open_whatsapp_web(self):
        phone = self.var_contact.get().strip()
        if not phone: 
            messagebox.showerror("Data doesn't exist", "No contact number found for this customer.")
            return
        
        if len(phone) == 10: phone = "91" + phone
        msg = f"""📢 *{BUSINESS_NAME} - PAYMENT REMINDER*
        
Hello *{self.var_name.get()}*,

🙏 *मराठी:*
आपले केबल/इंटरनेट बिल जनरेट झाले आहे. कृपया आपली सेवा अविरत सुरू ठेवण्यासाठी देय तारखेपूर्वी पैसे भरावे.
🔸 रक्कम: ₹{self.var_rental.get()}
🔸 देय तारीख: {self.var_recovery.get()}

🙏 *हिंदी:*
आपका केबल/इंटरनेट बिल जनरेट हो गया है। कृपया अपनी सेवा निर्बाध रखने के लिए देय तिथि से पहले भुगतान करें।
🔸 राशि: ₹{self.var_rental.get()}
🔸 देय तिथि: {self.var_recovery.get()}

🙏 *English:*
Your Cable/Internet bill has been generated. Please pay before the due date to enjoy uninterrupted services.
🔸 Amount: ₹{self.var_rental.get()}
🔸 Due Date: {self.var_recovery.get()}

💳 *Support:* {SUPPORT_CONTACT}
Thank you for choosing {BUSINESS_NAME}.
"""
        encoded_msg = urllib.parse.quote(msg)
        webbrowser.open(f"https://web.whatsapp.com/send?phone={phone}&text={encoded_msg}")

    def generate_receipt_pdf(self):
        html_content = f"""
        <html>
        <head><title>Invoice</title></head>
        <body style="font-family: Arial, sans-serif; padding: 40px; background: #f9f9f9;">
            <div style="background: white; padding: 30px; border: 1px solid #ccc; max-width: 800px; margin: auto;">
                <div style="display: flex; justify-content: space-between;">
                    <div>
                        <h1 style="color: #333;">{BUSINESS_NAME}</h1>
                        <p>{BUSINESS_ADDRESS}<br>Support: {SUPPORT_CONTACT}</p>
                    </div>
                    <div style="text-align: right;">
                        <h3>INVOICE</h3>
                        <p>Date: {datetime.date.today()}<br>Due: {self.var_recovery.get()}</p>
                    </div>
                </div>
                <hr>
                <div style="margin-top: 20px;">
                    <b>BILL TO:</b><br>
                    {self.var_name.get()} (CAN: {self.var_can.get()})<br>
                    {self.var_address.get()}<br>
                    {self.var_contact.get()}
                </div>
                <table style="width: 100%; margin-top: 30px; border-collapse: collapse;">
                    <tr style="background: #eee;">
                        <th style="text-align: left; padding: 10px;">Description</th>
                        <th style="text-align: right; padding: 10px;">Amount</th>
                    </tr>
                    <tr>
                        <td style="padding: 10px; border-bottom: 1px solid #ddd;">Monthly Rental ({self.var_status.get()})</td>
                        <td style="text-align: right; padding: 10px; border-bottom: 1px solid #ddd;">{self.var_rental.get()}</td>
                    </tr>
                    <tr>
                        <td style="padding: 10px;"><b>TOTAL DUE</b></td>
                        <td style="text-align: right; padding: 10px;"><b>{self.var_rental.get()}</b></td>
                    </tr>
                </table>
                <div style="margin-top: 50px; font-size: 12px; color: #777; text-align: center;">
                    {self.var_invoice_footer.get()}
                </div>
            </div>
            <script>window.print();</script>
        </body>
        </html>
        """
        try:
            with open("temp_receipt.html", "w", encoding="utf-8") as f:
                f.write(html_content)
            webbrowser.open("file://" + os.path.realpath("temp_receipt.html"))
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def export_report(self):
        area = self.var_filter_area.get()
        start = self.var_start_date.get().strip()
        end = self.var_end_date.get().strip()
        conn = self.get_db_connection()
        query = "SELECT * FROM customers WHERE 1=1"
        if area != "All": query += f" AND area='{area}'"
        df = pd.read_sql_query(query, conn)
        conn.close()
        if start and end:
            try:
                df['temp_date'] = pd.to_datetime(df['recovery_date'], errors='coerce')
                s_date = pd.to_datetime(start)
                e_date = pd.to_datetime(end)
                df = df[(df['temp_date'] >= s_date) & (df['temp_date'] <= e_date)]
                df = df.drop(columns=['temp_date'])
            except Exception as e:
                messagebox.showerror("Date Error", f"Invalid date format. Use YYYY-MM-DD.\n{e}")
                return
        if df.empty:
            messagebox.showinfo("Data doesn't exist", "No records found.")
            return
        filename = "Filtered_Report.xlsx"
        df.to_excel(filename, index=False)
        messagebox.showinfo("Export Successful", f"Saved as {filename}")

    def backup_db(self):
        filename = filedialog.asksaveasfilename(defaultextension=".db")
        if filename:
            shutil.copy(DB_FILE, filename)
            messagebox.showinfo("Backup", "Database Backed up!")

    def card(self, parent, title, val, color):
        f = ctk.CTkFrame(parent, fg_color=color)
        ctk.CTkLabel(f, text=title, text_color="white", font=("Arial", 12)).pack(pady=(10,5))
        ctk.CTkLabel(f, text=val, text_color="white", font=("Arial", 18, "bold")).pack(pady=(0,10))
        return f

    def create_entry(self, parent, label, variable, r, c, colspan=1):
        ctk.CTkLabel(parent, text=label).grid(row=r*2, column=c, sticky="w", padx=10)
        ctk.CTkEntry(parent, textvariable=variable, width=300 if colspan==1 else 620).grid(row=r*2+1, column=c, columnspan=colspan, sticky="ew", padx=10, pady=(0,10))

    def clear_content_frame(self):
        try: self.content_frame.destroy()
        except: pass
        self.content_frame = ctk.CTkFrame(self.main_view, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)

    def clear_form(self):
        self.current_customer_id = None
        for v in [self.var_can, self.var_name, self.var_address, self.var_contact, self.var_stb, 
                  self.var_smartcard, self.var_router, self.var_net_acc, self.var_rental, 
                  self.var_connections, self.var_recovery, self.var_install_date]:
            v.set("")
        self.var_area.set("Unassigned")
        self.var_status.set("Active")

if __name__ == "__main__":
    app = CableManagerApp()
    app.mainloop()