import pandas as pd

# 1. Define the Dummy Data
data = {
    'CAN': [1001, 1002, 1003, 1004, 1005],
    'Customer Name': [
        "Rahul Sharma", 
        "Priya Verma", 
        "Amitabh Patel", 
        "Sneha Gupta", 
        "Vikram Singh"
    ],
    'Address': [
        "Flat 101, Galaxy Apts, Nagpur", 
        "Plot 45, Shivaji Nagar, Nagpur", 
        "12, Main Road, Jaripatka", 
        "B-Wing, Sunrise Society, Sadar", 
        "Near Old Temple, Mahal"
    ],
    'Contact': [
        "9876543210", 
        "9123456789", 
        "9988776655", 
        "9000011111", 
        "8888822222"
    ],
    'STB No': [
        "STB-998877", 
        "STB-112233", 
        "STB-445566", 
        "STB-778899", 
        "STB-000000"
    ],
    'Payment Date': [
        "2025-12-01", 
        "2025-12-05", 
        "2025-12-10", 
        "2025-12-15", 
        "2025-12-20"
    ],
    'Paid': [
        "500", 
        "650", 
        "400", 
        "1000", 
        "500"
    ]
}

# 2. Create DataFrame
df = pd.DataFrame(data)

# 3. Save to Excel
filename = "Sample_Customer_List.xlsx"
df.to_excel(filename, index=False)

print(f"✅ Successfully created '{filename}'.")
print("You can now upload this file to GitHub.")