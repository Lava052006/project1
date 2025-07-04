import pandas as pd
from datetime import datetime

EXCEL_PATH = "pg_management.xlsx"

# Load Excel sheets
def load_data():
    rooms_df = pd.read_excel(EXCEL_PATH, sheet_name='Rooms')
    customers_df = pd.read_excel(EXCEL_PATH, sheet_name='Customers')
    payments_df = pd.read_excel(EXCEL_PATH, sheet_name='Payments')
    print(rooms_df)
    print(customers_df)
    print(payments_df)
    return rooms_df, customers_df, payments_df

# Save updated data
def save_data(rooms_df, customers_df, payments_df):
    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='w') as writer:
        rooms_df.to_excel(writer, sheet_name='Rooms', index=False)
        customers_df.to_excel(writer, sheet_name='Customers', index=False)
        payments_df.to_excel(writer, sheet_name='Payments', index=False)

# Add a new customer
def add_customer():
    name = input("Customer Name: ")
    phone = input("Phone Number: ")
    room = input("Room ID: ")
    checkin = input("Check-In Date (YYYY-MM-DD): ")
    deposit = int(input("Deposit (₹): "))
    electricity = int(input("Initial Electricity (₹): "))

    customers_df.loc[len(customers_df.index)] = [name, phone, room, checkin, "", deposit, electricity, "No", "No"]
    rooms_df.loc[rooms_df['Room ID'] == room, 'Status'] = "Occupied"
    print(f"{name} added successfully.")

# Record a payment
def record_payment():
    name = input("Customer Name: ")
    month = input("Month: ")
    rent = int(input("Rent Paid: "))
    electricity = int(input("Electricity Paid: "))
    total = rent + electricity
    date = datetime.today().strftime('%Y-%m-%d')

    payments_df.loc[len(payments_df.index)] = [name, month, rent, electricity, total, date]

    customers_df.loc[customers_df['Customer Name'] == name, 'Paid Rent'] = "Yes"
    customers_df.loc[customers_df['Customer Name'] == name, 'Paid Electricity'] = "Yes"
    print("Payment recorded.")

# Checkout customer
def checkout_customer():
    name = input("Customer Name: ")
    date = input("Checkout Date (YYYY-MM-DD): ")
    customers_df.loc[customers_df['Customer Name'] == name, 'Check-Out'] = date
    room = customers_df.loc[customers_df['Customer Name'] == name, 'Room ID'].values[0]
    rooms_df.loc[rooms_df['Room ID'] == room, 'Status'] = "Available"
    print(f"{name} checked out successfully.")

# View pending payments
def view_pending():
    pending = customers_df[
        (customers_df['Paid Rent'].str.lower() != 'yes') |
        (customers_df['Paid Electricity'].str.lower() != 'yes')
    ]
    print("\nPending Payments:")
    print(pending[['Customer Name', 'Room ID', 'Electricity (₹)', 'Paid Rent', 'Paid Electricity']])

# View total income
def total_income():
    income = payments_df['Total'].sum()
    print(f"\n💰 Total Income Collected: ₹{income}")

# Main menu
if __name__ == "__main__":
    rooms_df, customers_df, payments_df = load_data()

    while True:
        print("\n--- PG MANAGEMENT SYSTEM ---")
        print("1. Add New Customer")
        print("2. Record Payment")
        print("3. Checkout Customer")
        print("4. View Pending Payments")
        print("5. View Total Income")
        print("6. Exit")

        choice = input("Enter your choice: ")

        if choice == '1':
            add_customer()
        elif choice == '2':
            record_payment()
        elif choice == '3':
            checkout_customer()
        elif choice == '4':
            view_pending()
        elif choice == '5':
            total_income()
        elif choice == '6':
            save_data(rooms_df, customers_df, payments_df)
            print("✅ Data saved. Goodbye!")
            break
        else:
            print("❌ Invalid choice. Try again.")

