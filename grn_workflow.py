import pandas as pd
import os
from datetime import datetime

# Define the Excel file and sheet name
EXCEL_FILE = "grn_data.xlsx"
SHEET_NAME = "GRNs"

def get_grn_data():
    """
    Loads GRN data from the Excel file. If the file doesn't exist, it creates
    a new DataFrame with the required columns.
    """
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
            print(f"Loaded {len(df)} GRN records from {EXCEL_FILE}.")
            return df
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return create_new_grn_df()
    else:
        return create_new_grn_df()

def create_new_grn_df():
    """
    Creates a new DataFrame with the required columns for GRN data.
    All columns are created with a string dtype to prevent warnings when
    adding text to empty columns.
    """
    columns = [
        "GRN_ID", "Customer Name", "Warranty Status", "Gate Entry No", "Gate Entry Date",
        "CRN No", "DC No", "RGP/NRGP No", "Date", "Goods Description", "Qty Supplied",
        "UOM", "Received Qty", "QC Accepted", "QC Rejected", "Remarks", "Prepared By Stores",
        "Reviewed By PPC", "Inspected & Reworked By QA", "Acknowledged By Marketing",
        "General Remarks", "Status"
    ]
    # Creating a DataFrame with an explicit data type for all columns
    data_dict = {col: [] for col in columns}
    print(f"Created a new DataFrame for {EXCEL_FILE}")
    return pd.DataFrame(data_dict)

def save_grn_data(df):
    """Saves the DataFrame to the Excel file."""
    try:
        df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)
        print(f"Data saved successfully to {EXCEL_FILE}.")
    except Exception as e:
        print(f"Error saving data to Excel file: {e}")

def add_new_grn():
    """Prompts the user for details to add a new GRN record."""
    df = get_grn_data()
    print("\n--- Add New GRN ---")
    
    # Get user input for GRN details
    customer_name = input("Enter Customer Name & Address: ")
    warranty_status = input("Enter Warranty Status (warranty/out of warranty): ")
    gate_entry_no = input("Enter Gate Entry No: ")
    gate_entry_date = input("Enter Gate Entry Date (YYYY-MM-DD): ")
   
    #duplicates check using gate_entry_no
    if gate_entry_no in df['Gate Entry No'].values:
    print(f"GRN for Gate Entry No {gate_entry_no} already exists. Entry skipped.")
    return

    # Get material details
    goods_description = input("Enter Goods Description: ")
    qty_supplied = input("Enter Quantity Supplied: ")
    uom = input("Enter UOM: ")
    received_qty = input("Enter Received Quantity: ")
    qc_accepted = input("Enter QC Accepted Quantity: ")
    qc_rejected = input("Enter QC Rejected Quantity: ")
    remarks = input("Enter Remarks: ")

    # Get sign-off details
    prepared_by = input("Prepared by Stores: ")
    reviewed_by = input("Reviewed by PPC: ")
    general_remarks = input("Enter General Remarks: ")
    
    # Assign a unique GRN ID
    grn_id = f"GRN-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    # Create a new record as a dictionary
    new_record = {
        "GRN_ID": grn_id,
        "Customer Name": customer_name,
        "Warranty Status": warranty_status,
        "Gate Entry No": gate_entry_no,
        "Gate Entry Date": gate_entry_date,
        "CRN No": "",
        "DC No": "",
        "RGP/NRGP No": "",
        "Date": datetime.now().strftime('%Y-%m-%d'),
        "Goods Description": goods_description,
        "Qty Supplied": qty_supplied,
        "UOM": uom,
        "Received Qty": received_qty,
        "QC Accepted": qc_accepted,
        "QC Rejected": qc_rejected,
        "Remarks": remarks,
        "Prepared By Stores": prepared_by,
        "Reviewed By PPC": reviewed_by,
        "Inspected & Reworked By QA": "",
        "Acknowledged By Marketing": "",
        "General Remarks": general_remarks,
        "Status": "Pending"
    }

    # Append the new record to the DataFrame and save
    new_df = pd.DataFrame([new_record])
    df = pd.concat([df, new_df], ignore_index=True)
    save_grn_data(df)

def update_grn_status():
    """Allows a user to update the status of an existing GRN."""
    df = get_grn_data()
    print("\n--- Update GRN Status ---")
    grn_id = input("Enter the GRN ID to update: ")
    
    # Find the GRN record
    record_index = df.index[df['GRN_ID'] == grn_id].tolist()
    if not record_index:
        print(f"GRN with ID '{grn_id}' not found.")
        return

    record_index = record_index[0]
    current_status = df.loc[record_index, 'Status']
    print(f"Current status for {grn_id} is: {current_status}")
    
    # Provide status update options based on current status
    if current_status == "Pending":
        new_status = input("Enter new status (Approved/Rejected): ").capitalize()
        if new_status in ["Approved", "Rejected"]:
            qa_signature = input("Enter name of QA Inspector: ")
            df.loc[record_index, 'Status'] = new_status
            df.loc[record_index, 'Inspected & Reworked By QA'] = qa_signature
            save_grn_data(df)
        else:
            print("Invalid status. Please enter 'Approved' or 'Rejected'.")
            
    elif current_status == "Approved":
        final_status = input("Enter new status (Finalized): ").capitalize()
        if final_status == "Finalized":
            accounts_signature = input("Enter name of Accounts HOD: ")
            df.loc[record_index, 'Status'] = "Finalized"
            df.loc[record_index, 'Acknowledged By Marketing'] = accounts_signature
            save_grn_data(df)
        else:
            print("Invalid status. Only 'Finalized' can be set for Approved GRNs.")

    elif current_status == "Finalized":
        print("This GRN has already been finalized and cannot be updated.")
    
    else:
        print("Invalid status. No update options available.")

def generate_pending_report():
    """Generates a report of all pending GRNs."""
    df = get_grn_data()
    print("\n--- Generating Pending GRN Report ---")
    pending_grns = df[df['Status'] == "Pending"]
    
    if pending_grns.empty:
        print("No pending GRNs found.")
    else:
        print(f"Found {len(pending_grns)} pending GRN(s).")
        # Displaying a subset of columns for readability
        report_df = pending_grns[["GRN_ID", "Gate Entry No", "Date", "Customer Name", "Status"]]
        print("\n", report_df.to_string())

def view_grn_details():
    """Prints a single GRN record in a formatted way."""
    df = get_grn_data()
    print("\n--- View GRN Details ---")
    grn_id = input("Enter the GRN ID to view: ")
    
    record = df[df['GRN_ID'] == grn_id]
    
    if record.empty:
        print(f"GRN with ID '{grn_id}' not found.")
        return

    record = record.iloc[0]

    print("\n" + "=" * 70)
    print(f"GRN Details for ID: {record['GRN_ID']}")
    print("=" * 70 + "\n")
    
    print(f"Customer Name & Address: {record['Customer Name']}".ljust(50))
    print(f"Warranty Status: {record['Warranty Status']}\n")
    
    print(f"Gate Entry No: {record['Gate Entry No']}".ljust(50))
    print(f"Gate Entry Date: {record['Gate Entry Date']}\n")
    
    print("-" * 70)
    print("Goods Details:")
    print("-" * 70)
    
    print(f"  Description: {record['Goods Description']}")
    print(f"  Qty Supplied: {record['Qty Supplied']} {record['UOM']}")
    print(f"  Received Qty: {record['Received Qty']}")
    print(f"  QC Accepted: {record['QC Accepted']}")
    print(f"  QC Rejected: {record['QC Rejected']}")
    print(f"  Remarks: {record['Remarks']}\n")
    
    print("-" * 70)
    print("Sign-off & Status:")
    print("-" * 70)
    
    print(f"  Prepared By Stores: {record['Prepared By Stores']}".ljust(40))
    print(f"  Reviewed By PPC: {record['Reviewed By PPC']}")
    print(f"  Inspected & Reworked By QA: {record['Inspected & Reworked By QA']}")
    print(f"  Acknowledged By Marketing: {record['Acknowledged By Marketing']}")
    print(f"  General Remarks: {record['General Remarks']}\n")
    
    print("=" * 70)
    print(f"Current GRN Status: {record['Status']}")
    print("=" * 70 + "\n")

def main_menu():
    """Displays the main menu and handles user input."""
    while True:
        print("\n--- GRN Workflow Automation ---")
        print("1. Add New GRN")
        print("2. Update GRN Status (QC / Accounts)")
        print("3. Generate Pending GRN Report")
        print("4. View GRN Details (Formatted Report)")
        print("5. Exit")
        
        choice = input("Enter your choice (1-5): ")
        
        if choice == '1':
            add_new_grn()
        elif choice == '2':
            update_grn_status()
        elif choice == '3':
            generate_pending_report()
        elif choice == '4':
            view_grn_details()
        elif choice == '5':
            print("Exiting application.")
            break
        else:
            print("Invalid choice. Please enter a number from 1 to 5.")

if __name__ == "__main__":
    # Ensure pandas and openpyxl are installed
    try:
        import pandas as pd
        import openpyxl
    except ImportError:
        print("Required Python libraries 'pandas' and 'openpyxl' are not installed.")
        print("Please install them using the following command:")
        print("pip install pandas openpyxl")
        exit()

    main_menu()

