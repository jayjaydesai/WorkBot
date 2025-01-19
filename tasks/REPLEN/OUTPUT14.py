import pandas as pd
import os

# Define file paths
output13_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT13.xlsx"
output14_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT14.xlsx"

# Load the OUTPUT13 file
output13_df = pd.read_excel(output13_path)

# Normalize column names
output13_df.columns = output13_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = [
    'ITEM NUMBER', 'STOCK FOR REPLEN2', 'FINAL STATUS TO ADD',
    'AVAILABLE POSTED QTY TO ADD', 'BALANCE STOCK FOR REPLEN', 'LEVEL', 'POSTED QUANTITY'
]
for col in required_columns:
    if col not in output13_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT13.xlsx")

# Step 2: Function to allocate replen stock for "BALANCE" status
def allocate_replen_stock(group):
    balance_stock = group['BALANCE STOCK FOR REPLEN'].iloc[0]  # Get total balance stock
    sorted_group = group.sort_values(by='POSTED QUANTITY', ascending=False).copy()
    final_stock = []  # List to hold the allocated stock for each row

    for _, row in sorted_group.iterrows():
        if balance_stock > 0:  # If there is still stock to allocate
            if row['POSTED QUANTITY'] <= balance_stock:
                final_stock.append(row['POSTED QUANTITY'])  # Allocate full posted quantity
                balance_stock -= row['POSTED QUANTITY']  # Reduce balance stock
            else:
                final_stock.append(balance_stock)  # Allocate remaining balance stock
                balance_stock = 0  # Balance stock is now zero
        else:
            final_stock.append(0)  # No stock to allocate
    sorted_group['FINAL STOCK FOR REPLEN'] = final_stock
    return sorted_group

# Apply allocation logic to groups of ITEM NUMBER where FINAL STATUS TO ADD is BALANCE
output13_df = output13_df.groupby('ITEM NUMBER', group_keys=False).apply(
    lambda group: allocate_replen_stock(group) if group['FINAL STATUS TO ADD'].iloc[0] == "BALANCE" else group
)

# Step 3: For rows with FINAL STATUS TO ADD = "FULL", use POSTED QUANTITY
output13_df['FINAL STOCK FOR REPLEN'] = output13_df.apply(
    lambda row: row['POSTED QUANTITY'] if row['FINAL STATUS TO ADD'] == "FULL" else row['FINAL STOCK FOR REPLEN'],
    axis=1
)

# Step 4: Calculate "LEVEL STATUS"
def calculate_level_status(level):
    if level == "A":
        return "A LOCATION REPLENS"
    elif level in ["B", "C"]:
        return "PALLET STACKER REPLENS"
    else:
        return "REACH REPLENS"

# Apply the logic to calculate "LEVEL STATUS"
output13_df['LEVEL STATUS'] = output13_df['LEVEL'].apply(calculate_level_status)

# Save the final dataframe as OUTPUT14
try:
    output13_df.to_excel(output14_path, index=False)
    print(f"OUTPUT14 file saved at: {output14_path}")
except Exception as e:
    print(f"An error occurred while saving OUTPUT14: {e}")
