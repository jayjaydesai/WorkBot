import pandas as pd
import os

# Define file paths
output5_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT5.xlsx"
output6a_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT6A.xlsx"

# Load the OUTPUT5 file
output5_df = pd.read_excel(output5_path)

# Normalize column names
output5_df.columns = output5_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = [
    'ITEM NUMBER', 'LOCATION_X', 'POSTED QUANTITY', 
    'DIFF', 'LEVEL', 'LICENCE PLATE'
]
for col in required_columns:
    if col not in output5_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT5.xlsx")

# Step 2: Sort the DataFrame by LEVEL and POSTED QUANTITY in descending order
output5_df['LEVEL_PRIORITY'] = output5_df['LEVEL'].apply(lambda x: ord(x.upper()) - ord('A'))  # Convert LEVEL to a numerical priority
output5_df.sort_values(
    by=['ITEM NUMBER', 'LEVEL_PRIORITY', 'POSTED QUANTITY'],
    ascending=[True, True, False],  # Sort by LEVEL first, then by POSTED QUANTITY descending
    inplace=True
)

# Step 3: Define a function to calculate STOCK FOR REPLEN
def calculate_stock_for_replen(df):
    result = []
    grouped = df.groupby('ITEM NUMBER')

    for item_number, group in grouped:
        group = group.copy()
        diff = group['DIFF'].iloc[0]  # DIFF value is the same for all rows of an ITEM NUMBER
        remaining_diff = diff

        # Sort by LEVEL_PRIORITY and aggregate LICENSE PLATE numbers
        licence_plate_counts = group['LICENCE PLATE'].value_counts()
        locations_with_single_licence = licence_plate_counts[licence_plate_counts > 1].index.tolist()

        # Assign replen quantities based on LEVEL and LICENCE PLATE priorities
        for index, row in group.iterrows():
            if remaining_diff <= 0:
                result.append((index, 0))  # No replen needed if already fulfilled
                continue

            posted_quantity = row['POSTED QUANTITY']
            level_priority = row['LEVEL_PRIORITY']
            licence_plate = row['LICENCE PLATE']

            # Case 1: If the item is in a location with a shared LICENCE PLATE, prioritize it
            if licence_plate in locations_with_single_licence:
                replen_qty = min(posted_quantity, remaining_diff)
                result.append((index, replen_qty))
                remaining_diff -= replen_qty
                locations_with_single_licence.remove(licence_plate)  # Use only once
                continue

            # Case 2: Otherwise, use the location sequentially based on LEVEL
            replen_qty = min(posted_quantity, remaining_diff)
            result.append((index, replen_qty))
            remaining_diff -= replen_qty

        # Case 3: For rows not contributing to replen, set STOCK FOR REPLEN to 0
        for index, row in group.iterrows():
            if not any(idx == index for idx, _ in result):
                result.append((index, 0))

    # Update the dataframe with the calculated STOCK FOR REPLEN values
    for index, stock_value in result:
        df.at[index, 'STOCK FOR REPLEN'] = stock_value

    return df

# Step 4: Add the calculated column
output5_df['STOCK FOR REPLEN'] = 0  # Initialize the column with 0
output5_df = calculate_stock_for_replen(output5_df)

# Save the final dataframe as OUTPUT6A
if not os.path.exists(os.path.dirname(output6a_path)):
    os.makedirs(os.path.dirname(output6a_path))  # Create the output folder if it doesn't exist

output5_df.to_excel(output6a_path, index=False)
print(f"OUTPUT6A file saved at: {output6a_path}")
