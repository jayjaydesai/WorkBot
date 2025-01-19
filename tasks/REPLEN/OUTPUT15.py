import pandas as pd

# Define file paths
output14_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT14.xlsx"
output15_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT15.xlsx"

# Load the OUTPUT14 file
output14_df = pd.read_excel(output14_path)

# Normalize column names
output14_df.columns = output14_df.columns.str.strip().str.upper()

# Step 1: Replace "FINAL STOCK FOR REPLEN" value if it's 0 with "STOCK FOR REPLEN2" value
output14_df['FINAL STOCK FOR REPLEN'] = output14_df.apply(
    lambda row: row['STOCK FOR REPLEN2'] if row['FINAL STOCK FOR REPLEN'] == 0 else row['FINAL STOCK FOR REPLEN'],
    axis=1
)

# Step 2: Add the "NEW AS01011-AVAILABLE PHYSICAL" column
output14_df['NEW AS01011-AVAILABLE PHYSICAL'] = output14_df['AS01011-AVAILABLE PHYSICAL'] + output14_df['FINAL STOCK FOR REPLEN']

# Step 3: Add the "FURTHER REPLEN REQUIRED OR NOT" column
output14_df['FURTHER REPLEN REQUIRED OR NOT'] = output14_df.apply(
    lambda row: 'NO FURTHER REPLEN REQUIRED' if row['MIN'] < row['NEW AS01011-AVAILABLE PHYSICAL'] else 'FURTHER REPLEN REQUIRED',
    axis=1
)

# Save the updated DataFrame as OUTPUT15
try:
    output14_df.to_excel(output15_path, index=False)
    print(f"OUTPUT15 file saved at: {output15_path}")
except Exception as e:
    print(f"An error occurred while saving OUTPUT15: {e}")

