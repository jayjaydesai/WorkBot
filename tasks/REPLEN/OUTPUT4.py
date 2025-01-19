import pandas as pd
import os

# Define file paths
output3_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT3A.xlsx"
output4_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT4.xlsx"

# Load the OUTPUT3A file
output3_df = pd.read_excel(output3_path)

# Ensure column names are normalized
output3_df.columns = output3_df.columns.str.strip().str.upper()

# Check if required columns exist
required_columns = ['AS01011-AVAILABLE PHYSICAL', 'COVERAGE', 'MIN']
for col in required_columns:
    if col not in output3_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT3A.xlsx")

# Add the calculated column "DIFF"
output3_df['DIFF'] = output3_df['COVERAGE'] - output3_df['AS01011-AVAILABLE PHYSICAL']

# Add the calculated column "REPLEN REQUIRED STATUS"
output3_df['REPLEN REQUIRED STATUS'] = output3_df.apply(
    lambda row: "REPLEN REQUIRED" if row['MIN'] > row['AS01011-AVAILABLE PHYSICAL'] else "NO REPLEN REQUIRED",
    axis=1
)

# Save the final dataframe as OUTPUT4
if not os.path.exists(os.path.dirname(output4_path)):
    os.makedirs(os.path.dirname(output4_path))  # Ensure the output directory exists

output3_df.to_excel(output4_path, index=False)
print(f"OUTPUT4 file saved at: {output4_path}")
