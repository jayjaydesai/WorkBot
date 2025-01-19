import pandas as pd

# File paths
output_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT"
output1_file = f"{output_path}\OUTPUT1.xlsx"
output0a_file = f"{output_path}\OUTPUT0A.xlsx"
output1a_file = f"{output_path}\OUTPUT1A.xlsx"

# Load OUTPUT1 and OUTPUT0A into DataFrames
output1_df = pd.read_excel(output1_file)
output0a_df = pd.read_excel(output0a_file)

# Add CATEGORY STATUS column
output1_df["CATEGORY STATUS"] = "GENERAL"
output0a_df["CATEGORY STATUS"] = "EXPORT"

# Concatenate the two DataFrames
merged_df = pd.concat([output1_df, output0a_df], ignore_index=True)

# Save the merged DataFrame to OUTPUT1A.xlsx
merged_df.to_excel(output1a_file, index=False)

print(f"OUTPUT1A.xlsx has been created successfully at {output1a_file}")
