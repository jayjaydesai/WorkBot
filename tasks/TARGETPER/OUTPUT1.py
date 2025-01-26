import os
import pandas as pd

# Define dynamic paths
base_path = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", "CAPLOCATION", "Deployment", "bulk_report_webapp"))
output_path = os.path.join(base_path, "output", "TARGETPER")

# File paths
target_modified_file = os.path.join(output_path, "TARGET_MODIFIED.csv")
output1_file = os.path.join(output_path, "OUTPUT1.csv")

# Ensure the source file exists
if not os.path.exists(target_modified_file):
    raise FileNotFoundError(f"TARGET_MODIFIED.csv not found in {output_path}")

# Load the TARGET_MODIFIED.csv
target_df = pd.read_csv(target_modified_file)

# Add the specified columns and leave them blank
columns_to_add = [
    "2023-customer-sales",
    "2024-customer-sales",
    "number-of-customer-target",
    "number-of-part-number-target",
    "sales-after-target",
    "total-sales-after-first-target"
]

for col in columns_to_add:
    target_df[col] = ""  # Add blank columns

# Save the updated DataFrame to OUTPUT1.csv
target_df.to_csv(output1_file, index=False)

print(f"OUTPUT1.csv generated successfully at: {output1_file}")
