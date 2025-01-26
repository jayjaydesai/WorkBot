import os
import pandas as pd

# Define dynamic paths
base_path = os.getenv(
    "BASE_UPLOAD_PATH",
    os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", "CAPLOCATION", "Deployment", "bulk_report_webapp", "output", "TARGETPER")
)
output3_file = os.path.join(base_path, "OUTPUT3.csv")
output4_file = os.path.join(base_path, "OUTPUT4.xlsx")  # Using .xlsx for multiple sheets

# Ensure the input file exists
if not os.path.exists(output3_file):
    raise FileNotFoundError(f"OUTPUT3.csv not found at {output3_file}")

# Load the main sheet
df = pd.read_csv(output3_file)

# Standardize targetdate column
def validate_and_transform_date(date_str):
    """Validate and transform date to dd-mm-yyyy format."""
    try:
        # First attempt to parse as dd-mm-yyyy
        return pd.to_datetime(date_str, format='%d-%m-%Y', errors='raise').strftime('%d-%m-%Y')
    except ValueError:
        try:
            # Attempt parsing as yyyy-mm-dd and convert to dd-mm-yyyy
            return pd.to_datetime(date_str, format='%Y-%m-%d', errors='raise').strftime('%d-%m-%Y')
        except ValueError:
            # If both fail, return as invalid
            return None

# Apply date validation and transformation
df['targetdate'] = df['targetdate'].astype(str).str.strip()  # Remove any leading/trailing spaces
df['targetdate'] = df['targetdate'].apply(validate_and_transform_date)

# Handle invalid dates
invalid_dates = df['targetdate'].isna().sum()
if invalid_dates > 0:
    print(f"Found {invalid_dates} invalid dates in 'targetdate' column.")
    invalid_rows = df[df['targetdate'].isna()]
    invalid_rows_file = os.path.join(base_path, "INVALID_TARGETDATE.csv")
    invalid_rows.to_csv(invalid_rows_file, index=False)
    print(f"Saved rows with invalid 'targetdate' to: {invalid_rows_file}")

# Drop rows with truly invalid dates
df = df.dropna(subset=['targetdate'])

# Define quarters
quarters = {
    'q1': (pd.Timestamp('2024-01-01'), pd.Timestamp('2024-03-31')),
    'q2': (pd.Timestamp('2024-04-01'), pd.Timestamp('2024-06-30')),
    'q3': (pd.Timestamp('2024-07-01'), pd.Timestamp('2024-09-30')),
    'q4': (pd.Timestamp('2024-10-01'), pd.Timestamp('2024-12-31'))
}

# Create the summary table
summary = pd.DataFrame()

# Get unique customercode values and sort them in ascending order
summary['customercode'] = sorted(df['customercode'].drop_duplicates())

# Initialize quarterly counts
for q, (start, end) in quarters.items():
    summary[q] = summary['customercode'].apply(
        lambda code: df[
            (df['customercode'] == code) &
            (pd.to_datetime(df['targetdate'], format='%d-%m-%Y') >= start) &
            (pd.to_datetime(df['targetdate'], format='%d-%m-%Y') <= end)
        ].shape[0]
    )

# Calculate total
summary['total'] = summary[['q1', 'q2', 'q3', 'q4']].sum(axis=1)

# Calculate status counts
status_counts = df.groupby(['customercode', 'status']).size().unstack(fill_value=0)
summary['target_accepted'] = summary['customercode'].map(status_counts.get('Target Accepted', 0))
summary['best_offered'] = summary['customercode'].map(status_counts.get('Best Offered', 0))
summary['no_discount'] = summary['customercode'].map(status_counts.get('No Discount', 0))

# Save the main and summary sheets in a single Excel file
with pd.ExcelWriter(output4_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='main', index=False)
    summary.to_excel(writer, sheet_name='summary', index=False)

print(f"OUTPUT4.xlsx with main and summary sheets has been generated at: {output4_file}")

