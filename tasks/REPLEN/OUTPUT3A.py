import pandas as pd

def rename_columns_in_excel(input_file, output_file):
    try:
        # Load the Excel file
        df = pd.read_excel(input_file)
        
        # Rename the columns
        df.rename(columns={
            'RSUBRANGE': 'SUB RANGE',
            'RMIN': 'MIN',
            'RCOVERAGE': 'COVERAGE'
        }, inplace=True)
        
        # Save the modified DataFrame to a new Excel file
        df.to_excel(output_file, index=False)
        print(f"File has been saved successfully as {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Specify the path to your input and output files
input_file = r'C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT3.xlsx'
output_file = r'C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT3A.xlsx'

# Call the function to rename columns and save the new file
rename_columns_in_excel(input_file, output_file)
