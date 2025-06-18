import pandas as pd

# Load the entire Excel workbook
excel_path = "BS-FYP repository.xlsx"
sheets = pd.read_excel(excel_path, sheet_name=None)  # Load all sheets

# Loop through each sheet and save it as a new Excel file
for sheet_name, df in sheets.items():
    # Create a clean filename
    safe_name = sheet_name.replace(" ", "_").replace("/", "_")
    output_path = f"{safe_name}.xlsx"
    
    # Save to a new Excel file
    df.to_excel(output_path, index=False)
    print(f"Saved: {output_path}")