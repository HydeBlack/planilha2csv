import os
import pandas as pd

def excel_to_csv(input_file):
    # Load the Excel file
    excel_file = pd.ExcelFile(input_file)

    # Get the folder path of the input file
    folder_path = os.path.dirname(input_file)

    # Iterate over each sheet in the Excel file
    for sheet_name in excel_file.sheet_names:
        # Read the sheet into a DataFrame
        df = pd.read_excel(excel_file, sheet_name)
        
        # Save the DataFrame to a CSV file in the same folder
        csv_file = os.path.join(folder_path, f"{os.path.splitext(os.path.basename(input_file))[0]}_{sheet_name}.csv")
        df.to_csv(csv_file, index=False)

        print(f"Saved {sheet_name} from {input_file} as {csv_file}")

if __name__ == "__main__":
    # Get the path of the script
    script_path = os.path.dirname(os.path.abspath(__file__))

    # Find any Excel files in the folder
    for file_name in os.listdir(script_path):
        if file_name.endswith(('.xlsx', '.xls', '.ods', '.xlsm', '.xlsb', 'xlsm')):
            input_file = os.path.join(script_path, file_name)

            # Convert Excel to CSV
            excel_to_csv(input_file)

    print("Conversion completed.")

print("2024-02-22 - by Eder Castro - ver 1.0")