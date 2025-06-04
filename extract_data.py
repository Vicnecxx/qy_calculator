# -*- coding: utf-8 -*- 
# Fluorescence data file extraction functions

import pandas as pd
import os
# File extraction function supporting CSV, XLSX and XLS formats
# Set print_params=True to print parameters. Function takes file path as input and returns data table.
def extract_data(file_path, print_params=False):
    # Determine file type and read accordingly
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            df_raw = pd.read_csv(file_path, header=None)
        elif ext == '.xlsx':
            df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
        elif ext == '.xls':
            df_raw = pd.read_excel(file_path, header=None, engine='xlrd')
        else:
            print(f"Error: Unsupported file format {ext}. Only CSV, XLSX and XLS are supported.")
            return None
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

    # Extract instrument parameters
    params = {
        "Data mode:": None, "EX WL:": None, "EM Start WL:": None,
        "EM End WL:": None, "Scan speed:": None, "Delay:": None,
        "EX Slit:": None, "EM Slit:": None, "PMT Voltage:": None,
        "Response:": None, "Corrected spectra:": None, "Shutter control:": None
    }

    # Iterate through rows to find parameters
    data_start_row = None
    for idx, row in df_raw.iterrows():
        if row[0] in params:
            params[row[0]] = row[1]
        # Locate data starting row ("nm" and "Data" headers)
        if row[0] == "nm" and row[1] == "Data":
            data_start_row = idx + 1
            break

    # Check if data starting row was found
    if data_start_row is None:
        print("Error: Failed to locate data starting row.")
        return None

    # Extract data points
    df_data = df_raw.iloc[data_start_row:, [0, 1]]
    df_data.columns = ["nm", "Data"]
    df_data = df_data.dropna(subset=["nm"])  # Remove empty rows

    # Print parameters
    if print_params:
        print("Spectrometer parameters:")
        for key, value in params.items():
            if value is None:
                print(f"Warning: Parameter {key} was not successfully extracted.")
            else:
                print(f"{key}\t{value}")

    return df_data.reset_index(drop=True)

# Test run
if __name__ == "__main__":
    file_path = "D:\\Xuan\\Desktop\\Relative Quantum Yield Calculator\\test.xls"
    df_data = extract_data(file_path, print_params=True)
    if df_data is not None:
        print("\nData table:")
        print(df_data)