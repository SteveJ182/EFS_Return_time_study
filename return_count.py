import pandas as pd
#program to count the returns from movehistory, just specify the persons name and location to look for
def extract_and_sum_total_data(file_path, sheet_name="MoveHistory"):  
    # Load the Excel file without headers
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, header=None)

    # Fill down merged cells for columns B (1) and C (2)
    df.iloc[:, [1, 2]] = df.iloc[:, [1, 2]].fillna(method='ffill')

    # Ensure the file has enough columns
    if df.shape[1] < 17:  # Must have at least 17 columns (Q = index 16)
        print("Error: Excel file does not have enough columns.")
        return 0

    # Filter rows where:
    # - Column B (1) OR C (2) contains "Returns 001 - RET001"
    # - Column P (15) contains "Steven Cervera"
    filtered_data = df[
        ((df.iloc[:, 1].str.contains('Returns 001 - RET001', case=False, na=False)) |
         (df.iloc[:, 2].str.contains('Returns 001 - RET001', case=False, na=False))) &
        (df.iloc[:, 15].str.contains('Steven Cervera', case=False, na=False))
    ]

    # Extract and convert column Q (index 16) values to numbers
    extracted_values = pd.to_numeric(filtered_data.iloc[:, 16], errors='coerce').dropna()

    # Sum the values
    total_sum = extracted_values.sum()

    return total_sum

# Example usage
file_path = "MoveHistory.xlsx"  # Update with your file path
total = extract_and_sum_total_data(file_path)
print("Total Sum:", total)
