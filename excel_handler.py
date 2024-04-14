import pandas as pd
def search_excel(filename, sheet_name, search_column, search_value):
    """
    Search for a specific value in a specific column of an Excel file.

    Args:
        filename (str): Path to the Excel file.
        sheet_name (str): Name of the sheet in the Excel file.
        search_column (str): Name of the column to search.
        search_value: Value to search for.

    Returns:
        tuple: A tuple containing client name, file path, file name, and a message indicating whether the value was matched.
    """
    # Read Excel file into a DataFrame
    df = pd.read_excel(filename, sheet_name=sheet_name)

    # Check if search column exists
    if search_column not in df.columns:
        raise ValueError(f"Column '{search_column}' not found in the Excel file.")

    # Perform the search in the specified column
    # mask = df[search_column] == search_value
    mask = df[search_column].str.contains(search_value, case=False, na=False, regex=True)


    if not mask.any():
        # If not found in either column, return None for all values
        return None, None, None, f"No match found for '{search_value}' in '{search_column}'."

    # Get the row where the search value is found
    row = df.loc[mask].iloc[0]

    # Extract client name, file path, and file name
    client_name = row['Client Name']
    file_path = row['File Path']
    file_name = row['File Name']

    return client_name, file_path, file_name, f"Matched value in '{search_column}': {search_value}"
