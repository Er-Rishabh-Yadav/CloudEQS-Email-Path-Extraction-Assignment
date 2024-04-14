import os
import pandas as pd
from email_wrapper import EmailWrapper
from excel_handler import search_excel
import re

def convert_filename_to_pattern(filename):
    """
    Convert a filename with a flexible date format into a regex pattern.

    Args:
        filename (str): The filename to convert.

    Returns:
        str: The regex pattern for the filename.
    """
    # Extract the start string and date part from the filename
    match = re.match(r'(\w+)_(\d{4}-\d{2}-\d{2})\.txt', filename)
    if match:
        start_string = match.group(1)
        date_part = match.group(2)
        # Create a regex pattern with the extracted date
        pattern = f'{start_string}_\d{{4}}-\d{{2}}-\d{{2}}\.txt'
        return pattern
    else:
        # Return a generic pattern if the filename doesn't match expected format
        return '\w+_\d{4}-\d{2}-\d{2}\.txt'


def main():
    EMAIL='temp.for.rishabh@gmail.com'
    PASSWORD='jzwupdllmymgevhl'
    N = 1  # Fetch top 1 email

    # Initialize EmailWrapper
    email_wrapper = EmailWrapper(EMAIL, PASSWORD)
    email_wrapper.login()

    # Fetch emails
    emails = email_wrapper.fetch_emails(num_emails=N)

    # Get attachment file paths
    # attachment_paths = email_wrapper.get_attachment_file_paths(emails)

    # # Get body paths
    # email_body = email_wrapper.get_email_body(emails)
    # print(email_body)

    paths = email_wrapper.extract_paths_from_body(emails)
    print(f'paths recieve in emails {paths}')

    # Excel file path
    excel_file = "temp_client_data.xlsx"

    # Create an empty list to store results as dictionaries
    results_list = []

    try:
        for attachment_path in paths:
            filename = os.path.basename(attachment_path)
            search_path = os.path.dirname(attachment_path) + "/"

            print(f'Filename {filename} \n FilePath {search_path} \n')

            filename_pattern = convert_filename_to_pattern(filename)
            print(f'filename pattern {filename_pattern}')
            # Search in Excel for filename
            client_name, file_path, file_name, matched_message = search_excel(excel_file, 'Sheet1', 'File Name', filename_pattern)
            
            if client_name is None:
                # Search in Excel for file path if filename not found
                client_name, file_path, file_name, matched_message = search_excel(excel_file, 'Sheet1', 'File Path', search_path)

            if client_name is None:
                print(f"No matching data found for path: {attachment_path}")
            else:
                print("Result:")
                print(f"Client Name: {client_name}")
                print(f"File Path: {file_path}")
                print(f"File Name: {file_name}")
                print(matched_message)

                # Append result as a dictionary to the list
                results_list.append({
                    'Client Name': client_name,
                    'File Path': file_path,
                    'File Name': file_name,
                    'Matched Value': matched_message,
                    'Attachment Path': attachment_path
                })

        # Convert the list of dictionaries to a DataFrame
        results_df = pd.DataFrame(results_list)

        # Save results to CSV
        results_csv = "search_results.csv"
        results_df.to_csv(results_csv, index=False)
        print(f"Results saved to {results_csv}")

    except Exception as e:
        print("An error occurred:", e)

    email_wrapper.logout()

if __name__ == "__main__":
    main()
