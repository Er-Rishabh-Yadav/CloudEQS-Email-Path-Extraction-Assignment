from flask import Flask, send_file
import os
import pandas as pd
from email_wrapper import EmailWrapper
from excel_handler import search_excel

app = Flask(__name__)

# Initialize EmailWrapper
EMAIL = 'temp.for.rishabh@gmail.com'
PASSWORD = 'jzwupdllmymgevhl'
email_wrapper = EmailWrapper(EMAIL, PASSWORD)
email_wrapper.login()

# Main Script
def run_main():
    attachment_paths = []

    # Get all attachments from the 'attachments' folder
    for root, dirs, files in os.walk("attachments"):
        for file in files:
            attachment_paths.append(os.path.join(root, file))

    # Create an empty list to store results
    results_list = []

    try:
        for attachment_path in attachment_paths:
            # Search for each attachment
            filename = os.path.basename(attachment_path)
            search_path = os.path.dirname(attachment_path) + "/"
            print(f'Filename {filename} \n FilePath {search_path} \n')

            # Search in Excel for filename
            client_name, file_path, file_name, matched_message = search_excel("temp_client_data.xlsx", 'Sheet1', 'File Name', filename)

            if client_name is None:
                # Search in Excel for file path if filename not found
                client_name, file_path, file_name, matched_message = search_excel("temp_client_data.xlsx", 'Sheet1', 'File Path', search_path)

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
                    'Matched Value': matched_message
                })

        # Convert list to CSV
        if results_list:
            results_csv = "search_results.csv"
            with open(results_csv, "w") as f:
                f.write("Client Name,File Path,File Name,Matched Value\n")
                for result in results_list:
                    f.write(f"{result['Client Name']},{result['File Path']},{result['File Name']},{result['Matched Value']}\n")

            return results_csv
        else:
            return None

    except Exception as e:
        error_msg = f"An error occurred: {str(e)}"
        print(error_msg)
        return None

# Run the main script on Flask app startup
results_file = run_main()

@app.route('/download-results', methods=['GET'])
def download_results():
    if results_file is not None:
        return send_file(results_file, as_attachment=True, download_name='search_results.csv'), 200
    else:
        return "No results found.", 404


if __name__ == '__main__':
    app.run(debug=True)
