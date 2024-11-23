import pandas as pd
import re

def extract_emails_and_names(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Prepare a list to store the results
    results = []

    # Regular expression for valid email addresses
    email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

    # Process each cell in the DataFrame
    for column in df.columns:
        for entry in df[column]:
            if isinstance(entry, str):  # Ensure the entry is a string
                # Find all email addresses using regex
                emails = re.findall(email_regex, entry)
                
                # Create a dictionary for the current row
                row_data = {column: entry}
                
                for idx, email in enumerate(emails):
                    # Extract the human name from the email
                    name_part = email.split('@')[0]  # Get the part before '@'
                    # Clean the name by replacing dots and underscores with spaces
                    name = re.sub(r'[._]', ' ', name_part).title()  # Capitalize words

                    # Check if the extracted name contains any letters
                    if re.search(r'[a-zA-Z]', name):
                        row_data[f'{column} Email {idx + 1}'] = email
                        row_data[f'{column} Name {idx + 1}'] = name
                    else:
                        row_data[f'{column} Email {idx + 1}'] = email
                        row_data[f'{column} Name {idx + 1}'] = ''

                # Append the row data to the results
                results.append(row_data)
            else:
                # Append the original entry with empty email and name columns for non-strings
                results.append({column: entry})

    # Create a DataFrame from the results list
    results_df = pd.DataFrame(results)

    # Save the results to the same Excel file
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        results_df.to_excel(writer, sheet_name='Email Extraction Results', index=False)

    print(f"Emails and names saved to 'Email Extraction Results' in '{file_path}'.")

# Specify the path to your Excel file
file_path = 'C:\\Users\\ACER\\OneDrive - Pixeltruth\\Desktop\\Phonepe\\Testttt\\Test.xlsx'

# Call the function to extract emails and names, then save the output
extract_emails_and_names(file_path)
