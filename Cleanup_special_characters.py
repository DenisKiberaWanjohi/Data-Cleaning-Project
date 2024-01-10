import os
import re
import pandas as pd

def find_special_characters(row):
    """
    Find special characters in the row.

    Args:
        row (pd.Series): A row of the DataFrame.

    Returns:
        str: Special characters found in the row along with the column names.

    """
    special_chars = []
    for column, value in row.items():
        if isinstance(value, str):
            column_chars = re.findall(r'[!@#$%^&*()_+{}\[\]:;"/?.<,`~\-=\|]', value)
            if column_chars:
                special_chars.append(f"{column}: {', '.join(column_chars)}")
    return ', '.join(special_chars)

def clean_special_characters(text, keep_chars=None):
    """
    Clean up system-defined unacceptable special characters in text,
    while allowing customization to keep certain special characters.

    Args:
        text (str): The input text to clean.
        keep_chars (str or None): Special characters to keep in the text.
            If None, all special characters will be removed. Default is None.

    Returns:
        str: The cleaned text.

    """
    # Define the set of system-defined unacceptable special characters
    unacceptable_chars = r'/'

    # Remove unacceptable special characters from the text
    cleaned_text = re.sub(unacceptable_chars, '', text)

    # Keep specified special characters if provided
    if keep_chars is not None:
        cleaned_text = re.sub(f'[^{keep_chars}]', '', cleaned_text)

    return cleaned_text

# Read the Excel file
input_file = r"C:\UTest File.xlsx"
excel_file = pd.ExcelFile(input_file)

# Extract the input file name without extension
input_file_name = os.path.splitext(os.path.basename(input_file))[0]

# Specify the sheets and columns you want to clean, or set them to None to clean all
sheets_to_clean = None  # Add the names of sheets you want to clean, or set to None
columns_to_clean = None  # Add the names of columns you want to clean, or set to None

# Create an empty list to store the summary data
summary_data = []

# Create a writer to save all data to one Excel file with the input file name as a prefix
output_file = r"C:\SpecialCharactersCleaned_{}.xlsx".format(input_file_name)
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    for sheet_name in excel_file.sheet_names:
        # Read the sheet data into a DataFrame
        df = excel_file.parse(sheet_name)

        if sheets_to_clean is None or sheet_name in sheets_to_clean:
            # Clean the specified columns and get summary for each column
            cleaned_columns = []
            for column in df.columns:
                if columns_to_clean is None or column in columns_to_clean:
                    df[column] = df[column].apply(lambda x: clean_special_characters(str(x)))
                    cleaned_columns.append(column)

            # Add a new column to show special characters in each row
            #df['SpecialCharacters'] = df.apply(find_special_characters, axis=1)

            # Save the cleaned data for the sheet without the 'SpecialCharacters' column
            df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep='')

            # Collect summary data for each column in the current sheet
            for column in cleaned_columns:
                column_chars_set = set(re.findall(r'[!@#$%^&*()_+{}\[\]:;"/?.<,`~\-=\|]', ' '.join(df[column].dropna())))
                column_chars = ', '.join(column_chars_set)
                column_summary = {
                    'SheetName': sheet_name,
                    'Column': column,
                    'SpecialCharacters': column_chars
                }
                summary_data.append(column_summary)
        else:
            # Save the unchanged data for sheets that are not being cleaned
            df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep='')

    # Create a DataFrame for the summary data
    summary_df = pd.DataFrame(summary_data)

    # Save the summary to the same Excel file
    summary_df.to_excel(writer, sheet_name='Summary', index=False, na_rep='')

print("All cleaned data and summary saved to", output_file)