import re
import pandas as pd

# Load the Excel file
input_file = '../input_data/input.xlsx'

# Specify the desired file name (including the ".csv" extension)
file_name = 'output.csv'

# Specify the full path to the directory where you want to save the CSV file
directory_path = '../output_data/'

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(input_file)

# Rename columns for clarity
df = df.rename(columns={
    'Verze': 'outputDataFormat',
    'MIMETYPE': 'containerDataFormat',
    'Název': 'name',
    'Kategorie': 'category',
    'Přípona': 'extension',
    'Výstupní formát': 'outPuid',
    'Originál vždy (doporučeno)': 'keepOriginal',
    'Výstupní formát alternativně I': 'possibleOutPuid'
})

# Dictionaries for replacements
number_dict = {
    '2': 'fmt/477',
    '3': 'fmt/645,fmt/10,fmt/13',
    '4': 'fmt/4,fmt/649,fmt/640,fmt/199',
    '5': 'fmt/134,fmt/198,fmt/141',
    '6': 'fmt/101'
}

bold_number_dict = {
    '2': 'fmt/477',
    '3': 'fmt/645',
    '4': 'fmt/4',
    '5': 'fmt/134',
    '6': 'fmt/101'
}

exception_dict = {
    'ponechat': df['PUID'],
    'výstupní': df['PUID'],
    'individuální': df['PUID']
}


# Function to perform replacements
def replace_values(row):
    input_str = row['outPuid']

    if isinstance(input_str, float):
        # Try to convert input_str to float; if successful, return a predefined value
        float(input_str)
        return f"{row['PUID']}###{row['PUID']}###{'false'}###{'false'}"

    else:
        modified_parts_possible_out_puid = []
        modified_parts_out_puid = []
        modified_parts_output = []
        modified_parts_container = []

        # Split the input string into parts
        parts = re.findall(r'\d+|\w+', input_str, re.IGNORECASE)

        for part in parts:
            # Extract the number (e.g., '2' from '§ 23 odst. 2')
            word = part.strip().split()[-1]

            # Replace the number with the corresponding value from dictionaries
            if word in number_dict:
                modified_parts_possible_out_puid.append(number_dict[word])
                if len(modified_parts_container) < 1:
                    modified_parts_out_puid.append(bold_number_dict[word])
                    modified_parts_output.append('false')
                    modified_parts_container.append('false')

            elif word in exception_dict:
                modified_parts_possible_out_puid.append(row['PUID'])
                modified_parts_out_puid.append(row['PUID'])
                modified_parts_output.append('true')
                modified_parts_container.append('false')

            elif word == 'rozbalit':
                modified_parts_possible_out_puid.append(row['PUID'])
                modified_parts_out_puid.append(row['PUID'])
                modified_parts_output.append('false')
                modified_parts_container.append('true')

        # Concatenate the modified parts with commas
        result_out_puid = ','.join(modified_parts_out_puid)
        result_output = ','.join(modified_parts_output)
        result_container = ','.join(modified_parts_container)
        result_possible_out_puid = ','.join(modified_parts_possible_out_puid)

        # Combine values with a delimiter
        combined_result = f"{result_out_puid}###{result_possible_out_puid}###{result_output}###{result_container}"

        return combined_result


# Function to replace boolean values
def replace_bool(input_str):
    if input_str == 'ano':
        return 'true'
    else:
        return 'false'


# Define a function to split the combined string by '###'
def custom_split(combined_str):
    # Split the combined string by '###'
    parts = combined_str.split('###')

    # Ensure that there are at least tree parts
    if len(parts) >= 4:
        return parts[0], parts[1], parts[2], parts[3]
    else:
        return parts[0], None


# Apply the replace function to each row and store the result in a new column
df['Combined'] = df.apply(replace_values, axis=1)

# Split the 'Combined' column into 'outPuid', 'possibleOutPuid', 'outputDataFormat' and 'containerDataFormat' columns
df[['outPuid',
    'possibleOutPuid',
    'outputDataFormat',
    'containerDataFormat']] = df['Combined'].apply(custom_split).apply(pd.Series)

# Change values in column 'keepOriginal' based on previous values
df['keepOriginal'] = df['keepOriginal'].apply(replace_bool)

# Remove unnecessary columns
df = df.drop(columns=['id',
                      'Výstupní formát alternativně II',
                      'Výstupní formát alternativně III',
                      'Komentář',
                      'Combined'])

# Save the DataFrame as a CSV file
df.to_csv(directory_path + file_name, index=False)

print("Output file has been converted to CSV and saved as", file_name)
