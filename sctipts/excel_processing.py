import pandas as pd

# Load the Excel file
input_file = 'C:\\Users\\lepro\\IdeaProjects\\FIleFormatConverter\\input_data\\Formatova_pravidla_02_test.xlsx'
csv_output_file = 'output.xlsx'

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(input_file)

number_dict = {
    '2': 'fmt/477',
    '3': 'fmt/645,fmt/10,fmt/13',
    '4': 'fmt/4,fmt/649,fmt/640,fmt/199',
    '5': 'fmt/134,fmt/198,fmt/141',
    '6': 'fmt/101',
    'rozbalit (komponenty dle formátu)': df['PUID']
}

exception_dict = {
    'ponechat': df['PUID'],
    'výstupní': df['PUID'],
    'individuální': df['PUID'],
}

bool_dict = {
    'ano': 'true',
    '': 'false'
}


def replace_puid(input_str, row):
    parts = input_str.split()  # Split by commas
    modified_parts = []

    for part in parts:
        # Extract the number (e.g., '2' from '§ 23 odst. 2')
        word = part.strip().split()[-1]
        # Replace the number with the corresponding value from replacement_dict
        if word in number_dict:
            modified_parts.append(number_dict[word])
        if word in exception_dict:
            modified_parts.append(number_dict[word])
            row['Verze'] = True

    # Concatenate the modified parts with commas
    result = ','.join(modified_parts)

    return result


def replace_bool(input_str):
    if input_str in number_dict:
        return number_dict[input_str]


# Change values in column 'A' based on previous values
df['Výstupní formát'] = df['Výstupní formát'].apply(replace_puid, df['Verze'])
df['L'] = df['L'].apply(replace_bool)

# Remove unnecessary columns
# df = df.drop(columns=['A', 'D', 'F', 'I', 'J', 'K', 'M'])

# Save the DataFrame as a CSV file
df.to_csv(csv_output_file, index=False)

print("Output file has been converted to CSV and saved as", csv_output_file)
