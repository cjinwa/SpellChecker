import pandas as pd
from spellchecker import SpellChecker

# Load the Excel file
file_path = r"C:\Users\ACER\PycharmProjects\PythonProject3\investment_management_data - Copy.xlsx"  # Change to your actual file path
df = pd.read_excel(file_path,sheet_name=1)

# Initialize the spell checker
spell = SpellChecker()

# Define columns (Update these with actual column names)
id_column = "Entity_ID"  # Column with unique identifiers
description_column = "Asset_Description"  # Column with text to check

# List to store results
misspelled_entries = []

# Loop through each row
for index, row in df.iterrows():
    text = row[description_column]  # Ensure it's a string
    words = text.split()  # Split into words


    misspelled_words = [word for word in words if word.lower() not in spell]

    if misspelled_words:
        misspelled_entries.append({
            "ID": row[id_column],
            "Original Text": text,
            "Misspelled Words": ", ".join(misspelled_words)
        })

# Create a new DataFrame with misspelled words
misspelled_df = pd.DataFrame(misspelled_entries)

# Save to a new sheet in the same Excel file
with pd.ExcelWriter(file_path, mode="a", engine="openpyxl") as writer:
    misspelled_df.to_excel(writer, sheet_name="Misspelled_Words_test", index=False)

print("Misspelled words have been written to the new sheet: 'Misspelled_Words'.")
