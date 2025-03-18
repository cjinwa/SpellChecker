import pandas as pd
import win32com.client  # Microsoft Word's spell checker
from openpyxl import load_workbook

# 📌 Load Excel file
file_path = r"C:\Users\ACER\PycharmProjects\PythonProject3\investment_management_data - Copy.xlsx"  # Change to your actual file path

# Read first sheet
df = pd.read_excel(file_path, sheet_name=1)

# Define columns (update as needed)
id_column = "Entity_ID"  # Column with unique identifiers
description_column = "Asset_Description"

# Open Microsoft Word for spell checking
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # Keep Word hidden

# List to store misspelled entries
misspelled_entries = []

# Function to check spelling using MS Word
def check_spelling(text):
    doc = word.Documents.Add()
    doc.Content.Text = text
    doc.CheckSpelling()
    corrected_text = doc.Content.Text.strip()
    doc.Close(False)  # Close Word document without saving
    return corrected_text

# Iterate through asset descriptions
for index, row in df.iterrows():
    asset_text = str(row[description_column]).strip()
    corrected_text = check_spelling(asset_text)

    # If the text is changed, it means there was a spelling mistake
    if corrected_text != asset_text:
        misspelled_entries.append({
            "ID": row[id_column],
            "Asset Description": asset_text,
            "Corrected Version": corrected_text
        })

# Close Word application
word.Quit()

# Convert results to a DataFrame
misspelled_df = pd.DataFrame(misspelled_entries)

# Save results to a second sheet
with pd.ExcelWriter(file_path, mode="a", engine="openpyxl") as writer:
    misspelled_df.to_excel(writer, sheet_name="Misspelled_Words_2", index=False)

print("✅ Spell check completed! Results saved in 'Misspelled_Words' sheet.")