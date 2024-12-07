import pandas as pd
import re
import os

# Set absolute path for working directory or else it won't find relative path to raw_data and transformed data
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# Function to remove numbering from column title for example: 1.1 Theologie -> Theologie
def remove_numbering(column_name):
    return re.sub(r'^\d+(\.\d+)*\s*', '', column_name)

# In- and output path
input_path = "./raw_data/su-d-15.02.04.01.xlsx"
output_path = "./transformed_data/T10 Studierende nach Studientstufe und MINT Fach seit 2014.xlsx"

# read data from excel 
raw_data = pd.read_excel(
    input_path, 
    sheet_name="T10", 
    usecols="A:L", 
    skiprows=3, 
    nrows=228,
    index_col=None
)

# drop all na values (caused by formatting of excel)
raw_data.dropna(axis=0, how='all', inplace=True)  # Zeilen mit NaN entfernen
raw_data.dropna(axis=1, how='all', inplace=True)  # Spalten mit NaN entfernen

# Get drop condition for % of auslaender and frauen
drop_condition = raw_data["Unnamed: 1"].isin(["% Frau", "% Ausland"])

# Drop the data 
raw_data = raw_data[~drop_condition]

# Drop NaN to get a better overview
raw_data.dropna(axis=1, how='all', inplace=True)  # Spalten mit NaN entfernen

#
raw_data = raw_data.iloc[1:]

# transpose data
transposed_data = raw_data.T

# jahr is read as an index, we want it as a normal column
transposed_data = transposed_data.reset_index()

# Get Titles for columns
header_row = transposed_data.iloc[0]

# Set titles for columns
transposed_data.columns = header_row

# override first column to "jahr"
transposed_data.columns.values[0] = "jahr"

# Add Additional Information to header
for idx, col in enumerate(transposed_data.columns):
    if idx >= 8 and idx < 16:
        f"{col} Bachelor"
    elif idx >= 16 and idx < 24:
        f"{col} Master"
    elif idx >= 24 and idx < 32:
        f"{col} Diplom"
    elif idx >= 32 and idx < 40:
        f"{col} Doktorat"
    elif idx >= 40 and idx < 48:
        f"{col} Weiterbildung"
    else:
        col

transposed_data = transposed_data.iloc[1:]
transposed_data.dropna(axis=1, how='all', inplace=True)

transposed_data.columns = [
    f"{col} Bachelor" if idx >= 9 and idx <= 16 else 
    f"{col} Master" if idx >= 17 and idx <= 24 else
    f"{col} Diplom" if idx >= 25 and idx <= 32 else
    f"{col} Doktorat" if idx >= 33 and idx <= 40 else
    f"{col} Weiterbildung" if idx >= 41 and idx <= 48 else
    col
    for idx, col in enumerate(transposed_data.columns)
]

# interpret jahr as str so we can substr. the first 4 "letters"
transposed_data["jahr"] = transposed_data["jahr"].str[:4]

# cast back to int, because its a number (jahr)
transposed_data["jahr"] = transposed_data["jahr"].astype(int)

# export dataframe as an excel
transposed_data.to_excel(
    output_path,  # Pfad zur Ausgabedatei
    sheet_name="Tabelle1",  # Name des Sheets
    index=False  # Index weglassen
)
