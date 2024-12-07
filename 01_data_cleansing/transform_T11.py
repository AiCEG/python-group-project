import pandas as pd
import os

# Set absolute path for working directory or else it won't find relative path to raw_data and transformed data
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# In- and output path
input_path = "./raw_data/su-d-15.02.04.01.xlsx"
output_path = "./transformed_data/T11 Eintritte auf Stufe Diplom und Bachelor nach MINT-Fach (Seit 2014).xlsx"

# read data from excel 
raw_data = pd.read_excel(
    input_path, 
    sheet_name="T11", 
    usecols="A:L", 
    skiprows=3, 
    nrows=35,
    index_col=None
)

# drop all na values (caused by formatting of excel)
raw_data.dropna(axis=0, how='all', inplace=True)  # Zeilen mit NaN entfernen
raw_data.dropna(axis=1, how='all', inplace=True)  # Spalten mit NaN entfernen

print(raw_data)

# Get drop condition for % of auslaender and frauen
drop_condition = raw_data["Unnamed: 1"].isin(["% Frau", "% Ausland"])

# Drop the data 
raw_data = raw_data[~drop_condition]

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

# read from second line to last line - 1 
transposed_data = transposed_data.iloc[2:]

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