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
output_path = "./transformed_data/T7 Studierende auf Stufe Diplom und Bachelor nach Fachrichtung (seit 2014).xlsx"

# read data from excel 
raw_data = pd.read_excel(
    input_path, 
    sheet_name="T7", 
    usecols="A:W", 
    skiprows=2, 
    nrows=113,
    index_col=None
)

# drop all na values (caused by formatting of excel)
raw_data.dropna(axis=0, how='all', inplace=True)  # Zeilen mit NaN entfernen
raw_data.dropna(axis=1, how='all', inplace=True)  # Spalten mit NaN entfernen

# get columns with value 'F' in the first row
columns_to_drop = raw_data.columns[raw_data.iloc[0] == 'F']

# drop those columns
raw_data = raw_data.drop(columns=columns_to_drop)

# transpose data
transposed_data = raw_data.T

# jahr is read as an index, we want it as a normal column
transposed_data = transposed_data.reset_index()

# drop total column, is not needed
transposed_data = transposed_data.drop(transposed_data.columns[1], axis=1)

# set column names by finding value that is not NaN (column name is somewhere in the first three rows, but first value found that is not NaN will be taken)
new_column_names = transposed_data.apply(lambda col: col.dropna().iloc[0])

# Set the new column names
transposed_data.columns = new_column_names

# override first column to "jahr"
transposed_data.columns.values[0] = "jahr"

# Remove column numbering with function above
transposed_data.columns = [remove_numbering(col) for col in transposed_data.columns]

# Remove the first three rows which contained the header titles
transposed_data = transposed_data.iloc[3:].reset_index(drop=True)

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
