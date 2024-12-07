import pandas as pd
import os

# Set absolute path for working directory or else it won't find relative path to raw_data and transformed data
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# In- and output path
input_path = "./raw_data/su-d-15.02.04.01.xlsx"
output_path = "./transformed_data/T2.1 Studierende nach Studienstufe, Geschlecht und Stattsangehörigkeit (Seit 1990).xlsx"

# read data from excel 
raw_data = pd.read_excel(
    input_path, 
    sheet_name="T2.1", 
    usecols="B:M", 
    skiprows=3, 
    nrows=29,
    index_col=None
)

# drop all na values (caused by formatting of excel)
raw_data.dropna(how="all", inplace=True)

# transpose data
transposed_data = raw_data.T

# read from second line to last line - 1 
transposed_data = transposed_data.iloc[2:].iloc[:-1]

# jahr is read as an index, we want it as a normal column
transposed_data = transposed_data.reset_index()

# initialize list for column names
column_names = list(transposed_data.columns)

# set column names
column_names = [
    "jahr", "total", "total_frau", "total_auslaender", 
    "bachelor_total", "bachelor_frau", "bachelor_auslaender",
    "master_total", "master_frau", "master_auslaender",
    "diplom_total", "diplom_frau", "diplom_auslaender",
    "doktorat_total", "doktorat_frau", "doktorat_auslaender",
    "weiterbildung_total", "weiterbildung_frau", "weiterbildung_auslaender",
    "übrige_total", "übrige_frau", "übrige_auslaender"
]

# set column names in df
transposed_data.columns = column_names

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