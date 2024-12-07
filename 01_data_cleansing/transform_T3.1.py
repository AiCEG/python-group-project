import pandas as pd
import os

# Set absolute path for working directory or else it won't find relative path to raw_data and transformed_data
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# In- and output path
input_path = "./raw_data/su-d-15.02.04.01.xlsx"
output_path = "./transformed_data/T3.1 Studierende nach Fachbereichsgruppe, Geschlecht und Staatsangehörigkeit (Seit 1990).xlsx"

# read data from excel 
raw_data = pd.read_excel(
    input_path, 
    sheet_name="T3.1-T3.2", 
    usecols="A:K", 
    skiprows=3, 
    nrows=32,
    index_col=None
)

# drop all na values (caused by formatting of excel)
raw_data.dropna(axis=0, how='all', inplace=True)  # Zeilen mit NaN entfernen
raw_data.dropna(axis=1, how='all', inplace=True)  # Spalten mit NaN entfernen

# transpose data
transposed_data = raw_data.T

# jahr is read as an index, we want it as a normal column
transposed_data = transposed_data.reset_index()

# read from second line to last line - 1 
transposed_data = transposed_data.iloc[2:]

# initialize list for column names
column_names = list(transposed_data.columns)

# set column names
column_names = [
"jahr", "total", "total_frau_prozent", "total_auslaender_prozent",
"geist_und_sozial_total", "geist_und_sozial_frau_prozent", "geist_und_sozial_auslaender_prozent",
"wirtschaft_total", "wirtschaft_frau_prozent", "wirtschaft_auslaender_prozent",
"recht_total", "recht_frau_prozent", "recht_auslaender_prozent",
"natur_total", "natur_frau_prozent", "natur_auslaender_prozent",
"medizin_total", "medizin_frau_prozent", "medizin_auslaender_prozent",
"technik_total", "technik_frau_prozent", "technik_auslaender_prozent", "interdiszi_total",
"interdiszi_frau_prozent", "interdiszi_auslaender_prozent"
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
