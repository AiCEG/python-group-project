import pandas as pd
import os

# Set absolute path for working directory or else it won't find relative path to raw_data and transformed data
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# In- and output path
input_path = "./raw_data/su-d-15.02.04.01.xlsx"
output_path = "./transformed_data/T2.2 Studierende nach Hochschule (Seit 1990).xlsx"

# read data from excel 
raw_data = pd.read_excel(
    input_path, 
    sheet_name="T2.2-T2.3", 
    usecols="A:K", 
    skiprows=2, 
    nrows=19,
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
transposed_data = transposed_data.iloc[1:]

# initialize list for column names
column_names = list(transposed_data.columns)

# set column names
column_names = [
"Jahr", "total", "uni_basel", "uni_bern", "uni_freiburg", "uni_genf",
"iheid", "uni_lausanne", "uni_luzern", "uni_neuenburg", "uni_st_gallen",
"uni_ph_st_gallen", "uni_zurich", "uni_dell_svizzera_italiana",
"uni_fern_schweiz", "uni_institut_kurt_boesch", "eth_lausanne", "eth_zurich"
]

# set column names in df
transposed_data.columns = column_names

# interpret jahr as str so we can substr. the first 4 "letters"
transposed_data["Jahr"] = transposed_data["Jahr"].str[:4]

# cast back to int, because its a number (jahr)
transposed_data["Jahr"] = transposed_data["Jahr"].astype(int)

# export dataframe as an excel
transposed_data.to_excel(
    output_path,  # Pfad zur Ausgabedatei
    sheet_name="Tabelle1",  # Name des Sheets
    index=False  # Index weglassen
)
