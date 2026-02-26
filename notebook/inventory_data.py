#This script processes forestry data from an Excel file, renaming columns, 
# correcting scientific names, changing data types, classifying trees into 
# diameter classes, and dividing the data according to line numbers.

import pandas as pd
import numpy as np
import openpyxl
import os

# Load an Excel file
path = os.path.abspath(__file__)
dir_path = os.path.dirname(path)
df = pd.read_excel(os.path.join(dir_path, '..\data\BASE-YURUMANGUI-2013.xlsx'), sheet_name='base')

# Change the name of the columns
df = df.rename(
    columns={
        "N_Linea": "Line Number",
        "N_Parcela_general": "Plot",
        "N_Subparcela": "Subplot",
        "N_Arbol": "Tree number",
        "Nombre_común": "Tree name",
        "Nombre_cientifico": "Scientific name",
        "GRUPO COMERCIAL":"Commercial group",
        "DAP_(cm)": "DBH [cm]",
        "DAP (m)": "DBH [m]",
        "H_Comercial_(mt)": "Commercial height [m]",
        "Calidad_Fuste": "Stem quality",
        "Observaciones ": "Observations",
        "G (Área Basal)": "Basal area [m2]",
        "Volumen":"Volume [m3]",
    }
)

# translating commercial groups
df["Commercial group"] = df["Commercial group"].replace(
    {
        "FINA": "Fine",
        "ORDINARIA": "Ordinary",
        "PALMA": "Non-Timber Forest Product",
        "SIN USO COMERCIAL": "No commercial use",
    }    
)

# Correscting scientific names
df["Scientific name"] = df["Scientific name"].replace(
    {
        "Brosimun utile": "Brosimum utile",
        
    }    
)

# Changing object data types into nullable integer data types
df = df.astype(
    {
        "Line Number": pd.Int64Dtype(),
        "Plot": pd.Int64Dtype(),
        "Subplot": pd.Int64Dtype(),
        "Tree number": pd.Int64Dtype(),
    }
)

# changing object data types into categorical
df = df.astype(
    {
        "Tree name": pd.CategoricalDtype(ordered=False),
        "Scientific name": pd.CategoricalDtype(ordered=False),
        "Line Number": pd.CategoricalDtype(ordered=True),
        "Plot":pd.CategoricalDtype(ordered=True),
        "Subplot":pd.CategoricalDtype(ordered=True),
        "Tree number":pd.CategoricalDtype(ordered=True),
        "Commercial group":pd.CategoricalDtype(ordered=False),
        "Stem quality": pd.CategoricalDtype(ordered=True),
    }
)

df = df.astype(
    {
        "Line Number": pd.CategoricalDtype(categories=[1, 2, 3, 4, 5, 6, 7], ordered=True),
    }
)

#Data are divided according to line number
line1 = df[df["Line Number"] == 1]
line2 = df[df["Line Number"] == 2]
line3 = df[df["Line Number"] == 3]
line4 = df[df["Line Number"] == 4]
line5 = df[df["Line Number"] == 5]
line6 = df[df["Line Number"] == 6]
line7 = df[df["Line Number"] == 7]

dict_class = {
    "I": "10-20",
    "II": "20-30",
    "III": "30-40",
    "IV": "40-50",
    "V": "50-60",
    "VI": "60-70",
    "VII": "70-80",
    "VIII": "80-90",
    "IX":"90-100",
    "X": "100-110",
    "XI": "110-120",
    "XII": "120-130",
    "XIII": "130-140",
    "XIV": "140-150",
    "XV": "150-160",
    "XVI": "160-170",
    "XVII": "170-180",
    "XVIII": "180-190",
    "XIX": "190-200",
    "XX": "> 200",           
}

# Classify the trees into diameter classes using 10 cm intervals with roman numerals
df["DBH Class"] = np.select(
    [(df["DBH [cm]"] > 10) & (df["DBH [cm]"] <= 20),
     (df["DBH [cm]"] > 20) & (df["DBH [cm]"] <= 30),
     (df["DBH [cm]"] > 30) & (df["DBH [cm]"] <= 40),
     (df["DBH [cm]"] > 40) & (df["DBH [cm]"] <= 50),
     (df["DBH [cm]"] > 50) & (df["DBH [cm]"] <= 60),
     (df["DBH [cm]"] > 60) & (df["DBH [cm]"] <= 70),
     (df["DBH [cm]"] > 70) & (df["DBH [cm]"] <= 80),
     (df["DBH [cm]"] > 80) & (df["DBH [cm]"] <= 90),
     (df["DBH [cm]"] > 90) & (df["DBH [cm]"] <= 100),
     (df["DBH [cm]"] > 100) & (df["DBH [cm]"] <= 110),
     (df["DBH [cm]"] > 110) & (df["DBH [cm]"] <= 120),
     (df["DBH [cm]"] > 120) & (df["DBH [cm]"] <= 130),
     (df["DBH [cm]"] > 130) & (df["DBH [cm]"] <= 140),
     (df["DBH [cm]"] > 140) & (df["DBH [cm]"] <= 150),
     (df["DBH [cm]"] > 150) & (df["DBH [cm]"] <= 160),
     (df["DBH [cm]"] > 160) & (df["DBH [cm]"] <= 170),
     (df["DBH [cm]"] > 170) & (df["DBH [cm]"] <= 180),
     (df["DBH [cm]"] > 180) & (df["DBH [cm]"] <= 190),
     (df["DBH [cm]"] > 190) & (df["DBH [cm]"] <= 200),
     (df["DBH [cm]"] > 200)],
    ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX"],
    default="Other"
)

df = df.astype(
    {
        "DBH Class": pd.CategoricalDtype(categories=["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX"], ordered=True),
        "Line Number": pd.CategoricalDtype(categories=[1, 2, 3, 4, 5, 6, 7], ordered=True),
    }
)

def load_inventory():
    inventory = df.copy()
    return inventory