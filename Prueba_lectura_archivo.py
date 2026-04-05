import pandas as pd
import os


# Guarda la ruta de la carpeta actual
ruta_actual = os.getcwd() 

print(f"Mi script está corriendo en: {ruta_actual}")

columnas = ['start_date', 'date_from', 'date_until', 'concept', 'bp']
df = pd.read_excel('ArchivosCruce/Programador.xlsm', usecols=columnas)

#df['date_from'] = df['start_date'].str.strip()
df['date_from'] = df['date_from'].str.strip()
df['date_until'] = df['date_from'].str.strip()

df['start_date'] =pd.to_datetime(df['start_date'])
df['date_from'] = pd.to_datetime(df['date_from'], format='%H:%M').dt.time
df['date_until'] = pd.to_datetime(df['date_until'], format='%H:%M').dt.time
#df = pd.read_excel('ArchivosCruce/Programador.xlsm')

#print(df.head(9))
#print(df.info())

#print(df.columns.tolist())
#print(df.describe())
#print(df.shape)
