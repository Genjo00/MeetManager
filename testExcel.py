import pandas as pd

df = pd.read_excel("C:\\Users\\Gianluca Figini\\Documents\\Meeting_Chiasso\\Iscrizioni1.xlsx", sheet_name = "Foglio1")
print(df.shape[0])