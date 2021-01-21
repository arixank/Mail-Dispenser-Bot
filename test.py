import pandas as pd

data = pd.read_excel("data.xlsx", engine="openpyxl")

name_list = [name["Name"] for(index, name) in data.iterrows()]

for each in name_list:
    print(f"Hey {each},Hello")
