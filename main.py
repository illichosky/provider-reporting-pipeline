import os, re
import pandas as pd

file_list = os.listdir("input")

writer = pd.ExcelWriter('output.xlsx')
for file in file_list:
    data = pd.read_excel('input/'+file, sheet_name='result')
    print(data.head())
    title = re.search("_REPORT_ (.*?)_", str(file)).group(1)
    data.to_excel(writer,sheet_name=title, index=False)
writer.save()
