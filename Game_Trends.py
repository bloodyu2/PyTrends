# This script was created with the intent to automate the storing of Google Trends Data on Google Drive and Publish it
# on Google Data Studio:

from pytrends.request import TrendReq
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from openpyxl import Workbook

# Import the func
import pandas as pd
import time
import csv

start_time = time.time()
pytrends = TrendReq(hl='en-US', tz=360)

# Define de source of the keywords (columns and #s)
colnames = ["Keywords"]
df = pd.read_csv("Keywords.csv", names=colnames)
df2 = df["Keywords"].values.tolist()
df2.remove("Keywords")

dataset = []

# Define the range of the data
for x in range(0, len(df2)):
    keywords = [df2[x]]
    pytrends.build_payload(
        kw_list=keywords,
        cat=0,
        timeframe='today 5-y',
        geo='')
    data = pytrends.interest_over_time()
    if not data.empty:
        data = data.drop(labels=['isPartial'], axis='columns')
        dataset.append(data)

result = pd.concat(dataset, axis=1)
result.to_csv('search_trends.csv')

# Transform the CSV into an XLSX file for later storage
wb = Workbook()
ws = wb.active
with open('search_trends.csv', 'r') as f:
    for row in csv.reader(f):
        ws.append(row)
wb.save('Resultados_de_pesquisa.xlsx')

# Access Google Drive
gauth = GoogleAuth()
drive = GoogleDrive(gauth)

# Upload the file to the drive
file1 = ['Resultados_de_pesquisa.xlsx']
file1 = drive.CreateFile({'title': 'Resultados_de_pesquisa.xlsx'})
file1.SetContentFile('Resultados_de_pesquisa.xlsx')
file1.Upload()

#End of the Script
executionTime = (time.time() - start_time)
print('Execution time in sec.:' + str(executionTime))
print("Seu arquivo foi carregado com sucesso em seu Drive!")
