import requests, webbrowser
from bs4 import BeautifulSoup
import openpyxl
import json

wb=openpyxl.load_workbook("C:\\Users\\PRAJWAL\\PycharmProjects\\WebScapping\\excel\\Audit.xlsx")
sh1=wb['Sheet1']
startUps=[]
for i in range (2,sh1.max_row+1):
    if (sh1.cell(row=i, column=2).value == None):
        break
    startUps.append(sh1.cell(row=i,column=2).value)

myjsonfile=open("C:\\Users\\PRAJWAL\\PycharmProjects\\WebScapping\\excel\\Passion_Process.json","r")
jsondata=myjsonfile.read()

obj=json.loads(jsondata)

li=obj['Process_areas']
process_areas=[]
for i in range(len(li)):
    process_areas.append(li[i].get("Process"))
#print(process_areas)

rows=15;
col=3
for list in startUps:
 for value in process_areas:
    print(list+" "+ value + " " + "Management")
    google_search = requests.get("https://www.google.co.in/search?q=" + list+" "+ value )
    htmlContent = google_search.content
    soup = BeautifulSoup(htmlContent, 'html.parser')
    anchors = soup.find_all('a')
    for link in anchors:
        a = link.get('href')
        if (link.get('href')[0:4] == '/url'):
            print(link.get('href'))
            c=sh1.cell(row=rows,column=col)
            c.value=a
            col=col+1
        rows=rows+2

    print("-----------------------------------------------------------------------------------\n")

wb.save("C:\\Users\\PRAJWAL\\PycharmProjects\\WebScapping\\excel\\Audit.xlsx")

