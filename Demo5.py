
import requests, webbrowser
from bs4 import BeautifulSoup

user_input= input("Enter something to search: ")
print("googling")
process={1:"idea", 2:"knowledge"  }
google_search=requests.get("https://www.google.co.in/search?q="+user_input)
htmlContent=google_search.content
soup= BeautifulSoup(htmlContent,'html.parser')

anchors=soup.find_all('a')

#Get href of anchor tag



file=open('file1.txt','w')
for link in anchors:
 a=link.get('href')
 substring=user_input
 if(link.get('href')[0:4]=='/url' and substring in a ) :
    #print(link.get('href'))
    file.write(a)
    file.write('\n')
#file.close()
file=open('file1.txt',"r")
print(file.read())

file.close()