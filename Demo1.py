import requests
from bs4 import BeautifulSoup
url="https://www.codewithharry.com"
r=requests.get(url)
htmlContent=r.content
soup=BeautifulSoup(htmlContent,'html.parser')
#print(soup.prettify())
title=soup.title



#Commonly used type of objects:
#print(type(title))#1.Tag
#print(type(title.string))#2.Navigable String
#print(type(soup))#3.BeatifulSoup
#4.Comment

#Get all paragraphs from page
paras=soup.find_all('p')
print(paras)

#Get all anchor from page
anchors=soup.find_all('a')
print(anchors)

for link in anchors:
    print(link)

#Get href of anchor tag
for link in anchors:
    print(link.get('href'))
print("------------------------------")
#remove duplicate
all_links=set()
for link in anchors:
   # if(link.get('href')!="#"):                                  //to remove any link
        link="https://www.codewithharry.com"+link.get('href')
        all_links.add(link)
        print(link)


#Get first paragraph from page and its class
#print(soup.find('p')['class'])

#find all the elements with class lead
#print(soup.find_all("p",class_="mt-2"))

#print(soup.find('p').get_text())

#print(soup.get_text())

