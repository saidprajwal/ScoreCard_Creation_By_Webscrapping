import requests
from bs4 import BeautifulSoup
url="https://codewithharry.com"
r=requests.get(url)
htmlContent=r.content
soup=BeautifulSoup(htmlContent,'html.parser')
print(soup.prettify())

# print div by id
s1=soup.find(id='__next')
#print(s1)

#get div children
#.contents- A tag children are available as list    //stores in memory
#.children- A tag content are available as a generator //fast

print(s1.children)
content=s1.contents
for element in content:
    print(element)

for item in s1.strings:            #stripped_string
    print(item)

print(s1.parent)
print("----------------------------------")
#print(s1.parents)  return object
for item in s1.parents:
    print(item.name)

#.next_siblings is the next tag after tag. It also considers next space as sibling
#.precious_sibling


#css selector - # is used for id
# class selector - . is used
ele=soup.select('#__next')
print(ele)
