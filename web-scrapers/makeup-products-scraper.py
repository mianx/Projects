from bs4 import BeautifulSoup
import requests
import re
import os
import pandas as pd

url = 'https://www.makeupcityshop.com/collections/make-up-palette'
response = requests.get(url)
soup = BeautifulSoup(response.content)
prodData = []
products = soup.find_all('a',{'class':'product-title'})
for i in products:
    url = 'https://www.makeupcityshop.com/' + i['href']
    r = requests.get(url)
    soup = BeautifulSoup(r.content)
    #Image Details
    image = soup.find_all('img')
    name = image[4]['alt']
    link = 'http:' + image[4]['src']
    
    #Product Details
    dct = {'Title' : (i.text.strip()),
    'Category' : soup.find('div',{"class": 'product-type'}).find('span').text.strip(),
    'SKU' : soup.find('div',{"class": 'sku-product'}).find('span').text.strip(),
    'Price' : soup.find('div',{"class": 'prices'}).find('span').text.strip().split('.')[1],
    'Description' : soup.find('div',{"class": 'short-description'}).text.strip()
               }
    prodData.append(dct)
    
    #Getting Image
    with open(name.replace(' ','-').replace('/','-') + '.jpg','wb') as f:
        im = requests.get(link)
        f.write(im.content)

dataframe = pd.DataFrame(prodData)
df = dataframe.iloc[::2]
df.set_index(df.columns[0],inplace= True)
txt = 'http://abdi.dreamsglow.xyz/wp-content/uploads/2022/04/'
imgs = os.listdir('..//Pallete')
lst = [txt + imgs[i] for i in range(0,len(imgs))]
df['images'] = lst

df.to_csv('Product-Pallete.csv',sep=";")


