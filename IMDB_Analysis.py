#!/usr/bin/env python
# coding: utf-8

# In[25]:


import requests, openpyxl
from bs4 import BeautifulSoup


# In[31]:


excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Rank', 'Movie Name', 'Year Of Release', 'IMDB Rating'])


try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find('tbody', class_ = 'lister-list').find_all('tr')

    for movie in movies:
        rank = movie.find('td', class_ = 'titleColumn').get_text(strip = True).split('.')[0]
        name = movie.find('td', class_ = 'titleColumn').find('a').text
        year = movie.find('td', class_ = 'titleColumn').span.text.strip('()')
        rating = movie.find('td', class_ = 'ratingColumn imdbRating').strong.text
        
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])
    

except Exception as e:
    print(e)
    
    
excel.save('IMDB_Movie_Ratings.xlsx')    


# In[ ]:




