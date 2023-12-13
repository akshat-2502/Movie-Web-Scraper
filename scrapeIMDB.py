from bs4 import BeautifulSoup
import requests
from urllib.request import urlopen
import openpyxl




excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Trending Movies'

sheet.append(['Movie Name' , 'IMDB Rating' , 'Available'])


try:
    source = requests.get('https://www.91mobiles.com/entertainment/trending-movies')
    source.raise_for_status()

    soup = BeautifulSoup(source.text,'html.parser')
    
    movies = soup.find('div', class_="result_items").find_all('div', class_="pro_item card_box bg-inner-c txt-white")
    
    for movie in movies:

        name = movie.find('div', class_="target_link ga_tracking").a.text

        rating = movie.find('span', class_="t-b-700").text

        stream = movie.find('div', class_="target_link_ext").text
        

        print(name)
        print(rating)
        print(stream)
        sheet.append([name, rating, stream])
        



except Exception as e:
    print(e)

excel.save('Trending Movies.xlsx')
