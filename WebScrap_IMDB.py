from bs4 import BeautifulSoup
import requests, openpyxl 
# request module to get the website
# openpyxl for creating and saving the file to exel

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top 250 Movies'
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'Rating'])

try:
    source = requests.get('https://www.imdb.com/chart/top/?ref_=nv_mv_250')
    source.raise_for_status() #To throw error if website link is invalid

    soup = BeautifulSoup(source.text, 'html.parser')
    #Getting all the id tages of movies
    movies  = soup.find('tbody', class_="lister-list").find_all('tr')
    
    for movie in movies:

       # scaping required details from 'td' tag in html 
        name = movie.find('td', class_="titleColumn").a.text
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])
      

except Exception as e:
    print(e)

excel.save('IMDB Top 250 Movies.xlsx')