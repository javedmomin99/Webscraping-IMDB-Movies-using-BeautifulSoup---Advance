import requests, openpyxl
from bs4 import BeautifulSoup
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top Rated Movies"
sheet.append(["rank", "name", "year", "runtime", "category", "star_rating", "description"])


try:
  source = requests.get("https://www.imdb.com/list/ls055386972/")
  source.raise_for_status()

  soup = BeautifulSoup(source.text,"html.parser")
  movies = soup.find_all('div',class_="lister-item-content")

  for movie in movies:
    name = movie.find('a').text
    rank = movie.find('span',class_="lister-item-index unbold text-primary").text
    year = movie.find('span',class_="lister-item-year text-muted unbold").text.strip("()")
    runtime = movie.find('span',class_="runtime").text
    category = movie.find('span',class_="genre").get_text(strip=True)
    star_rating = movie.find('span',class_="ipl-rating-star__rating").text
    description = movie.find('p', class_="").get_text(strip=True)
    print(rank, name, year, runtime, category, star_rating, description)
    sheet.append([rank, name, year, runtime, category, star_rating, description])

except Exception as e:
  print(e)    

excel.save('IMDB Ratings.xlsx')
