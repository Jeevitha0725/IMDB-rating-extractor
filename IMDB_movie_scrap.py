import pandas as pd
from bs4 import BeautifulSoup
import requests, openpyxl

excel=openpyxl.Workbook()
print(excel.sheetnames)
sheet=excel.active
sheet.title="Top rated Movies 250"
print(excel.sheetnames)
sheet.append(["Movie Rank","Movie Name","Movie Year","Movie Rank"])


try:
    r=requests.get("https://www.imdb.com/chart/top/")
    #print(r.raise_for_status())
    print(r)
    soup=BeautifulSoup(r.content,"html.parser")
    #print(soup.prettify())
    items=soup.find("tbody",class_="lister-list").find_all("tr")
    print(len(items))

    for item in items:
        movie_name=item.find("td",class_="titleColumn").a.text
        movie_rank=item.find("td",class_="titleColumn").get_text(strip=True).split(".")[0]
        movie_year=item.find("td",class_="titleColumn").span.text.strip("()")
        movie_rating=item.find("td", class_="ratingColumn imdbRating").strong.text
        print(movie_rank,movie_name,movie_year,movie_rating)
        sheet.append([movie_rank,movie_name,movie_year,movie_rating])

except Exception as e:
    print(e)

excel.save(r"C:\Users\sjeev\Documents\Python\Web_scraping\IMDB Movie Rating.xlsx")
