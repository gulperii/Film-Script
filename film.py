from bs4 import BeautifulSoup
import requests
import xlsxwriter
from datetime import datetime

pageUrl = "https://boxofficeturkiye.com/vizyon/"
mainPage = requests.get(pageUrl, headers={'User-agent': 'Mozilla/5.0'})
mainPageContent = BeautifulSoup(mainPage.content, "html.parser")
links = []

fileName = "film" + datetime.today().strftime("%Y-%m") + ".xlsx"
workbook = xlsxwriter.Workbook(fileName)
worksheet = workbook.add_worksheet()
rowTitle = 0
colTitle = 0
bold = workbook.add_format({'bold': True})

titles = ['Film Adı', 'Vizyon Tarihi', 'TR Dağıtım', 'Şirket', 'Film Türü', 'Konusu', 'Ülke', 'Yönetmen', 'Oyuncular']

for title in titles:
    worksheet.write(rowTitle, colTitle, title, bold)
    colTitle += 1

for item in mainPageContent.findAll('a', {'class': 'film'}):
    link = pageUrl[:-8] + item['href']
    name = item['title']
    pair = (name, link)
    links.append(pair)

rowFilm = 1
for pair in links:
    colFilm = 0
    (name, link) = pair
    page = requests.get(link, headers={'User-agent': 'Mozilla/5.0'})
    pageContent = BeautifulSoup(page.content, "html.parser")
    items = pageContent.findAll('td', {'class': 'movie-summary-value'})

    vizyonTarihi = items[0].get_text().replace("\n", " ")
    trDagitim = items[1].get_text()
    sirket = items[2].get_text()
    tur = items[3].get_text().replace(" ", "").replace('\r\n', " ")

    try:
        konu = pageContent.findAll('span', {'class': 'spot'}, limit=1)[0].get_text().replace("\n", " ").split("Devamı")[
            0]
    except IndexError:
        konu = 'null'

    ulke = pageContent.find('img', {'class': 'cercevesiyah'}, width=25).get('title')

    cast = pageContent.find('div', {'id': 'movieCast'}).get_text().split("\n")

    mCast = list(filter(lambda x: x != "", cast))

    yonetmenler = []
    oyuncular = []

    actorIndex = 0

    try:
        directorIndex = mCast.index('Yönetmen')
    except ValueError:
        directorIndex = -1
    try:
        actorIndex = mCast.index('Oyuncular')
    except ValueError:
        actorIndex = -1
    try:
        screenwriterIndex = mCast.index('Senaryo')
    except ValueError:
        screenwriterIndex = -1

    if (directorIndex != -1):
        if (actorIndex != -1):
            yonetmenler = mCast[directorIndex + 1:actorIndex]
            if (screenwriterIndex != -1):
                oyuncular = mCast[actorIndex + 1:screenwriterIndex]
            else:
                oyuncular = mCast[actorIndex + 1:]
        elif (screenwriterIndex != -1):
            yonetmenler = mCast[directorIndex + 1:screenwriterIndex]
        else:
            yonetmenler = mCast[directorIndex + 1:]

    allnfo = [name, vizyonTarihi, trDagitim, sirket, tur, konu, ulke, str(yonetmenler)[1:-1], str(oyuncular)[1:-1]]

    for info in allnfo:
        worksheet.write(rowFilm, colFilm, info)
        colFilm += 1
    rowFilm += 1

workbook.close()
