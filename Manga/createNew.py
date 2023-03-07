import openpyxl as px
import requests as r
import datetime
import shutil
import re

from openpyxl.styles import *
from bs4 import BeautifulSoup
from selenium import webdriver 
from selenium.webdriver.chrome.service import Service as ChromeService 
from webdriver_manager.chrome import ChromeDriverManager
import PIL
import io
import urllib3
from openpyxl.drawing.image import Image
# from openpyxl_image_loader import SheetImageLoader

def main_def(manga_name, manga_have_count, excel_file):
    def get_manga(name):
        # Format search
        name = name.lower().replace(" ", "+")

        # Get acces to website
        manga_website = r.get("https://www.mangaguide.de/index.php?include=24&suche=" + name)
        soup = BeautifulSoup(manga_website.content, "html.parser")

        # Get the first search result link
        manga_link = "https://www.mangaguide.de/" + soup.find(id="inhalt").find_all("a")[0]['href']
        manga_page = r.get(manga_link)
        manga_data = BeautifulSoup(manga_page.content, "html.parser").find(id="inhalt")
        
        # Getting Title, Author and Max Count
        temp = manga_data.find("table").find_all("tr")

        title = temp[0].find("td").text
        author = manga_data.find_all("a")
        for i in author:
            try:
                check = i["href"]
            except:
                pass
            if "index.php?include=6&mangaka_id=" in check:
                author = i.text
                break

        max_count = int(re.findall(r'\d+', manga_page.text.split("nglich erschien")[1])[0])

        # Getting Genre
        genre = manga_data.find_all("a")
        for i in genre:
            try:
                check = i["href"]
            except:
                pass
            if "index.php?include=5&suchen=1&kategorie=" in check:
                genre = check.split("index.php?include=5&suchen=1&kategorie=")[1]
                break
        
        # German Count
        temp = re.findall(r'\d+', manga_page.text.split("auf Deutsch erschienen.")[0][-20:])
        if len(temp) == 0:
            german_count = 1
        else:
            german_count = int(temp[0])

        # Cost
        for i in manga_data.find_all("td", {"class": "bandtext"}):
            try:
                cost_str = re.findall(r'\d+', i.text.split("Kaufpreis: ")[1])
                cost = int(cost_str[0]) + (int(cost_str[1]) / 100)
                break
            except Exception as e:
                cost = 0
        
        # Cover
        cover = "https://www.mangaguide.de" + manga_data.find("td", {"class": "cover"}).find("a")["href"][1:]

        # Next realease
        state = "-"
        if not german_count == max_count:
            seite = int((german_count - 1) / 10) + 1

            manga_website = r.get(manga_link + "&seite=" + str(seite))
            soup = BeautifulSoup(manga_website.content, "html.parser")

            table = soup.find_all("table", {"class": "mitte"})[1]
            mangas = table.find_all("tr")
            
            temp = german_count - int(german_count / 10) * 10 + 1
            selected = mangas[temp * 2 - 1 - 1]

            if selected.find("span")["class"][0] == "angekuendigt":
                isbn = selected.text.split("ISBN ")[1][:18]
                state = get_acces_with_isbn(isbn)[13:]
                if "." in state:
                    splitet = state.split(".")
                    state = convert_date_to_excel_ordinal(int(splitet[0]), int(splitet[1]), int(splitet[2]))
                else:
                    state = "-"
            else:
                state = "NaN"

        return {"name": title, "author": author, "max_count": max_count, "german_count": german_count, "genre": genre, "cost": cost, "cover": cover, "state": state}

    def convert_date_to_excel_ordinal(day, month, year) :
        offset = 693594
        current = datetime.datetime(year,month,day)
        n = current.toordinal()
        return (n - offset)

    def get_acces_with_isbn(isbn):
        # Get acces to website
        manga_website = "https://www.thalia.de/suche?sq=ISBN%20" + str(isbn)

        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        driver = webdriver.Chrome(options=options)

        driver.get(manga_website)

        soup = BeautifulSoup(driver.page_source, "html.parser")

        return soup.find("a", {"class": "element-link-toplevel tm-produkt-link"}).find("dl-product")["product-avail"]


    wb = px.load_workbook(excel_file)
    wb.iso_dates = True
    sheet = wb.active

    count = manga_have_count
    data = get_manga(manga_name)

    for i, row in enumerate(sheet['B']):
        if i < 4:
            continue

        if row.value is None:
            cur = str(i + 1)
            break

    fill = PatternFill("solid", fgColor="D9E1F2")
    thin = Side(border_style="thin", color="000000")
    border = Border(right=thin, left=thin, top=thin, bottom=thin)
    aline = Alignment(horizontal="center", vertical="center")

    # Cover
    http = urllib3.PoolManager()
    req = http.request('GET', data["cover"])
    image_file = io.BytesIO(req.data)
    img = px.drawing.image.Image(image_file)
    img.anchor = "A" + cur
    img.width = 96
    img.height = 134
    sheet.add_image(img, "A" + cur)

    # Name
    name_font = Font(name="Calibri", size=14, bold=True)
    name_aline = Alignment(horizontal="left", vertical="center")
    name_cell = "B" + cur

    sheet[name_cell].font = name_font
    sheet[name_cell].alignment = name_aline
    sheet[name_cell].fill = fill
    sheet[name_cell].border = border
    sheet[name_cell] = data["name"]

    # Genre
    genre_font = Font(name="Calibri", size=16, bold=True)
    genre_cell = "C" + cur

    sheet[genre_cell].font = genre_font
    sheet[genre_cell].alignment = aline
    sheet[genre_cell].fill = fill
    sheet[genre_cell].border = border
    sheet[genre_cell] = data["genre"]

    # Author
    author_cell = "D" + cur

    sheet[author_cell].font = name_font
    sheet[author_cell].alignment = aline
    sheet[author_cell].fill = fill
    sheet[author_cell].border = border
    sheet[author_cell] = data["author"]

    # Count
    count_font = Font(name="Calibri", size=20, bold=True)
    count_cell = "E" + cur

    sheet[count_cell].font = count_font
    sheet[count_cell].alignment = aline
    sheet[count_cell].fill = fill
    sheet[count_cell].border = border
    sheet[count_cell] = count

    # Counts
    counts_cell = "F" + cur

    sheet[counts_cell].font = count_font
    sheet[counts_cell].alignment = aline
    sheet[counts_cell].fill = fill
    sheet[counts_cell].border = border
    if data["max_count"] == data["german_count"]:
        sheet[counts_cell] = data["max_count"]
    else:
        sheet[counts_cell] = str(data["german_count"]) + "/" + str(data["max_count"])

    # Next
    next_cell = "G" + cur

    sheet[next_cell].font = count_font
    sheet[next_cell].alignment = aline
    sheet[next_cell].fill = fill
    sheet[next_cell].border = border
    sheet[next_cell] = data["state"]
    sheet[next_cell].number_format = "[$-de-DE]D. MMMM YYYY;@"

    # Cost
    cost_cell = "H" + cur

    sheet[cost_cell].font = count_font
    sheet[cost_cell].alignment = aline
    sheet[cost_cell].fill = fill
    sheet[cost_cell].border = border
    sheet[cost_cell] = data["cost"]
    sheet[cost_cell].number_format = "0.00€"

    wb.save(excel_file)

mangas_names = ("Barefoot Angel",
"BJ Alex",
"Black Butler",
"Café Liebe",
"Chainsaw Man",
"Colorful line",
"Cupid is Struck by Lightning",
"Der Metalhead von nebenan",
"Die Rippe des adam",
"Free hugs for you only",
"Given",
"Hang out Crisis",
"I Hear the Sunspot",
"My Genderless Boyfriend",
"Neon Genesis Evangelion - Perfect Edition",
"Never good enough",
"Pheromoneholic",
"To your Eternity",
"Tokyo Ghoul",
"Uns trennen Welten")

manga_counts = (1,5,1,2,5,1,2,1,2,1,7,1,2,1,1,1,2,5,14,1)

shutil.copyfile("C:\\Users\\milan\\Documents\\Manga\\Blank.xlsx", "C:\\Users\\milan\\Documents\\Manga\\NewMangaListe.xlsx")

name = "NewMangaListe.xlsx"

for i in range(len(mangas_names)):
    print("Loading '", mangas_names[i], "'...")
    main_def(mangas_names[i], manga_counts[i], "C:\\Users\\milan\\Documents\\Manga\\" + name)
    print("Succsesfully loaded '", mangas_names[i], "'![", i + 1, "/", len(mangas_names), "]")