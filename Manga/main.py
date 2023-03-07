import openpyxl as px
import requests as r
import datetime
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
from TITLES import bcolors

def printC(title, name, message, color=None):
    # Adding spaces so that the string is 9 chars long to make it look better
    for i in range(11 - len(title)):
        title += ' '

    if not color == None:
        print(f"     > {title} '{name}': {color}{message}{bcolors.ENDC}")
    else:
        print(f"     > {title} '{name}': {message}")

def get_manga(name):
    print("Loading '{}'...".format(name))

    # Format the search name
    formatet_name = name.lower().replace(" ", "+")

    # Searching the reqestet managa on 'www.mangaguide.de' and setting up BeautifulSoup
    mangaguide_website = r.get("https://www.mangaguide.de/index.php?include=24&suche=" + formatet_name)
    soup = BeautifulSoup(mangaguide_website.content, "html.parser")

    """Acces to the Website"""
    # Get the first search result link
    try:
        first_result_link = soup.find(id="inhalt").find_all("a")[0]['href']
        printC("Manga", name, "Successfully loaded the Manga from the website!", bcolors.OKGREEN)
    except IndexError as e:
        printC("Manga", name, "Failed to find the Manga on the website!", bcolors.FAIL)
        return None

    manga_link = "https://www.mangaguide.de/" + first_result_link
    manga_page = r.get(manga_link)
    manga_data = BeautifulSoup(manga_page.content, "html.parser").find(id="inhalt")

    """Title"""
    # Getting basic Information [Title, Author, Max_Count]
    tr_tags = manga_data.find("table").find_all("tr")

    # The title is always the first tr tag
    manga_title = tr_tags[0].find("td").text
    printC("Title", name, "Successfully loaded the Manga title!", bcolors.OKGREEN)

    """Author"""
    # Searching for the a tag with the 'mangaka_id' in the 'href' to get the author
    author = None
    a_tags = manga_data.find_all("a")
    for a in a_tags:
        if a.has_attr('href'):
            if "mangaka_id=" in a['href']:
                author = a.text
                printC("Author", name, "Successfully loaded the Manga author!", bcolors.OKGREEN)
                break
    
    # Checking if the author has been found
    if author == None:
        printC("Author", name, "Failed to load the Manga author!", bcolors.FAIL)
        author = "Error"

    """Max Count"""
    try:
        # Getting the text where the maximal count of a Manga is stored
        max_count_text = manga_page.text.split("nglich erschien")[1]
        # Getting from the text the number
        max_count = int(re.findall(r'\d+', max_count_text)[0])
        printC("MaxCount", name, "Successfully loaded the Manga max count!", bcolors.OKGREEN)
    except Exception as e:
        printC("MaxCount", name, "Failed to load the Manga max count!", bcolors.FAIL)
        max_count = -1

    """Genre"""
    # Searching for the a tag with the 'kategorie=' in the 'href' to get the genre
    genre = None
    for a in a_tags:
        if a.has_attr('href'):
            if "kategorie=" in a['href']:
                genre = a['href'].split("kategorie=")[1]
                printC("Genre", name, "Successfully loaded the Manga genre!", bcolors.OKGREEN)
                break

    # Checking if the genre has been found
    if genre == None:
        printC("Genre", name, "Failed to load the Manga genre!", bcolors.FAIL)
        genre = "Error"

    """German Count"""
    try:
        # Getting the text where the german count is stored
        german_count_text = manga_page.text.split("auf Deutsch erschienen.")[0][-20:]
        # Getting from the text the numbers
        temp = re.findall(r'\d+', german_count_text)
        # Checking if there is actually a number or just 'ein'
        if len(temp) == 0:
            german_count = 1
        else:
            german_count = int(temp[0])

        printC("GermanCount", name, "Successfully loaded the Manga german count!", bcolors.OKGREEN)
    except Exception as e:
        printC("GermanCount", name, "Failed to load the Manga max count!", bcolors.FAIL)
        german_count = -1

    """Cost"""
    # Going throgh all of the volumes of the Manga to get a cost if the fist few dont have one given
    for i in manga_data.find_all("td", {"class": "bandtext"}):
        try:
            cost_text = i.text.split("Kaufpreis: ")[1]
            cost_nums = re.findall(r'\d+', cost_text)
            cost = int(cost_nums[0]) + (int(cost_nums[1]) / 100)
            printC("Cost", name, "Successfully loaded the Manga cost!", bcolors.OKGREEN)
            break
        except Exception as e:
            cost = -1
    
    # If it dosnt find any price in all the volumes
    if cost == -1:
        printC("Cost", name, "Failed to load the Manga cost!", bcolors.FAIL)
    
    """Cover"""
    # Getting the cover-link beging with the 2nd char to not get the '.' at the beginning
    cover_link = manga_data.find("td", {"class": "cover"}).find("a")["href"][1:]
    cover = "https://www.mangaguide.de" + cover_link

    """Next realease"""
    state = "-"
    state_date = "-"
    # If the Manga is not finshied yet in german
    if not german_count == max_count:
        # Calculating to wich page of the website we have to go to get the information (only 10 volumes per page)
        page = int((german_count - 1) / 10) + 1

        # Setting up the page of the website we need
        manga_website = r.get(manga_link + "&seite=" + str(page))
        soup = BeautifulSoup(manga_website.content, "html.parser")

        # Getting acces to alle the volumes
        table = soup.find_all("table", {"class": "mitte"})[1]
        mangas = table.find_all("tr")
        
        # Getting the next new Manga going to be released dependent on the page
        temp = german_count - int(german_count / 10) * 10 + 1
        # Selecting the Manga and skipping every 2nd tr tag, because the are only white lines
        selected = mangas[temp * 2 - 1 - 1]

        state_date = "-"
        # If the selected Manga has been announced
        if selected.find("span")["class"][0] == "angekuendigt":
            # Getting the isbn of the announced manga
            isbn = selected.text.split("ISBN ")[1][:18]
            # Getting the realse date from thalia.de
            state = get_acces_with_isbn(isbn)[13:]
            state_date = state
            # If a '.' is in the date, we know that it is a real date and i worked
            if "." in state:
                # Convert the date to the format for excel
                splitet = state.split(".")
                state = convert_date_to_excel_ordinal(int(splitet[0]), int(splitet[1]), int(splitet[2]))
                printC("Next", name, "Successfully loaded the Manga next release!", bcolors.OKGREEN)
            else:
                printC("Next", name, "Failed to get the next release from thalia!", bcolors.FAIL)
                state = "-"
        else:
            printC("Next", name, "Successfully loaded the Manga next release!", bcolors.OKGREEN)
            state = "NaN"
            state_date = "Noch nicht Angekündigt"

    return {"name": manga_title, "author": author, "max_count": max_count, "german_count": german_count, "genre": genre, "cost": cost, "cover": cover, "state": state, "state_date": state_date}

# Converts a date to a number for excel
def convert_date_to_excel_ordinal(day, month, year) :
    offset = 693594
    current = datetime.datetime(year, month, day)
    n = current.toordinal()
    return (n - offset)

# Returns release date from given isbn
def get_acces_with_isbn(isbn):
    # Search the isbn on thalia
    manga_website = "https://www.thalia.de/suche?sq=ISBN%20" + str(isbn)

    # Options to not show logs
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    driver = webdriver.Chrome(options=options)

    driver.get(manga_website)

    soup = BeautifulSoup(driver.page_source, "html.parser")

    # Return the availiblity of the reqested isbn
    return soup.find("a", {"class": "element-link-toplevel tm-produkt-link"}).find("dl-product")["product-avail"]

def add_to_excel_file(excel_file, manga_have_count, manga_data):
    wb = px.load_workbook(excel_file)
    sheet = wb.active

    count = manga_have_count
    data = manga_data

    # Starting from the next empty row beginning at row 4
    for i, row in enumerate(sheet['B']):
        if i < 4:
            continue

        if row.value is None:
            cur = str(i + 1)
            break

    # Setting the cell style
    fill = PatternFill("solid", fgColor="D9E1F2")
    thin = Side(border_style="thin", color="000000")
    border = Border(right=thin, left=thin, top=thin, bottom=thin)
    aline = Alignment(horizontal="center", vertical="center")

    # Loading the cover into the 'A' columne
    http = urllib3.PoolManager()
    req = http.request('GET', data["cover"])
    image_file = io.BytesIO(req.data)
    img = Image(image_file)

    img.anchor = "A" + cur
    img.width = 96
    img.height = 134
    sheet.add_image(img, "A" + cur)

    # Loading the name into the 'B' columne
    name_font = Font(name="Calibri", size=14, bold=True)
    name_aline = Alignment(horizontal="left", vertical="center")
    name_cell = "B" + cur

    sheet[name_cell].font = name_font
    sheet[name_cell].alignment = name_aline
    sheet[name_cell].fill = fill
    sheet[name_cell].border = border
    sheet[name_cell] = data["name"]

    # Loading the genre into the 'C' columne
    genre_font = Font(name="Calibri", size=16, bold=True)
    genre_cell = "C" + cur

    sheet[genre_cell].font = genre_font
    sheet[genre_cell].alignment = aline
    sheet[genre_cell].fill = fill
    sheet[genre_cell].border = border
    sheet[genre_cell] = data["genre"]

    # Loading the author into the 'D' columne
    author_cell = "D" + cur

    sheet[author_cell].font = name_font
    sheet[author_cell].alignment = aline
    sheet[author_cell].fill = fill
    sheet[author_cell].border = border
    sheet[author_cell] = data["author"]

    # Loading the have count into the 'E' columne
    count_font = Font(name="Calibri", size=20, bold=True)
    count_cell = "E" + cur

    sheet[count_cell].font = count_font
    sheet[count_cell].alignment = aline
    sheet[count_cell].fill = fill
    sheet[count_cell].border = border
    sheet[count_cell] = count

    # Loading the german- and max count into the 'F' columne
    counts_cell = "F" + cur

    sheet[counts_cell].font = count_font
    sheet[counts_cell].alignment = aline
    sheet[counts_cell].fill = fill
    sheet[counts_cell].border = border
    if data["max_count"] == data["german_count"]:
        sheet[counts_cell] = data["max_count"]
    else:
        sheet[counts_cell] = str(data["german_count"]) + "/" + str(data["max_count"])

    # Loading the next release date into the 'G' columne
    next_cell = "G" + cur

    sheet[next_cell].font = count_font
    sheet[next_cell].alignment = aline
    sheet[next_cell].fill = fill
    sheet[next_cell].border = border
    sheet[next_cell] = data["state"]

    # Loading the cost into the 'H' columne
    cost_cell = "H" + cur

    sheet[cost_cell].font = count_font
    sheet[cost_cell].alignment = aline
    sheet[cost_cell].fill = fill
    sheet[cost_cell].border = border
    sheet[cost_cell] = data["cost"]
    sheet[cost_cell].number_format = "0.00€"

    wb.save(excel_file)

if __name__ == "__main__":
    get_manga("Bj Alex")