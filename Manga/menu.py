from TITLES import *
from main import get_manga

from os import system
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager

def get_choice(bound):
    while True:
        choice = input("> ")
        try:
            choice = int(choice)
        except ValueError:
            print("Du must eine Zahl eingeben!")
            continue

        if choice not in range(1, bound + 1):
            print("Du must eine Zahl zwischen 1 und", bound, "eingeben!")
            continue
        return choice

clear = lambda: system("cls")
delay = lambda: sleep(0.1)
printC = lambda mes, color: print(color, mes, bcolors.ENDC)

states = {0: "Menu", 3: "Infos"}
state = states[0]
while True:
    clear()
    # Menu
    if state == states[0]:
        printC(MENU_TITLE, bcolors.OKCYAN)
        printC("Willkommen in dem Manga-Maneger!!! Hier kannst du deine Mangas in Excel-Tabellen speichern und neue Listen erstellen! Viel Spaß :D\n\n", bcolors.WARNING)
        printC("[1] Manga zu einer exestierenden Liste hinzufügen.", bcolors.WARNING)
        printC("[2] Neue Liste erstellen.", bcolors.WARNING)
        printC("[3] Informationen zu einem Manga ausgeben.", bcolors.WARNING)
        printC("[4] Bedinungsanweisung.\n", bcolors.WARNING)

        state = states[get_choice(4)]
    # Infos
    elif state == states[3]:
        printC(INFOS_TITLE, bcolors.OKCYAN)
        printC("Hier im Info-Menü kannst du über einen Manga alle Möglichen Informationen ganz einfach bekommen! ('Enter' um wieder ins Menü zu kommen)\n\n", bcolors.WARNING)
        name = input("Gib den namen eines Mangas > ")

        if name == "":
            state = states[0]
            continue

        print("\n------------------------------------------------------------------------------------------------------------------------")
        data = get_manga(name)
        print("------------------------------------------------------------------------------------------------------------------------\n")

        if data == None:
            printC("[ERROR] Es gab Fehler beim Laden der Informationen des Mangas! Bitte Melden sie den Fehler beim Creator.\n", bcolors.FAIL)
        else:
            print(f"Informationen über '{data['name']}':")
            print(f"   Titel: {data['name']}")
            print(f"   Author: {data['author']}")
            if not data['german_count'] == data['max_count']:
                print(f"   Anzahl: {data['german_count']}/{data['max_count']}")
            else:
                print(f"   Anzahl: {data['max_count']}")
            print(f"   Genre: {data['genre']}")
            print(f"   Preis: {data['cost']}€")
            print(f"   Cover Link: {data['cover']}")
            print(f"   Nächster Release: {data['state_date']}")

        input("Drücke 'Enter' um wieder ins Menü zu kommen...")
        state = states[0]

