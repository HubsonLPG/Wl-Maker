import pandas as pd
import glob
import os
import time

def save_to_csv():
    base_name = "WL"
    ext = ".csv"
    filename = base_name + ext
    i = 1

    while os.path.exists(filename):
        filename = f"{base_name}_{i}{ext}"
        i += 1

    df.to_csv(filename, sep=';', index=False)
    return filename


class color:
   PURPLE = '\033[95m'
   CYAN = '\033[96m'
   DARKCYAN = '\033[36m'
   BLUE = '\033[94m'
   GREEN = '\033[92m'
   YELLOW = '\033[93m'
   RED = '\033[91m'
   BOLD = '\033[1m'
   UNDERLINE = '\033[4m'
   END = '\033[0m'

print(f"{color.GREEN}{color.BOLD}\nWitaj kolego, słuchaj się grzecznie poleceń bo inaczej zepsujesz a po co!\n{color.END}")
time.sleep(1)

while True:
    path = glob.glob(os.path.join(os.getcwd(), "*.xlsx"))
    list_index = 0
    for item in path:
        print(f'index: {list_index} path: "{item}"')
        list_index += 1

    ind = int(input('Wybierz indeks ścieżki:'))
    path = path[ind]
    df = pd.read_excel(path)

#logika
    menu_1 = input(f"\nCo chcesz zrobić słodki książe? Jeśli tylko wygenerować CSV z gotowego pliku XLSX wybierz 'q'. Jeśli chcesz uzupełnić zestawienie wyciągnięte z dziekanatu, wybierz 'r'. Jeśli chcesz zamknąć kliknij enter.\n")
    if menu_1.lower() == 'q':
        df = df.loc[:, ['Imię', 'Nazwisko', 'Company', 'e-mail', 'Start Date', 'End Date', 'Timezone ID', 'Trainer']]
        print('zapisuje do csv')
        save_to_csv()

    elif menu_1.lower() == 'r':
        start_date = input(f"{color.RED}{color.BOLD}WAŻNE!{color.END} Datę podawaj tylko w formacie yy-mm-dd\npodaj datę startową: ")
        end_date = input("podaj datę końcową: ")
        teacher_name = input("podaj imię wykładowcy: ")
        teacher_surname = input("podaj nazwisko wykładowcy: ")
        teacher_mail = input("podaj mail wykładowcy: ")

        df['Company'] = 'AWSB'
        df['Start Date'] = start_date
        df['End Date'] = end_date
        df['Timezone ID'] = 54
        df['Trainer'] = False
        df = df.loc[:, ['Imię', 'Nazwisko', 'Company', 'e-mail', 'Start Date', 'End Date', 'Timezone ID', 'Trainer']]
        df.loc[len(df)] = {
            'Imię': teacher_name,
            'Nazwisko': teacher_surname,
            'Company': 'AWSB',
            'e-mail': teacher_mail,
            'Start Date': start_date,
            'End Date': end_date,
            'Timezone ID': 54,
            'Trainer': True
            }
        input(f'{df.head()}\n\n{df.tail()}')
        save_to_csv()
    else:
        break
    action = input(f'\nNo i wariacie co robimy?\nJeśli chcesz ponownie skorzystać wybierz "q". Jeśli chcesz zakończyć program wybierz "r"')
    if action.lower() == 'r':
        break
    elif action.lower() == 'q':
        continue






