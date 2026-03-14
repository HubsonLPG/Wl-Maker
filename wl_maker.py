import pandas as pd
import glob
import os
import time

def save_to_csv(raw_name, df):
    file_name = raw_name.replace("/", "_")
    os.makedirs("csvki", exist_ok=True)
    save_path = os.path.join("csvki", f"{file_name}.csv")
    df.to_csv(save_path, sep=';', index=False)
    print("\nZapisano plik! ♿")
    return file_name

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

gr = color.GREEN
bd = color.BOLD
eol = color.END

print(f"{gr}{bd}\nWitaj kolego, słuchaj się grzecznie poleceń bo inaczej zepsujesz a po co!\n{eol}")
time.sleep(1)

while True:
    path = glob.glob(os.path.join(os.getcwd(), "zestawienia", "*.xlsx"))
    list_index = 0
    for item in path:
        print(f'index: {list_index} path: "{item}"')
        list_index += 1

    ind = int(input('Wybierz indeks ścieżki:\n'))
    path = path[ind]
    df = pd.read_excel(path)
#logika
    file_name = str(f"{df['Tok - nazwa'][0]} {df['Grupa - nazwa'][0]} {(df['Prowadzący zajęcia, imię'][0])[0]}{df['Prowadzący zajęcia, nazwisko'][0]}")
    menu_1 = input(f"\nCo chcesz zrobić słodki książe? Jeśli tylko wygenerować CSV z gotowego pliku XLSX wybierz {gr}{bd}'q'{eol}. Jeśli chcesz uzupełnić zestawienie wyciągnięte z dziekanatu, "
                   f"wybierz {gr}{bd}'r'{eol}. Jeśli chcesz zamknąć kliknij {gr}{bd}enter{eol}.\n")
    if menu_1.lower() == 'q':
        df = df.loc[:, ['Imię', 'Nazwisko', 'Company', 'e-mail', 'Start Date', 'End Date', 'Timezone ID', 'Trainer']]
        print('zapisuje do csv\n')

        save_to_csv(file_name, df)

    elif menu_1.lower() == 'r':
        teacher_mail = input(f"Wybierz typ maila wykładowcy, {gr}{bd}'q'{eol} dla {(df['Prowadzący zajęcia, imię'][0]).lower()[0]}{df['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl, "
                             f"{gr}{bd}'r'{eol} dla {(df['Prowadzący zajęcia, imię'][0].lower())}.{df['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl."
                             f"{gr}{bd} Jeśli mail jest niestandardowany, uzupełnij pole: {eol}\n")
        if teacher_mail.lower() == 'q':
            teacher_mail = f"{(df['Prowadzący zajęcia, imię'][0]).lower()[0]}{df['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl"
        elif teacher_mail.lower() == 'r':
            teacher_mail = f"{(df['Prowadzący zajęcia, imię'][0].lower())}.{df['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl"


        df.loc[len(df)] = {
            'Imię': df['Prowadzący zajęcia, imię'][0],
            'Nazwisko': df['Prowadzący zajęcia, nazwisko'][0],
            'e-mail': teacher_mail,
            }
        

        df['Company'] = 'AWSB'
        df['Start Date'] = pd.Timestamp.now().normalize()
        df['End Date'] = df['Start Date'] + pd.Timedelta(days=14)
        df['Timezone ID'] = 54
        df['Trainer'] = False
        df = df.loc[:, ['Imię', 'Nazwisko', 'Company', 'e-mail', 'Start Date', 'End Date', 'Timezone ID', 'Trainer']]
        df.loc[df.index[-1], 'Trainer'] = True


        input(f'{df.tail()}')
        save_to_csv(file_name, df)
    else:
        break
    action = input(f'\nNo i wariacie co robimy?\nJeśli chcesz ponownie skorzystać wybierz {gr}{bd}"q"{eol}. Jeśli chcesz zakończyć program wybierz {gr}{bd}"r"{eol}\n')
    if action.lower() == 'r':
        break
    elif action.lower() == 'q':
        continue






