from color import color
import pandas as pd
import glob
import os
import time

gr = color.GREEN
bd = color.BOLD
eol = color.END
cn = color.DARKCYAN


def generate_csv(df):
    df_copy = df.copy()
    df_copy = df_copy.loc[
        :,
        [
            "Imię",
            "Nazwisko",
            "Company",
            "e-mail",
            "Start Date",
            "End Date",
            "Timezone ID",
            "Trainer",
        ],
    ]
    return df_copy


def fill_tally_into_csv(df):
    df_copy = df.copy()
    teacher_mail = str(input(
        f"Wybierz typ maila wykładowcy\n{gr}{bd}'1'{eol} - {(df_copy['Prowadzący zajęcia, imię'][0]).lower()[0]}{df_copy['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl\n"
        f"{gr}{bd}'2'{eol} - {(df_copy['Prowadzący zajęcia, imię'][0].lower())}.{df_copy['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl\n"
        f"{gr}{bd}Jeśli mail jest niestandardowany, uzupełnij pole: {eol}\n"
    ))
    time.sleep(0.3)
    if teacher_mail.lower() == "1":
        teacher_mail = f"{(df_copy['Prowadzący zajęcia, imię'][0]).lower()[0]}{df_copy['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl"
    elif teacher_mail.lower() == "2":
        teacher_mail = f"{(df_copy['Prowadzący zajęcia, imię'][0].lower())}.{df_copy['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl"

    df_copy.loc[len(df_copy)] = {
        "Imię": df_copy["Prowadzący zajęcia, imię"][0],
        "Nazwisko": df_copy["Prowadzący zajęcia, nazwisko"][0],
        "e-mail": teacher_mail,
    }

    df_copy["Company"] = "AWSB"
    df_copy["Start Date"] = pd.Timestamp.now().normalize()
    df_copy["End Date"] = df_copy["Start Date"] + pd.Timedelta(days=14)
    df_copy["Timezone ID"] = 54
    df_copy["Trainer"] = False
    df_copy = df_copy.loc[
        :,
        [
            "Imię",
            "Nazwisko",
            "Company",
            "e-mail",
            "Start Date",
            "End Date",
            "Timezone ID",
            "Trainer",
        ],
    ]
    df_copy.loc[df_copy.index[-1], "Trainer"] = True
    return df_copy


def merge_tallys_into_csv(df):
    df_copy = df.copy()
    print(
        f"{cn}{bd}Wybierz pliki do dołączenia\n"
        f"enter - zakończenie wybierania{eol}\n"
    )
    while True:
        path = glob.glob(os.path.join(os.getcwd(), "zestawienia", "*.xlsx"))
        for i, item in enumerate(path):
            print(f'index: {i} plik: "{os.path.basename(item)}"')
        ind = input("\nWybierz indeks ścieżki:\n")
        time.sleep(0.3)
        if ind != "":
            path = path[int(ind)]
            df_to_add = pd.read_excel(path)
            df_copy = pd.concat([df_copy, df_to_add], ignore_index=True)
            continue
        else:
            break
    return df_copy
