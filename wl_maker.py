import pandas as pd
import glob
import os
import time
from color import color
import modules_wl


gr = color.GREEN
bd = color.BOLD
eol = color.END
cn = color.DARKCYAN
lcn = color.CYAN
bl = color.BLUE
pr = color.PURPLE

print(
    f"{gr}{bd}\nWitaj kolego, słuchaj się grzecznie poleceń bo inaczej zepsujesz a po co!\n{eol}"
)
time.sleep(0.5)

while True:
    path = glob.glob(os.path.join(os.getcwd(), "zestawienia", "*.xlsx"))
    for i, item in enumerate(path):
        df = pd.read_excel(item)
        try:
            print(
                f'index: {i} plik: "{os.path.basename(item)}" '
                f"{bd}{cn}\n{df['Tok - nazwa'][0]}{eol} "
                f"{bd}{gr}{df['Prowadzący zajęcia, nazwisko'][0]}{eol} "
                f"{bd}{bl}liczba rekordów: {len(df)}{eol} "
                f"{bd}{pr}{df['Przedmiot'][0]}{eol}\n"
            )
        except Exception as e:
            print(
                f'WYSTĄPIŁ BŁĄD W PLIKU {os.path.basename(item)} {e}'
            )

    ind = int(input("\nWybierz indeks ścieżki:\n"))
    time.sleep(0.3)
    path = path[ind]
    df = pd.read_excel(path)
    df = df.drop_duplicates(subset='e-mail', ignore_index=True)
    file_name = str(
        f"{df['Tok - nazwa'][0]} {df['Grupa - nazwa'][0]} {(df['Prowadzący zajęcia, imię'][0])[0]}{df['Prowadzący zajęcia, nazwisko'][0]}"
    )
    menu_1 = str(input(
        f"\nCo chcesz zrobić słodki książe?"
        f"\n{gr}{bd}'1'{eol} - Scalić wybrany plik z innymi plikami CSV"
        f"\n{gr}{bd}'2'{eol} - uzupełnić zestawienie wyciągnięte z dziekanatu"
        f"\n{gr}{bd}enter{eol} - zamknij program\n"
    ))
    time.sleep(0.3)
# --- MENU option 1 ---
    if menu_1.lower() == "1":
        while True:
            file_name = str(
                f"Łączenie_{(df['Prowadzący zajęcia, imię'][0])[0]}{df['Prowadzący zajęcia, nazwisko'][0]}"
            )
            df_copy = modules_wl.merge_tallys_into_csv(df, ind)
            df_copy = modules_wl.fill_tally_into_csv(df_copy)
            if not modules_wl.summary(df_copy, ind, df):
                continue
            modules_wl.save_to_csv(file_name, df_copy)
            break
# --- MENU option 1 ---

# --- MENU option 2 ---
    elif menu_1.lower() == "2":
        while True:
            df_copy = modules_wl.fill_tally_into_csv(df)
            if not modules_wl.summary(df_copy, ind, df):
                continue
            modules_wl.save_to_csv(file_name, df_copy)
            break
# --- MENU option 2 ---

    else:
        break
    action = str(input(
        f"\nNo i wariacie co robimy?\n"
        f"{gr}{bd}'1'{eol} - Jeśli chcesz ponownie skorzystać\n"
        f"{gr}{bd}'2'{eol} -  Jeśli chcesz zakończyć program\n"
    ))
    time.sleep(0.3)
    if action.lower() == "2":
        break
    elif action.lower() == "1":
        continue
