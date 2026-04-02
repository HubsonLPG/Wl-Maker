import pandas as pd
import glob
import os
import time
from color import color
import modules_wl


def save_to_csv(raw_name, df):
    file_name = raw_name.replace("/", "_")
    os.makedirs("csvki", exist_ok=True)
    save_path = os.path.join("csvki", f"{file_name}.csv")
    df.to_csv(save_path, sep=";", index=False)
    print("\nZapisano plik! ♿")
    return file_name


gr = color.GREEN
bd = color.BOLD
eol = color.END
cn = color.DARKCYAN

print(
    f"{gr}{bd}\nWitaj kolego, słuchaj się grzecznie poleceń bo inaczej zepsujesz a po co!\n{eol}"
)
time.sleep(0.5)

while True:
    path = glob.glob(os.path.join(os.getcwd(), "zestawienia", "*.xlsx"))
    for i, item in enumerate(path):
        print(f'index: {i} plik: "{os.path.basename(item)}"')

    ind = int(input("\nWybierz indeks ścieżki:\n"))
    time.sleep(0.3)
    path = path[ind]
    df = pd.read_excel(path)
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
        file_name = str(
            f"{(df['Prowadzący zajęcia, imię'][0])[0]}{df['Prowadzący zajęcia, nazwisko'][0]}"
        )
        # df_copy = modules_wl_maker.generate_csv(df)
        df_copy = modules_wl.merge_tallys_into_csv(df)
        df_copy = modules_wl.fill_tally_into_csv(df)
        save_to_csv(file_name, df_copy)
# --- MENU option 1 ---

# --- MENU option 2 ---
    elif menu_1.lower() == "2":
        while True:
            df_copy = modules_wl.fill_tally_into_csv(df)
            if (
                str(input(
                    f"Sprawdź podsumowanie bratku czy jest okej."
                    f"\n{gr}{bd}'enter' - okej{eol}"
                    f"\n{gr}{bd}'1' - nie okej{eol}\n"
                    f"\n{cn}{bd}ostatni index:{eol} {ind},{cn}{bd} przedmiot:{eol}"
                    f"{df['Przedmiot'][0]}\n\n{df_copy.tail()}\n"
                ))
                == "2"
            ):
                continue
            else:
                save_to_csv(file_name, df_copy)
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
