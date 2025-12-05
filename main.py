import pandas as pd
from sqlalchemy import create_engine
import urllib

# ================= KONFIGURACJA =================
# Ustawienia bazy danych (MSSQL)
SERVER = 'NAZWA_TWOJEGO_SERWERA'  # np. localhost\SQLEXPRESS lub 192.168.x.x
DATABASE = 'NAZWA_BAZY_DANYCH'
DRIVER = 'ODBC Driver 17 for SQL Server' # Sprawdź w 'ODBC Data Sources', jaki masz sterownik

# Jeśli logujesz się przez Windows Authentication (zazwyczaj w firmach):
TRUSTED_CONNECTION = 'yes'
USERNAME = ''
PASSWORD = ''

# Jeśli logujesz się loginem i hasłem SQL, wypełnij powyższe i ustaw TRUSTED_CONNECTION = 'no'

# Ścieżka do pliku Excel
EXCEL_FILE = r'C:\Sciezka\Do\Twojego\Pliku\dane.xlsx'
SHEET_NAME = 'Arkusz1' # Nazwa zakładki w Excelu

# Nazwa tabeli w bazie
TABLE_NAME = 'produktyfudal'

# ================= MAPOWANIE KOLUMN =================
# Słownik: 'Nazwa Kolumny w Excelu' : 'Nazwa Kolumny w SQL'
# UWAGA: Pomiń kolumnę ID, ponieważ w bazie jest ona IDENTITY (tworzy się sama)
column_mapping = {
    'Kod Produktu': 'KOD_PRODUKTU',
    'Nazwa Produktu': 'NAZWA_PRODUKTU',
    'Indeks 1': 'INDEKS1',
    'Indeks 2': 'INDEKS2',
    'Nazwa Indeksu': 'NAZWA_INDEKSU',
    'Cena': 'CENA_JEDN',
    'Jm': 'JM',
    'Ilość': 'ILOSC',
    'TKW': 'TKW',
    'Opis/Uwagi': 'UWAGI'
}

def main():
    try:
        print("1. Wczytywanie pliku Excel...")
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        
        # Sprawdzenie czy dane zostały wczytane
        print(f"   Wczytano {len(df)} wierszy.")

        print("2. Mapowanie i czyszczenie danych...")
        # Zmieniamy nazwy kolumn z Excelowych na SQLowe zgodnie z mapą
        df = df.rename(columns=column_mapping)
        
        # Wybieramy tylko te kolumny, które istnieją w mapowaniu (odrzucamy śmieci z Excela)
        # Tworzymy listę wartości ze słownika mappingu
        sql_columns = list(column_mapping.values())
        
        # Filtrujemy DataFrame, żeby zawierał tylko kolumny zdefiniowane w SQL (i istniejące w Excelu)
        # intersection sprawdza część wspólną, żeby uniknąć błędu jeśli jakiejś kolumny brakuje w Excelu
        available_columns = df.columns.intersection(sql_columns)
        df = df[available_columns]

        # Zamiana pustych wartości (NaN) na None (NULL w SQL), inaczej mogą być błędy
        df = df.where(pd.notnull(df), None)

        print("3. Łączenie z bazą danych...")
        # Tworzenie Connection String
        if TRUSTED_CONNECTION == 'yes':
            params = urllib.parse.quote_plus(
                f'DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;'
            )
        else:
            params = urllib.parse.quote_plus(
                f'DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};'
            )
            
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

        print(f"4. Wysyłanie danych do tabeli {TABLE_NAME}...")
        # if_exists='append' -> dodaje dane do istniejących
        # index=False -> nie dodaje kolumny z numerem wiersza z Excela
        df.to_sql(TABLE_NAME, con=engine, if_exists='append', index=False)

        print("SUKCES! Dane zostały zaimportowane pomyślnie.")

    except Exception as e:
        print("\nWYSTĄPIŁ BŁĄD:")
        print(e)

if __name__ == "__main__":
    main()