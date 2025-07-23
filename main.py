import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np

def load_file():
    file_path = filedialog.askopenfilename()  # Możesz wybrać dowolny plik dla testu
    if file_path:
        process_file(file_path)

def process_file(file_path):
    try:
        # Wczytaj dane z Excela
        df = pd.read_excel(file_path)

        # Wyświetl nagłówki kolumn, aby zweryfikować ich poprawność
        print("Columns in the file:", df.columns.tolist())

        # Wyświetl pierwsze kilka wierszy danych, aby zobaczyć ich strukturę
        print("Preview of the data:")
        print(df.head())

        # Przykład: Filtruj dane na podstawie kolumny "Date" na jeden tydzień
        if 'Date' not in df.columns:
            messagebox.showerror("Error", "Column 'Date' not found in the file.")
            return

        df['Date'] = pd.to_datetime(df['Date'])
        start_date = df['Date'].min()
        end_date = start_date + pd.Timedelta(days=6)
        weekly_data = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

        # Przykładowa logika przetwarzania danych
        def calculate_hours(group):
            take_actions = group[group['Action'] == 'take']
            accept_reject_actions = group[(group['Action'] == 'accept') | (group['Action'] == 'reject')]
            total_hours = 0
            for index, take_action in take_actions.iterrows():
                corresponding_actions = accept_reject_actions[accept_reject_actions['Xdeliverable'] == take_action['Xdeliverable']]
                if not corresponding_actions.empty:
                    accept_reject_time = corresponding_actions['Date'].iloc[0]
                    time_diff = accept_reject_time - take_action['Date']
                    hours = np.ceil(time_diff.total_seconds() / 3600 / 0.25) * 0.25  # zaokrąglij w górę do 0.25 godziny
                    total_hours += hours
            return total_hours

        results = weekly_data.groupby(['Campaign name', 'Date']).apply(calculate_hours).reset_index(name='Hours')

        # Eksport wyników do nowego pliku Excel
        results.to_excel('processed_data.xlsx', index=False)
        print("Plik przetworzony i zapisany jako 'processed_data.xlsx'.")
    except Exception as e:
        print(f"An error occurred: {e}")

app = tk.Tk()
app.title("CRD Processor")

# Ustawienie większego rozmiaru okna, np. 400x300 pikseli
app.geometry("400x300")

load_button = tk.Button(app, text="Load Excel File", command=load_file)
load_button.pack()

app.mainloop()