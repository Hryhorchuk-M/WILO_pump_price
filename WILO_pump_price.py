import tkinter as tk
from tkinter import messagebox
import openpyxl


def save_currency_rate():
    new_rate = entry.get()

    new_rate = new_rate.replace(',', '.')

    if new_rate.replace('.', '', 1).isdigit():
        try:
            with open('currency_rate.txt', 'w') as file:
                file.write(new_rate)
            messagebox.showinfo("Успішно", f"Курс валют оновлено: {new_rate}")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося зберегти курс валют: {str(e)}")
    else:
        messagebox.showwarning("Увага", "Введіть коректний курс валют (число) перед збереженням!")


def load_currency_rate():
    try:
        with open('currency_rate.txt', 'r') as file:
            rate = file.read()
            if rate:
                entry.delete(0, 'end')
                entry.insert(0, rate)
    except Exception as e:
        pass


def get_pump_price():

    pump_info = pump_info_entry.get()

    try:
        # Відкриваємо Excel-файл з цінами на насоси
        workbook = openpyxl.load_workbook('price.xlsx')
        sheet = workbook.active
        # Шукаємо відповідний рядок за артикулом або назвою насосу
        for row in sheet.iter_rows(values_only=True):
            if pump_info in row:
                # Отримуємо ціну насосу (припускаємо, що вона знаходиться у третьому стовпці)
                price = row[2]
                return price
        # Якщо насос не знайдено, повертаємо None
        return None
    except Exception as e:
        # Обробка можливих помилок при зчитуванні Excel-файлу
        print(f"Помилка при зчитуванні цін: {e}")
        return None


root = tk.Tk()
root.title("Курс валют")

label = tk.Label(root, text="Курс валют:")
label.pack()

entry = tk.Entry(root)
entry.pack()

save_button = tk.Button(root, text="Зберегти", command=save_currency_rate)
save_button.pack()

# Створення мітки та поля для введення артикулу або назви насосу
pump_info_label = tk.Label(root, text="Назва або артикул насосу:")
pump_info_label.pack()

pump_info_entry = tk.Entry(root)
pump_info_entry.pack()

# Кнопка для зчитування ціни насосу
read_pump_price_button = tk.Button(root, text="Зчитати ціну насосу", command=get_pump_price)
read_pump_price_button.pack()

load_currency_rate()

root.mainloop()
