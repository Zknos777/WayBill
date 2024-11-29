import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
from tkcalendar import DateEntry
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
from babel.dates import format_date
import os

# Глобальная переменная для хранения загруженных данных
data_df, vehicle_data_df, organisation_data_df = None, None, None
context = {}
selected_name, selected_vehicle, selected_organisation = None, None, None


def print_file(file_path):
    # Печать файла (можно использовать команду системы)
    try:
        os.startfile(file_path, "print")
        messagebox.showinfo("Информация", f"Файл '{file_path}' отправлен на печать.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось отправить файл на печать: {e}")

def make_card():
    global selected_name
    # Создаем папку, если она не существует
    if not os.path.exists(selected_name.split()[0]):
        os.makedirs(selected_name.split()[0])
    else:
        clear_output_folder(selected_name.split()[0])

    doc = DocxTemplate("card_template.docx")
    context["fullname"] = selected_name
    # подставляем контекст в шаблон
    doc.render(context)
    # сохраняем и смотрим, что получилось
    file_path = os.path.join(selected_name.split()[0], f"Карточка водителя.docx")
    doc.save(file_path)
    print_file(file_path)


def clear_output_folder(folder_path):
    # Проверяем, существует ли папка
    if os.path.exists(folder_path):
        # Получаем список всех файлов в папке
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            # Проверяем, является ли это файлом (а не папкой)
            if os.path.isfile(file_path):
                os.remove(file_path)  # Удаляем файл
        print(f"Все файлы в папке '{folder_path}' были удалены.")
    else:
        print(f"Папка '{folder_path}' не существует.")


def make_docs_files():
    global selected_name
    # Создаем папку, если она не существует
    if not os.path.exists(selected_name.split()[0]):
        os.makedirs(selected_name.split()[0])
    else:
        clear_output_folder(selected_name.split()[0])
    doc = DocxTemplate("waybill_template.docx")
    start_data = datetime.strptime(entries[0].get(), '%d.%m.%Y')
    stop_data = datetime.strptime(entries[1].get(), '%d.%m.%Y')
    for day in range(int((stop_data - start_data).days) + 1):
        context["data_start_short"] = (start_data + timedelta(days=day)).strftime('%d.%m.%Y')
        context["data_start"] = format_date(start_data + timedelta(days=day), format='long', locale='ru_RU')
        context["data_stop"] = format_date(start_data + timedelta(days=day + 1), format='long', locale='ru_RU')
        context["surname"], context["name"], context["middlename"]  = selected_name.split()
        context["fullname"] = context["surname"] + " " + context["name"][0] + "." + context["middlename"][0] + "."
        # подставляем контекст в шаблон
        doc.render(context)
        # сохраняем и смотрим, что получилось
        file_path = os.path.join(selected_name.split()[0], f"{context['data_start_short']}.docx")
        doc.save(file_path)
        try:
            os.startfile(file_path, "print")
            messagebox.showinfo("Информация", f"Файл '{file_path}' отправлен на печать.")
        except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось отправить файл на печать: {e}")


def load_data():
    """Загрузка данных из Excel файла."""
    global data_df
    try:
        data_df = pd.read_excel('Водители.xlsx')
        name_list = data_df['ФИО'].tolist()
        name_combobox['values'] = name_list
    except FileNotFoundError:
        messagebox.showerror("Ошибка", "Файл Водители.xlsx не найден.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")


def load_vehicle_data():
    """Загрузка данных из второго Excel файла."""
    global vehicle_data_df
    try:
        vehicle_data_df = pd.read_excel('Автомобили.xlsx')
        vehicle_number_list = vehicle_data_df['Гос номер'].tolist()
        vehicle_combobox['values'] = vehicle_number_list
    except FileNotFoundError:
        messagebox.showerror("Ошибка", "Файл Автомобили.xlsx не найден.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")


def load_organisation_data():
    """Загрузка данных организаций из Excel файла."""
    global organisation_data_df
    try:
        organisation_data_df = pd.read_excel('Организации.xlsx')
        organisation_number_list = organisation_data_df['Организация'].tolist()
        #print(1, organisation_number_list)
        organisation_combobox['values'] = organisation_number_list
    except FileNotFoundError:
        messagebox.showerror("Ошибка", "Файл Организации.xlsx не найден.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")

def on_vehicle_select(event):
    global selected_number
    """Заполнение полей при выборе гос номера из списка."""
    selected_number = vehicle_combobox.get()
    context["number_auto"] = selected_number
    if selected_number in vehicle_data_df['Гос номер'].values:
        row = vehicle_data_df[vehicle_data_df['Гос номер'] == selected_number].iloc[0]
        vehicle_entries[0].delete(0, tk.END)
        vehicle_entries[0].insert(0, row['Модель'])
        context["model_auto"] = row["Модель"]
        vehicle_entries[1].delete(0, tk.END)
        vehicle_entries[1].insert(0, row['Разрешение'])
        context["permission"] = row["Разрешение"]
        vehicle_entries[2].delete(0, tk.END)
        vehicle_entries[2].insert(0, row['Номер в реестре'])
        context["number_reestr"] = row["Номер в реестре"]
    else:
        messagebox.showwarning("Предупреждение", "Выбранный гос номер не найден.")

def on_organisation_select(event):
    global selected_organisation
    """Заполнение полей при выборе организации из списка."""
    selected_organisation = organisation_combobox.get()
    context["organisation"] = selected_organisation
    if selected_organisation in organisation_data_df['Организация'].values:
        row = organisation_data_df[organisation_data_df['Организация'] == selected_organisation].iloc[0]
        # organisation_entries[0].delete(0, tk.END)
        # organisation_entries[0].insert(0, row['Адрес'])
        context["adress"] = row["Адрес"]
        # organisation_entries[1].delete(0, tk.END)
        # organisation_entries[1].insert(0, row['Действителен с'])
        context["valid_from"] = row["Действителен с"]
        # organisation_entries[2].delete(0, tk.END)
        # organisation_entries[2].insert(0, row['Действителен до'])
        context["valid_until"] = row["Действителен до"]
        context["organisation_phone"] = row["Телефон"]
        context["OGRN"] = row["ОГРН"]
        context["INN"] = row["ИНН"]
        context["OKPO"] = row["ОКПО"]
    else:
        messagebox.showwarning("Предупреждение", "Выбранная организация не найдена.")


def on_name_select(event):
    global selected_name, context
    """Заполнение полей при выборе имени из списка."""
    selected_name = name_combobox.get()

    if selected_name in data_df['ФИО'].values:
        row = data_df[data_df['ФИО'] == selected_name].iloc[0]
        entries[2].delete(0, tk.END)
        entries[2].insert(0, row['Права'])
        context["drive_license"] = row["Права"]
        entries[3].delete(0, tk.END)
        entries[3].insert(0, row['Снилс'])
        context["snils"] = row["Снилс"]
        entries[4].delete(0, tk.END)
        entries[4].insert(0, row['Телефон'])
        context["phone"] = row["Телефон"]
    else:
        messagebox.showwarning("Предупреждение", "Выбранное имя не найдено.")


# Создаем основное окно
root = tk.Tk()
root.title("Генератор путевых листов")

# Поля для первой формы
entries = []
labels = []

#Создание форм даты
label = tk.Label(root, text="Дата начала")
label.grid(row=1, column=0, sticky="w", padx=5, pady=5)  # Выравнивание по левому краю
entry = DateEntry(root, locale="ru", datepattern="mm/dd/y", width=19, background='darkblue', foreground='white', borderwidth=2)
entry.insert(0, "")  # Заполнение поля данными
entry.grid(row=1, column=1, padx=5, pady=5)  # Поле ввода рядом с меткой
labels.append(label)
entries.append(entry)

label = tk.Label(root, text="Дата окончания")
label.grid(row=2, column=0, sticky="w", padx=5, pady=5)  # Выравнивание по левому краю
entry = DateEntry(root, locale="ru", datepattern="mm/dd/y", width=19, background='darkblue', foreground='white', borderwidth=2)
entry.insert(0, "")  # Заполнение поля данными
entry.grid(row=2, column=1, padx=5, pady=5)  # Поле ввода рядом с меткой
labels.append(label)
entries.append(entry)

# Комбобокс для выбора водители
label = tk.Label(root, text="Водитель")
label.grid(row=3, column=0, sticky="w", padx=5, pady=5)  # Выравнивание по левому краю
labels.append(label)
name_combobox = ttk.Combobox(root, width=19)
name_combobox.grid(row=len(entries)+1, column=1, columnspan=2, pady=5)
name_combobox.bind("<<ComboboxSelected>>", on_name_select)

for title in ["Права", "Снилс", "Телефон"]:
    label = tk.Label(root, text=title)
    label.grid(row=len(entries)+2, column=0, sticky="w", padx=5, pady=5)

    entry = tk.Entry(root)
    entry.grid(row=len(entries)+2, column=1, padx=5, pady=5, ipadx=3)

    labels.append(label)
    entries.append(entry)



# Поля для второй формы (транспортные средства)
vehicle_entries = []
vehicle_labels = []

# Комбобокс для выбора гос номера
label = tk.Label(root, text="Автомобиль")
label.grid(row=len(vehicle_entries) + len(entries) + 2, column=0, sticky="w", padx=5, pady=5)  # Выравнивание по левому краю
labels.append(label)
vehicle_combobox = ttk.Combobox(root, width=19)
vehicle_combobox.grid(row=len(vehicle_entries) + len(entries) + 2, column=1, columnspan=2,padx=5, pady=5)
vehicle_combobox.bind("<<ComboboxSelected>>", on_vehicle_select)


for title in ["Модель", "Разрешение", "Номер в реестре"]:
    label = tk.Label(root, text=title)
    label.grid(row=len(vehicle_entries) + len(entries) + 3, column=0, sticky="w", padx=5, pady=5)

    entry = tk.Entry(root)
    entry.grid(row=len(vehicle_entries) + len(entries) + 3, column=1, ipadx=3, padx=5, pady=5)

    vehicle_labels.append(label)
    vehicle_entries.append(entry)


#Выбор организации
organisation_entries = []
organisation_labels = []

# Комбобокс для выбора организаций
label = tk.Label(root, text="Организация")
label.grid(row=len(vehicle_entries) + len(entries) + 3, column=0, sticky="w", padx=5, pady=5)  # Выравнивание по левому краю
labels.append(label)
organisation_combobox = ttk.Combobox(root, width=19)
organisation_combobox.grid(row=len(vehicle_entries) + len(entries) + 3, column=1, columnspan=2,padx=5, pady=5)
organisation_combobox.bind("<<ComboboxSelected>>", on_organisation_select)


#Чек_вар галочка для печати
check_var = tk.BooleanVar()
label_state = tk.Label(root, text='Состояние: ' + str(check_var.get()))
label_state.grid(row=len(vehicle_entries) + len(entries) + len(organisation_entries) + 4, column=1, pady=5)
check = tk.Checkbutton(root, text='Распечатать', variable=check_var)
check.grid(row=len(vehicle_entries) + len(entries) + len(organisation_entries) + 4, column=0, pady=5)



# Кнопка "Создать карточку"
create_card_button = tk.Button(root, text="Создать карточку", command=make_card, width=19)
create_card_button.grid(row=len(vehicle_entries) + len(entries) + len(organisation_entries) + 5, column=0, pady=5)


# Кнопка "Создать путевые"
create_card_button = tk.Button(root, text="Создать путевые", command=make_docs_files, width=19)
create_card_button.grid(row=len(vehicle_entries) + len(entries) + len(organisation_entries) + 5, column=1, pady=5)

# Кнопка "Выход"
exit_button = tk.Button(root, text="Выход", command=root.quit)
exit_button.grid(row=len(vehicle_entries) + len(entries) + len(organisation_entries)+ 7, columnspan=2, pady=5)

# Запускаем главный цикл приложения
load_data()
load_vehicle_data()
load_organisation_data()
root.mainloop()
