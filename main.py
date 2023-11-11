import ftplib
import platform
import re
import socket
import uuid
import customtkinter
from tkinter import messagebox
import wmi
from wmi import GetObject

c = wmi.WMI()

customtkinter.set_default_color_theme("dark-blue")
window = customtkinter.CTk()
window.title("Осмотр ПК")
window.geometry('400x400')
window.resizable(False, False)

userid_entry = customtkinter.CTkEntry(window, placeholder_text="ПІБ співробітника", width=370)
userid_entry.grid(row=1, column=0, padx=15, pady=5, sticky='nw')
userid_label = customtkinter.CTkLabel(window,
                                      text=f"Введіть тільки ПРІЗВИЩЕ співробітника ЦКС \n(користувача ПК)")
userid_label.grid(row=0, column=0, padx=15, pady=5, sticky='n')

userid1_label = customtkinter.CTkLabel(window, text="ПРИКЛАД: Шевченко", text_color="red")
userid1_label.grid(row=2, column=0, padx=15, pady=5, sticky='n')

address_entry = customtkinter.CTkEntry(window, placeholder_text="Адреса ПК", width=370)
address_entry.grid(row=4, column=0, padx=15, pady=5, sticky='nw')
address_label = customtkinter.CTkLabel(window,
                                       text="Введіть АДРЕСУ підрозділу ЦКС (лише назву вул.)")
address_label.grid(row=3, column=0, padx=15, pady=5, sticky='n')
address1_label = customtkinter.CTkLabel(window, text="ПРИКЛАД: Донця", text_color="red")
address1_label.grid(row=5, column=0, padx=15, pady=5, sticky='n')

phone_entry = customtkinter.CTkEntry(window, placeholder_text="Номер телефону", width=370)
phone_entry.grid(row=7, column=0, padx=15, pady=5, sticky='nw')
phone_label = customtkinter.CTkLabel(window,
                                     text="Будь ласка, введіть ВАШ НОМЕР ТЕЛЕФОНУ ДЛЯ ЗВ’ЯЗКУ")
phone_label.grid(row=6, column=0, padx=15, pady=5, sticky='n')
phone1_label = customtkinter.CTkLabel(window, text="ПРИКЛАД: 067-123-45-67", text_color="red")
phone1_label.grid(row=8, column=0, padx=15, pady=0, sticky='n')


def monic_model():
    objWMI = GetObject('winmgmts:\\\\.\\root\\WMI').InstancesOf('WmiMonitorID')
    for obj in objWMI:
        if obj.UserFriendlyName is not None:
            res = ''.join(chr(i) for i in obj.UserFriendlyName)
            return str(res)


def monic_serial():
    objWMI = GetObject('winmgmts:\\\\.\\root\\WMI').InstancesOf('WmiMonitorID')
    for obj in objWMI:
        if obj.serialnumberid != None:
            res = ''.join(chr(i) for i in obj.serialnumberid)
            return str(res)


def get_ram_info():
    ram_info = []

    try:
        ram_modules = c.Win32_PhysicalMemory()

        for i, module in enumerate(ram_modules, start=1):
            capacity_mb = module.Capacity if hasattr(module, 'Capacity') else "N/A"
            if capacity_mb != "N/A":
                capacity_gb = round(
                    float(capacity_mb) / (1024.0 ** 3))  # Преобразование из МБ в ГБ и округление до целого числа
                capacity_str = str(capacity_gb)
            else:
                capacity_str = "N/A"
            ram_info.append({
                "RAM Module": i,
                "Manufacturer": module.Manufacturer if hasattr(module, 'Manufacturer') else "N/A",
                "Speed (MHz)": module.Speed if hasattr(module, 'Speed') else "N/A",
                "Capacity (GB)": capacity_str
            })

    except Exception as e:
        print(f"Error: {str(e)}")

    return ram_info


def get_virtual_memory_info():
    for item in c.Win32_OperatingSystem():
        total_virtual_memory_kb = int(item.TotalVisibleMemorySize)
        total_virtual_memory_gb = round(total_virtual_memory_kb / (1024.0 ** 2))
        return total_virtual_memory_gb


def getMachine_addr():
    for item in c.Win32_BIOS():
        serial_number = item.SerialNumber
        return serial_number


def get_cpu_type():
    try:
        cpus = c.Win32_Processor()

        if cpus:
            return cpus[0].Name
        else:
            return "N/A"
    except Exception as e:
        return str(e)


def get_hdd_info():
    hdd_info = []

    for index, item in enumerate(c.Win32_DiskDrive(), start=1):
        tagReplacedSlashes = item.Name.replace("\\", "")
        if 'PHYSICALDRIVE' in tagReplacedSlashes:
            hdd_model = item.Model
            hdd_size_bytes = int(item.Size)
            hdd_size_gb = round(hdd_size_bytes / (1024.0 ** 3))
            hdd_info.append({"Index": index, "Model": hdd_model, "SizeGB": hdd_size_gb})

    return hdd_info


def getSystemInfo():
    mb_manufacturer = c.Win32_BaseBoard()[0].Manufacturer
    mb_model = c.Win32_Baseboard()[0].Product
    ram_info = get_ram_info()
    hdd_all = get_hdd_info()
    lines = [
        "-------------------------------------------------------------",
        f"Пренадлежность к РСЦ(адрес): {address_entry.get()}",
        f"Фамилия сотрудника: {userid_entry.get()}",
        f"Номер: {phone_entry.get()}",
        "-----------------------  Windows  -------------------------",
        f"ОС: {platform.system()} {platform.release()}",
        f"Версия ОС: {platform.version()}",
        "                                                      "
        "\n----------------------  Материнская плата  ----------------",
        f"Производитель: {mb_manufacturer}",
        f"Модель: {mb_model}",
        "                                                       "
        "\n----------------------  Инфо о ПК  ------------------------",
        f"Имя ПК: {socket.gethostname()}",
        f"Серийный номер: {getMachine_addr()}",
        f"IP Адрес: {socket.gethostbyname(socket.gethostname())}",
        f"Мак-Адрес: {':'.join(re.findall('..', '%012x' % uuid.getnode()))}",
        "                                                      "
        "\n-----------------------  Инфо о CPU  ----------------------",
        f"CPU: {get_cpu_type()}",
        "                                                              "
        "\n---------------------  Инфо о HDD/SSD  ---------------------",
    ]
    if hdd_all:
        for hdd in hdd_all:
            lines.extend([
                f"HDD {hdd['Index']}",
                f"Модель: {hdd['Model']}",
                f"Размер: {hdd['SizeGB']} ГБ",
                "                                                           ",
            ])
        lines.extend([
            "-------------------  Оперативная память  -------------------",
            f"Объём оперативы (Общий): {get_virtual_memory_info()} Гб"]),

    if ram_info:
        for info in ram_info:
            lines.extend([
                f"Модуль: {info['RAM Module']}:",
                f"Производитель: {info['Manufacturer']}",
                f"Частота (MHz): {info['Speed (MHz)']}",
                f"Объём: {info['Capacity (GB)']} ГБ",
                "------------------------------------------------------------",
            ])

    lines.extend([
        "                                                       "
        "\n----------------------  Монитор  ------------------------",
        f"Модель: {monic_model()}",
        f"Серийный номер: {monic_serial()}"
    ])

    with open('invent.txt', 'w', encoding='utf-8') as f:
        for line in lines:
            f.write(line)
            f.write('\n')

    return lines


def ftp_send():
    global ftp
    try:
        getSystemInfo()
        ftp = ftplib.FTP('ftp.src.kiev.ua')
        ftp.login("ftp4pc", "0uKT8whQ3")
        filename = "invent.txt"
        with open(filename, "rb") as file:
            ftp.storbinary(f"STOR {filename}", file)

        ftp.rename('invent.txt',
                   f"{str(socket.gethostname())} {address_entry.get()} {userid_entry.get()} ({str('-'.join(re.findall('..', '%012x' % uuid.getnode())))}).txt")
        messagebox.showinfo(title="Файл вiдправлено", message="Файл з характеристиками вiдправлено.", icon="info")
    except Exception as e:
        messagebox.showinfo(title="Помилка", message=f"Помилка при вiдправцi файлу. Перевірте з'єднання з Інтернетом", icon="error")

    finally:
        try:
            file.close()
            ftp.quit()
        except:
            pass


search_button = customtkinter.CTkButton(window, text="Отправить", command=ftp_send, corner_radius=10, width=150)
search_button.grid(row=9, column=0, padx=5, pady=5)

if __name__ == '__main__':
    window.mainloop()
