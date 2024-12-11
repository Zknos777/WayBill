import win32print

printer_name = "HP Laser 103 107 108"
PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}

hprinter = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)

attributes = win32print.GetPrinter(hprinter, 2)
## Настройка двухсторонней печати
attributes['pDevMode'].Nup = 2
print(attributes['pDevMode'].DriverVersion)

print(type(attributes['pDevMode']))## Передаем нужные значения в принтер
win32print.SetPrinter(hprinter, 2, attributes, 0)
attributes = win32print.GetPrinter(hprinter, 2)
print(attributes['pDevMode'].Nup)
win32print.ClosePrinter(hprinter)
print(1111111)