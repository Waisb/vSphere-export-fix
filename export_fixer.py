try:
    import xlwt
    import csv
    import argparse
    import os.path
except:
    print("Установите необходимые библиотеки!\n\nВыполните в консоли команду:\npip install -r requirements.txt")
    exit()

parser = argparse.ArgumentParser(description='Скрипт для превращения \"кривого\" экспорта vSphere в \"нормальный\" .xls файл')
parser.add_argument("-i", "--input", type=str, default="ExportList.csv", help="Указать свое название файла для обработки.")
parser.add_argument("-o", "--output", type=str, default="Output.xls", help="Указать свое имя выходного файла.")
args = parser.parse_args()

InputFilename = args.input
OutputFilename = args.output

print (f"Имя входного файла: {InputFilename}\nИмя выходного файла: {OutputFilename}")

if os.path.exists(InputFilename) == False:
    print(f"Файл {InputFilename} не найден!")
    exit()
    
if os.path.exists(OutputFilename):
    print(f"Внимание! Файл {OutputFilename} уже существует, он будет перезаписан!")
    answer = input("Продолжить? y/n?\n> ") 
    if answer != "y":
        print("Отмена.")
        exit()
        
    
with open(InputFilename,encoding="utf8") as fp:
    reader = csv.reader(fp, delimiter=",", quotechar='"')
    data_read = [row for row in reader]

Book = xlwt.Workbook(encoding="utf-8")
Sheet = Book.add_sheet("Выгрузка")

for row, data in enumerate(data_read):
    for column, item in enumerate(data):
        Sheet.write(row, column, item)
        
Book.save(OutputFilename)

print("Файл обработан!")
