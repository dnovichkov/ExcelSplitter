# Excel Splitter

Скрипт разбивает листы Excel-файла на несколько файлов по заданному числу строк.
На вход скрипт принимает путь к файлу (относительный или абсолютный) - _filename_, количество строк (по умолчанию - 20) - _linescount_, место размещения результатов (по умолчанию - _results_).

После этого скрипт читает все листы в файле и формирует выходные файлы. Скрипт разбивает все листы в файле на несколько таблиц по _linescount_ каждый, сохранив первую строку с заголовками.
Формат именования файла - _filename_sheetname_linesfrom_linesto.xls_.
Здесь _filename_ - это название исходного файла. _sheetname_ - название листа.
_linesfrom_ - номер строки, откуда начинаются записи, _linesto_ - номер строки, которой заканчиваются записи в файле.

Допустим, был загружен файл orders.xlsx, в котором на листе Sheet1 есть 30 строк, а на листе SheetOther - 45 строк.
Тогда на выходе должен получиться такой список файлов:
* orders_Sheet1_1_20.xlsx
* orders_Sheet1_21_30.xlsx
* orders_SheetOther_1_20.xlsx
* orders_SheetOther_21_40.xlsx
* orders_SheetOther_41_45.xlsx

В каждом файле будет сначала идти строка заголовка из исходного файла с исходного листа, потом соответствующие строки.

Установка зависимостей:
    `pip install pandas openpyxl`

Запуск скрипта:
    `python excel_splitter.py <путь_к_файлу.xlsx> [-l LINESCOUNT] [-o OUTPUT_DIR]`

Пример:
    `python excel_splitter.py orders.xlsx --linescount 20 --output results`

## Создание исполняемого .exe для Windows:

    1. Установите PyInstaller:
       pip install pyinstaller

    2. Соберите exe (один файл):
       pyinstaller --onefile --name excel_splitter main.py

    3. В каталоге dist появится файл excel_splitter.exe
       Скопируйте его вместе с любыми Excel-файлами и запускайте без установки Python.
