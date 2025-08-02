Excel Splitter

Описание:
    Скрипт разбивает листы Excel-файла на несколько файлов по заданному числу строк.

Установка зависимостей:
    pip install pandas openpyxl

Запуск скрипта:
    python excel_splitter.py <путь_к_файлу.xlsx> [-l LINESCOUNT] [-o OUTPUT_DIR]

Пример:
    python excel_splitter.py orders.xlsx --linescount 20 --output results

Создание исполняемого .exe для Windows:
    1. Установите PyInstaller:
       pip install pyinstaller

    2. Соберите exe (один файл):
       pyinstaller --onefile --name excel_splitter main.py

    3. В каталоге dist появится файл excel_splitter.exe
       Скопируйте его вместе с любыми Excel-файлами и запускайте без установки Python.
