# PS Excel Merge

PS Excel Merge is written in Python openpyxl to deal with automating VLOOKUP .dbf files (DELIVERY, DRIVER and DELISITE). 
Creates a new excel file "PS DELIVERY MMDD". MMDD is today's date.


## How to Run
### 1. Executable
- Download the `PS Excel Merge` zip file and unzip it
- Put DELIVERY, DRIVER and DELISITE in the `PS Excel Merge` folder
- Run `main.exe` **(WINDOWS only)**


### 2. CLI (cmd)
- Download the `main` zip file and unzip it
- Download dependencies in `cmd`

```cmd
pip install dbfread pandas openpyxl
```
- Change directory to target folder
```cmd
cd C:\Users\user\Downloads\PS Excel Merge
```
- Put DELIVERY, DRIVER and DELISITE in the same folder as `main.py`
- Run `main.py`

```cmd
py main.py
```


### 3. Android
- Get `Pydroid 3` from Google Play
- Open `main.py`
- Navigate to `Pip` > `Install`, uncheck `'Use prebuilt libraries repository'`, enter `dbfread`, `pandas`, `openpyxl` separately and install each at a time
- Use `Quick Install` if stuck
- Put DELIVERY, DRIVER and DELISITE in the same folder as `main.py`
- Run `main.py`


## Usage

```cmd
Starting... Merging...
DRIVER, FROM, TO¡G columns added
Sheet1, Sheet2 created and copied over from DRIVER and DELISITE
All sheets sorted by column A!
Column G formatted to 'D-MMM' format successfully!
VLOOKUP done!
Added Subtotal rows, grouped rows
Added Grand Total rows
Column width adjusted
Header freezed
Header: no border, bold, align left
Done! File 'PS DELIVERY MMDD.xlsx' created!
Press Enter key to close this window
```


## Misc
1. Only reads .dbf files
2. Files have to be in the same folder as `main.py` or `main.exe`
3. `Pyinstaller` is used to create the .exe. Use `pip` to install, navigate to target folder then run `pyinstaller+hidden imports` on `main.py`

```cmd
pip install pyinstaller
pyinstaller --paths "C:\Python39\Lib\site-packages" --hidden-import pkg-resources --hidden-import pandas --hidden-import dbfread --hidden-import openpyxl -F main.py
```

4. `Pip freeze` or `pipreqs` is used to create 'requirements.txt'

```cmd
pip freeze > requirements.txt
```
```cmd
pipreqs ./
```
