# X2TM

Map Excel table to template.

## Install

You have to install [python3](https://www.python.org/) first.

Then, if you are on Windows run ```install.bat``` or if you are on POSIX (MacOS/Linux) run ```install.bash``` both can be found in install folder.

## How to use

### Excel File

Table head must be on row 1 only (currently, you can change it in the upcomming update). See `test/data.xlsx` for example.

### Template File

Template file must be plain text file with strings in table head surrounded by `[]` in places where you want each element in each row to go. See `test/temp.txt` for example.

### GUI

Nah, you can just figure it out. xD

### CLI

Run the scirpt via terminal.

On Windows

```batch
py xlsx-to-text-mapper.py /x path\to\excel.xlsx /t path\to\template.txt /o path\to\output.txt
```

On POSIX

```bash
python3 ./xlsx-to-text-mapper.py -x path/to/excel.xlsx -t path/to/template.txt -o path/to/output.txt
```

## Dependencies

This program uses: `xlrd` and `tkinter`
