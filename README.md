# LaTeXfromExcel
## Intro
Writing LaTeX to do paper typesetting seems not a good deal, but hard for normal paper writers.

This progect helps to convert Excel data to LaTeX source code, to reduce work load of table writers.
All the.

It's developed under Python3(3.7). Tkinter, xlrd, xlwt and openpyxl are needed.
## Structure
```
.
├── console.py     
├── interface.py            
├── pyexcel.py              
├── README.md               
├── readTable.py            
└── writeTable.py           

0 directories, 5 files
```
console.py            --The console mode 

interface.py            --The gui mode

pyexcel.py              --Used to produce '.xls' or '.xlsx' samples for test

readTable.py            --Utils to read  '.xls' or '.xlsx' files to my own data format

writeTable.py           --Utils to output '.xls', '.xlsx' or '.csv' even 'LaTeX' and 'HTML' 

## Usage
### Run Source Code
Make sure you are using Python3+

Install the dependencies
```sh
pip install -r requirement.txt 
```
Run the console mode
```sh
python console.py -i inputfile -o outputfile
# e.g.
python console.py -i input.xls -o output.tex # generate LaTeX source codes
# or
python console.py -i input.csv -o output.xlsx # do file format conversion
```
Run the gui mode
```sh
python interface.py 
```
### Run Executable File
Get executable files at [this link](https://github.com/TyeYeah/LaTeXfromExcel/releases) according to your platform. They are sometimes buggy because they're not built from the latest source codes. 
## Feature
Now I developed two modules: 
1. File Format Conversion
2. Output LaTeX Source

If you want to get LaTeX format data table on a laptop without office suites and LaTeX compilation engine, this project helps to: 
* do simple file format conversion(No color supports)
* produce Table's LaTeX source code.