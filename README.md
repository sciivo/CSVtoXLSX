# CSVtoXLSX
Convert CSV files to XLSX

Used for converting CSV files generated from Microsofts SQL Server Management Studio command line tool 'sqlcmd'.
eg; sqlcmd -S *SERVER* -U *USERNAME* -P *PASSWORD* -i *INPUT.SQL* -o *OUTPUT.CSV* -s "," -W -f o: 1252 -m 1

The above command will generate a CSV file given an input SQL file. Row 2, containing spacing characters will be removed, along with the last two lines which contain total row count.

Syntax:
CSVtoXLSX.exe *INPUT CSV* *OUTPUT XLSX* *ERROR LOG OUTPUT* *MERGE*

Example usage:
CSVtoXLSX.exe C:\Data.csv C:\Data.xlsx C:\Errors.txt no
Data.csv will be opened, aforementioned rows will removed, and saved as Data.xlsx.


CSVtoXLSX.exe C:\Data1.csv C:\Data.xlsx C:\Errors.txt no
CSVtoXLSX.exe C:\Data2.csv C:\Data.xlsx C:\Errors.txt yes
CSVtoXLSX.exe C:\Data3.csv C:\Data.xlsx C:\Errors.txt yes
Same as before but data contained in Data2.csv and Data3.csv will be appeneded as new sheets in Data.xlsx is the file already exists.
