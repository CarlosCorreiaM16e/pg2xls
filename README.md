# pg2xls

Create Excel files from Postgresql databases

### Dependencies
```psycopg2``` and ```openpyxl``` (v. >= 2.3.2)

#### Debian/Ubuntu installing

```
$ sudo apt-get install psycopg2 python-pip
$ sudo pip install openpyxl
```
**Installing openpyxl with pip is needed as Debian newest
stable version is 1.7.0.**

## Usage
```
python pg2xls.py -d db_name \
  [-h localhost] [-p 5432] [-U username] [-W password] \
  [-t table] [-q "query"] [-f filename] [-T "sheet title" ]
```

**Optional parameters:**

```-h``` : defaults to 'localhost'

```-p``` : defaults to 5432

```-U``` : defaults to os.environ[ 'USER' ]

```-f``` : defaults to 'sheet.xlsx'

```-T``` : defaults to '' (empty)

```-W``` : read from ~/.pgpass or from console

If ```-t``` option is given, the sheet is created with all data
from the given table ordered by its first field, otherwise
```-q``` option must be given with a query to fill the spreadsheet.

## Examples
```
$ python pg2xls.py -d mydb -t mytable -f sheet1 -T "List of mytable"

$ python pg2xls.py -d mydb \
-q "select * from table1 join table2 using( some_key ) order by 1" \
-f sheet2 -T "List of table1 and table2"
```

