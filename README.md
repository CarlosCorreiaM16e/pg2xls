# pg2xls

Create Excel files from Postgresql databases

## USAGE
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

## EXAMPLES 
```
$ python pg2xls.py -d mydb -t mytable -f sheet1 -T "List of mytable"

$ python pg2xls.py -d mydb \
-q "select * from table1 join table2 using( some_key ) order by 1" \
-f sheet2 -T "List of table1 and table2"
```
