# Excel×SQL

This project is design to make possible that use Excel Sheets as database tables and run SQL query on them.
For those are familiar with database and SQL but not expert for Excel, it might help.

## How this works?

Simple, read the excel and set up a database of SQLite3, then you can run query SQL, export result and more.

## Usage

See `example`.

Note:

1. Prepare the source `.xlsx` files. Excel sheet cells are sometimes expressing strangely, especially on numeric values.
1. Prepare the sqlite3 database file. As default memory embedded database instance might be a waste of memory, you may determine a file.
1. Use a proper load method to load. For almost situations, `loadSheetWithTypeDefinition` works best.
1. You may need to output as xlsx file. Note currently only one sheet might be generated. 

## Donate

You can donate through this alipay account.

![http://www.everstray.com/resources/img/AlipayUkanokan258.png](http://www.everstray.com/resources/img/AlipayUkanokan258.png)