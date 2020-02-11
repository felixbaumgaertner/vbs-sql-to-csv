# vbs-sql-to-csv
This simple VBS script gets SQL results and writes them to a CSV file.


## Setup the variables

Make sure to change the variables at the beginning of the file.

```
sql = "SELECT * FROM your_table"
csvFilePath = "path\to\destination\results.csv"
seperator = ";"
server = "192.168.178.100"
database = "your_database"
user = "your_user"
password = "your_password"
```

All of them have to be changed in order to run, only the seperator will probably work (you maybe want to change it to "," or customize).

### Where to get the information

* *sql* - here goes your sql statement
* *csvFilePath* - this is the **absolute filepath** to the place where the csv file will be created (e.g. "C:\Users\felixbaumgaertner\Desktop\TestExport.csv")
* *seperator* - default seperator ";" can be changed to "," or whatever you like
* *server*, *database*, *user* & *password* - those information define which database server you want to connect to (ask your administrator or technical support for this details)
