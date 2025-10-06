Currently in concept phase

Which tool to use on Andoid tablet for DB design? DB Designer

Which tool to use on Windows PC for database management? SQLiteStudio

How to setup ODBC entry for SQLite db?<br>
    Download sqlite odbc driver from http://www.ch-werner.de/sqliteodbc/<br>
    Start ODBC Datenquellen Administrator<br>
    Select File DSN and add an entry for your SQLite db using the SQLite3 ODBC Driver <br>
        Save the following definition in C:\Users\Volker\Documents\My Data Sources\pssService-sqlite.dsn<br>
        #Datasource Name: pssService-sqlite<br>
        Database Name: C:\MyReports\DataSources\pssService\sql-db<br>
        Lock Timeout: 100000<br>
        Sync Mode: NORMAL<br>
        




