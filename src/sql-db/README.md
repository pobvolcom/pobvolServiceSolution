Currently in concept phase

Which tool to use on Andoid tablet for DB design? DB Designer

Which tool to use on Windows PC for database management? SQLiteStudio

How to setup ODBC entry for SQLite db?
    Download sqlite odbc driver from http://www.ch-werner.de/sqliteodbc/
    Start ODBC Datenquellen Administrator
    Select File DSN and add an entry for your SQLite db using the SQLite3 ODBC Driver 
        Save the following definition in C:\Users\Volker\Documents\My Data Sources\pssService-sqlite.dsn
        #Datasource Name: pssService-sqlite
        Database Name: C:\MyReports\DataSources\pssService\sql-db
        Lock Timeout: 100000
        Sync Mode: NORMAL
        



