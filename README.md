# Metrology-Database
A script to upload data residing in excel files to a SQLite database. It will automatically create the necessary table for each excel file if they don't already exist. it will also once per day check the integrity of the data of the database and email the necessary individuals if the program runs into an error or bad data.
The first 4 rows of the excel file will always be headers for column names and so the actual data will always start at row 5. The program automatically determines how many rows there are and will populate a new table in the database if it requires it.
