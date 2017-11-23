@echo off

sqlite3.exe CsvData.sqlite < CsvData.sql
sqlite3.exe CsvData.sqlite < CsvData_index.sql

exit