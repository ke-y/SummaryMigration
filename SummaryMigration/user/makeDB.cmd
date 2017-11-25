@echo off

sqlite3.exe SummaryData.sqlite < SummaryData.sql
sqlite3.exe SummaryData.sqlite < SummaryData_Index.sql

exit