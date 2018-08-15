@echo off
setlocal

:: Only for Navision Version above 2009! Tested for Navision 2017

:: Auditor Name       :  EY
:: Creation Date      :  03.07.2017
:: Version	      :  2.0
:: Last Change Date   :  15.08.2018
:: Change history     :  06.07.2017 Combining all queries into one, make it more dynamic and IPE secure 
::						 10.07.2017 Insert audit period to job log entries
::						 20.09.2017 Change output file type into .txt and add an separator

echo 'Start now...'
cd C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn
						 
:: Declaration of variables
:: DON'T change the audit period. It is only allowed by the EY auditor. A change of this value can be easily detected.
set auditperiod='2018-01-01 00:00:00' AND '2018-12-31 23:59:59'

:: Set your parameters here
set dbName="Demo Database NAV (10-0)"
set dbName2=Demo Database NAV (10-0)
set servername=nst4-poc\NAVDEMO
set path=C:\Users\Silvia.Inan\Documents\

:: User Group
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT * FROM [%dbName2%].dbo.[User Group]" -o "%path%14a_EY_company.txt"

:: User Group Member
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT * FROM [%dbName2%].dbo.[User Group Member]" -o "%path%15a_EY_company.txt"

:: User Group Access Control
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT * FROM [%dbName2%].dbo.[User Group Access Control]" -o "%path%16a_EY_company.txt"

:: User Group Permission Set
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT * FROM [%dbName2%].dbo.[User Group Permission Set]" -o "%path%17a_EY_company.txt"

:: Permission Set(Table 22000000004)
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT [Role ID], Name FROM [%dbName2%].dbo.[Permission Set]" -o "%path%5a_EY_company.txt"

:: Permission(Table 22000000005)
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT [Role ID], [Object Type], [Object ID], [Read Permission], [Insert Permission], [Modify Permission], [Delete Permission], [Execute Permission] FROM [%dbName2%].dbo.Permission" -o "%path%6a_EY_company.txt"

:: User (Table 2000000120)
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT [User Security ID], [User Name], [Full Name], State, [Expiry Date], [Windows Security ID], [Change Password], [License Type] FROM [%dbName2%].dbo.[User]" -o "%path%7a_EY_company.txt"

:: Access Control (Table 2000000053)
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT * FROM [%dbName2%].dbo.[Access Control]" -o "%path%8a_EY_company.txt"

:: Object designer table
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT Type, [Company Name], ID, Name, Modified, Compiled, Date, Time, [Version List], Locked, [Locked By] FROM [%dbName2%].dbo.Object WHERE CONVERT(varchar(30), CONVERT( datetime, [Date]), 120) BETWEEN %auditperiod%" -o "%path%9a_EY_company.txt"

:: All Company codes
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT * FROM [%dbName2%].dbo.Company" -o "%path%10a_EY_company.txt"

:: Headers
sqlcmd -S %servername% -d %dbName% -E -Q "SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME FROM [%dbName2%].dbo.INFORMATION_SCHEMA.COLUMNS" -o "%path%13a_EY_company.txt"

pause
endlocal