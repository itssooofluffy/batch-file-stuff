@echo off
setlocal enabledelayedexpansion

:: Only for Navision Version above 2009! Tested for Navision 2017

:: Auditor Name       :  EY
:: Creation Date      :  03.07.2017
:: Version	      :  2.0
:: Last Change Date   :  15.08.2018
:: Change history     :  06.07.2017 Combining all queries into one, make it more dynamic and IPE secure 
::						 10.07.2017 Insert audit period to job log entries
::						 20.09.2017 Change output file type into .txt and add an separator
::						 16.08.2018 Change output file type into .csv and export method into sqlcmd
::						 22.08.2018 Insert Company Loop 
	
echo 'Start at %DATE%'

:: This is the default location of SQLCMD.EXE; in case this file is stored somewhere else change the path here 
cd C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn
						 
:: Declaration of variables
:: DON'T change the audit period. It is only allowed by the EY auditor. A change of this value can be easily detected.
set auditperiod='2018-01-01 00:00:00' AND '2018-12-31 23:59:59'

:: Set your parameters here
set dbName=Demo Database NAV (10-0)
set servername=nst4-poc\NAVDEMO
set dirPath=C:\Users\Silvia.Inan\Documents\
set companiesInScope="CRONUS International Ltd","TEST"

:: Start SQL Commands

:: User Group
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[User Group]" -o "%dirPath%14a_EY_company.csv" -s"," -w 999 -W 

:: User Group Member
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[User Group Member]" -o "%dirPath%15a_EY_company.csv" -s"," -w 999 -W 

:: User Group Access Control
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[User Group Access Control]" -o "%dirPath%16a_EY_company.csv" -s"," -w 999 -W 

:: User Group Permission Set
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[User Group Permission Set]" -o "%dirPath%17a_EY_company.csv" -s"," -w 999 -W 

:: Permission Set(Table 22000000004)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Role ID], Name FROM [%dbName%].dbo.[Permission Set]" -o "%dirPath%5a_EY_company.csv" -s"," -w 999 -W  

:: Permission(Table 22000000005)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Role ID], [Object Type], [Object ID], [Read Permission], [Insert Permission], [Modify Permission], [Delete Permission], [Execute Permission] FROM [%dbName%].dbo.Permission" -o "%dirPath%6a_EY_company.csv" -s"," -w 999 -W 

:: User (Table 2000000120)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [User Security ID], [User Name], [Full Name], State, [Expiry Date], [Windows Security ID], [Change Password], [License Type] FROM [%dbName%].dbo.[User]" -o "%dirPath%7a_EY_company.csv" -s"," -w 999 -W 

:: Access Control (Table 2000000053)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[Access Control]" -s"," -o "%dirPath%8a_EY_company.csv" -h-1 -w 999 -W 

:: Object designer table
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT Type, [Company Name], ID, Name, Modified, Compiled, Date, Time, [Version List], Locked, [Locked By] FROM [%dbName%].dbo.Object WHERE CONVERT(varchar(30), CONVERT( datetime, [Date]), 120) BETWEEN %auditperiod%" -o "%dirPath%9a_EY_company.csv" -s"," -w 999 -W 

:: All Company codes
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.Company" -o "%dirPath%10a_EY_company.csv" -s"," -w 999 -W 
sqlcmd -S %servername% -d "%dbName%" -E -Q "SET NOCOUNT ON SELECT Name FROM [%dbName%].dbo.Company" -s"," -o "%dirPath%HelperFile.csv" -h-1 -w 999 -W

:: Headers
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME FROM [%dbName%].dbo.INFORMATION_SCHEMA.COLUMNS" -o "%dirPath%13a_EY_company.csv" -s"," -w 999 -W 

set count=0
for %%j in (%companiesInScope%) do set /A count+=1 
	for /L %%A in (1,1,%count%) do (
		call :myLoopFunc %%A
)

:myLoopFunc
for /F "tokens=%1 delims=," %%I in ("%companiesInScope%") do (
	set currentParameter=%%I
	set currentParameter=!currentParameter:~1,-1!
	call :checkForSpacesInTableName
)
exit /b

:checkForSpacesInTableName
if "%currentParameter%"=="%currentParameter: =%" (call :SQLTableNamesWithoutSpace) else (call :SQLTableNamesWithSpace)
exit /b

:SQLTableNamesWithoutSpace

:: Change Log Entries (Table 405)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Entry No_],[Date and Time],[User ID],[Table No_],[Field No_],[Type of Change],[Old Value],[New Value],[Primary Key],[Primary Key Field 1 No_],[Primary Key Field 1 Value],[Primary Key Field 2 No_],[Primary Key Field 2 Value],[Primary Key Field 3 No_],[Primary Key Field 3 Value],[Record ID] FROM [%dbName%].dbo.[!currentparameter!$Change Log Entry] WHERE CONVERT(varchar(30), CONVERT( datetime, [Date and Time]), 120) BETWEEN %auditperiod%"  -s"," -o "%dirPath%1a_EY_!currentParameter!_Change_Log_Entry.csv" -w 999 -W

:: Change Log Entries (Table 405) filtered to Access Control Table 
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Entry No_],[Date and Time],[User ID],[Table No_],[Field No_],[Type of Change],[Old Value],[New Value],[Primary Key],[Primary Key Field 1 No_],[Primary Key Field 1 Value],[Primary Key Field 2 Value],[Primary Key Field 3 No_],[Primary Key Field 3 Value],[Record ID] FROM [%dbName%].dbo.[!currentparameter!$Change Log Entry] WHERE CONVERT(varchar(30), CONVERT( datetime, [Date and Time]), 120) BETWEEN %auditperiod% AND [Table No_] = '2000000053'" -s"," -o "%dirPath%1b_EY_!currentParameter!_Change_Log_Entry_filtered.csv" -w 999 -W

:: Change Log Setup (Table 402)
:: 1 = true; 0 = NULL (not set)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Change Log Activated] FROM [%dbName%].dbo.[!currentparameter!$Change Log Setup]" -s"," -o "%dirPath%2a_EY_!currentParameter!_Change_Log_Setup.csv" -w 999 -W

::Change Log Tables (Table 403)
:: 2 = All Fields; 1 = Some Fields; 0 0 NULL (not set)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Table No_],[Log Insertion],[Log Modification],[Log Deletion] FROM [%dbName%].dbo.[!currentparameter!$Change Log Setup (Table)]" -s"," -o "%dirPath%3a_EY_!currentParameter!_Change_Log_Table.csv" -w 999 -W

:: Change Log Fields (Table 404)
:: 1 = true; 0 = NULL (not set)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Table No_],[Log Insertion],[Log Modification],[Log Deletion] FROM [%dbName%].dbo.[!currentparameter!$Change Log Setup (Field)]" -s"," -o "%dirPath%4a_EY_!currentParameter!_Change_Log_Field.csv" -w 999 -W

:: Job Queue Entry 
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[!currentparameter!$Job Queue Entry]" -s"," -o "%dirPath%11a_EY_!currentParameter!_Job_Queue_Entry.csv" -w 999 -W

:: Job Queue Log Entry 
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[!currentparameter!$Job Queue Log Entry] WHERE CONVERT(varchar(30), CONVERT( datetime, [Start Date_Time]), 120) BETWEEN %auditperiod%" -s"," -o "%dirPath%12a_EY_!currentParameter!_Job_Queue_Log_Entry.csv" -w 999 -W
exit /b

:SQLTableNamesWithSpace

:: Change Log Entries (Table 405)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Entry No_],[Date and Time],[User ID],[Table No_],[Field No_],[Type of Change],[Old Value],[New Value],[Primary Key],[Primary Key Field 1 No_],[Primary Key Field 1 Value],[Primary Key Field 2 No_],[Primary Key Field 2 Value],[Primary Key Field 3 No_],[Primary Key Field 3 Value],[Record ID] FROM [%dbName%].dbo.[!currentparameter!_$Change Log Entry] WHERE CONVERT(varchar(30), CONVERT( datetime, [Date and Time]), 120) BETWEEN %auditperiod%"  -s"," -o "%dirPath%1a_EY_!currentParameter!_Change_Log_Entry.csv" -w 999 -W

:: Change Log Entries (Table 405) filtered to Access Control Table 
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Entry No_],[Date and Time],[User ID],[Table No_],[Field No_],[Type of Change],[Old Value],[New Value],[Primary Key],[Primary Key Field 1 No_],[Primary Key Field 1 Value],[Primary Key Field 2 Value],[Primary Key Field 3 No_],[Primary Key Field 3 Value],[Record ID] FROM [%dbName%].dbo.[!currentparameter!_$Change Log Entry] WHERE CONVERT(varchar(30), CONVERT( datetime, [Date and Time]), 120) BETWEEN %auditperiod% AND [Table No_] = '2000000053'" -s"," -o "%dirPath%1b_EY_!currentParameter!_Change_Log_Entry_filtered.csv" -w 999 -W

:: Change Log Setup (Table 402)
:: 1 = true; 0 = NULL (not set)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Change Log Activated] FROM [%dbName%].dbo.[!currentparameter!_$Change Log Setup]" -s"," -o "%dirPath%2a_EY_!currentParameter!_Change_Log_Setup.csv" -w 999 -W

::Change Log Tables (Table 403)
:: 2 = All Fields; 1 = Some Fields; 0 0 NULL (not set)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Table No_],[Log Insertion],[Log Modification],[Log Deletion] FROM [%dbName%].dbo.[!currentparameter!_$Change Log Setup (Table)]" -s"," -o "%dirPath%3a_EY_!currentParameter!_Change_Log_Table.csv" -w 999 -W

:: Change Log Fields (Table 404)
:: 1 = true; 0 = NULL (not set)
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT [Table No_],[Log Insertion],[Log Modification],[Log Deletion] FROM [%dbName%].dbo.[!currentparameter!_$Change Log Setup (Field)]" -s"," -o "%dirPath%4a_EY_!currentParameter!_Change_Log_Field.csv" -w 999 -W

:: Job Queue Entry 
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[!currentparameter!_$Job Queue Entry]" -s"," -o "%dirPath%11a_EY_!currentParameter!_Job_Queue_Entry.csv" -w 999 -W

:: Job Queue Log Entry 
sqlcmd -S %servername% -d "%dbName%" -E -Q "SELECT * FROM [%dbName%].dbo.[!currentparameter!_$Job Queue Log Entry] WHERE CONVERT(varchar(30), CONVERT( datetime, [Start Date_Time]), 120) BETWEEN %auditperiod%" -s"," -o "%dirPath%12a_EY_!currentParameter!_Job_Queue_Log_Entry.csv" -w 999 -W
exit /b 

endlocal
