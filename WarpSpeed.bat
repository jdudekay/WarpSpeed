@echo off
rem    WarpSpeed: Night Audit Accelerator - A tool made with the intention 
rem    of streamlining the Night Audit Process by renaming and sorting a large 
rem    number of .pdf files and generating reports based off of their contents.
rem
rem    Copyright (C) 2020-2021 John Dudek
rem
rem    This program is free software: you can redistribute it and/or modify
rem    it under the terms of the GNU General Public License as published by
rem    the Free Software Foundation, either version 3 of the License, or
rem    (at your option) any later version.
rem
rem    This program is distributed in the hope that it will be useful,
rem    but WITHOUT ANY WARRANTY; without even the implied warranty of
rem    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
rem    GNU General Public License for more details.
rem
rem    You should have received a copy of the GNU General Public License
rem    along with this program.  If not, see <https://www.gnu.org/licenses/>.
set ver=2.3.0
call :warpSpeed
:updateWS
cls
echo WarpSpeed: Night Audit Accelerator by John Dudek v%ver%
echo Downloading WarpSpeed v%webVer% . . .
cd..
powershell -Command "Invoke-WebRequest https://github.com/jdudekay/WarpSpeed/releases/latest/download/WarpSpeed.zip -OutFile WarpSpeed.zip"
cls
echo WarpSpeed: Night Audit Accelerator by John Dudek v%ver%
echo Extracting and Installing WarpSpeed v%webVer% . . . (WarpSpeed will restart when finished)
powershell Expand-Archive -Force WarpSpeed.zip
del WarpSpeed.zip
cd WarpSpeed
start cmd /K WarpSpeed.bat
exit
:warpSpeed
cls
echo -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
echo +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
echo -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-                                                                                                            -+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-                                                                                                            -+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-                888       888                        .d8888b.                              888              -+-+-+
echo -+-+-+               888   o   888                       d88P  Y88b                             888               +-+-+-
echo +-+-+-               888  d8b  888                       Y88b.                                  888               -+-+-+
echo -+-+-+               888 d888b 888 8888b. 888d88888888b.  "Y888b.  88888b.  .d88b.  .d88b.  .d88888               +-+-+-
echo +-+-+-               888d88888b888    "88b888P"  888 "88b    "Y88b.888 "88bd8P  Y8bd8P  Y8bd88" 888               -+-+-+
echo -+-+-+               88888P Y88888.d888888888    888  888      "888888  8888888888888888888888  888               +-+-+-
echo +-+-+-               8888P   Y8888888  888888    888 d88PY88b  d88P888 d88PY8b.    Y8b.    Y88b 888               -+-+-+
echo -+-+-+               888P     Y888"Y888888888    88888P"  "Y8888P" 88888P"  "Y8888  "Y8888  "Y88888               +-+-+-
echo +-+-+-                                           888               888                                            -+-+-+
echo -+-+-+                                           888               888                                            +-+-+-
echo +-+-+-                                           888               888                                            -+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-                               N i g h t   A u d i t   A c c e l e r a t o r                                -+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-                                                   v %ver%                                                  -+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-                                     Copyright (C) 2020-2021 John Dudek                                     -+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
echo -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
echo +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
pause
:restart
cls

setlocal EnableExtensions DisableDelayedExpansion
set /a count=1

rem Main program prompts
echo WarpSpeed: Night Audit Accelerator by John Dudek v%ver%
echo Checking for updates . . .
curl https://github.com/jdudekay/WarpSpeed/releases/latest/download/WarpSpeed.zip -s > temp.txt
setlocal enabledelayedexpansion
for /f "tokens=*" %%a in (temp.txt) do (
  set webVer=%%a
  set webVer=!webVer:~92,-45!
  if "!webVer!" NEQ "%ver%" ( 
    del temp.txt
	echo.
	echo Current Version: %ver%
	echo Latest Release: !webVer!
	echo.
	echo WarpSpeed will now update.
	pause
    call :updateWS 
    )
)
del temp.txt
setlocal DisableDelayedExpansion

cls
echo WarpSpeed: Night Audit Accelerator by John Dudek v%ver%
echo Type "help" for instructions
echo.

if not exist "tools\xpdf\pdftotext.exe" (
echo FATAL ERROR: pdftotext.exe is missing. 
echo WarpSpeed cannot run. 
echo WarpSpeed will attempt to update and repair itself automatically.
echo.
pause
call :updateWS
exit
)
if not exist "tools\xpdf\pdfinfo.exe" (
echo FATAL ERROR: pdfinfo.exe is missing. 
echo WarpSpeed cannot run. 
echo WarpSpeed will attempt to update and repair itself automatically.
echo.
pause
call :updateWS
exit
)

if exist "rmrtver.pdf" (
	echo Room Rate Verification Report detected: Engaging WarpDrive
	echo.
	pause
	set wdRep="rmrtver"
	call :warpDrive
	pause
	exit
)

if exist "exparvl.pdf" (
	echo Expected Arrivals Report detected: Engaging WarpDrive
	echo.
	pause
	set wdRep="exparvl"
	call :warpDrive
	pause
	exit
)

call :getAuditDate
echo Audit Pack will be generated for %nameMonth% %auditDay%, %year%.
echo.
set /p input=" Would you like to begin? (y/n): "

if "%input%" == "n" (
  echo.
  echo Closing WarpSpeed . . .
  echo.
  pause
  exit
)
rem Manual user input, can input date, or run help or update functions
if "%input%" == "debug" (
  set input=y
  set /a debug=1
)
rem Help function
if "%input%" == "help" (
start notepad "README.md" 
goto :restart
)
rem Update function
if "%input%" == "update" (
call :updateWS
exit
)

if "%input%" NEQ "y" (
  echo.
  echo Input not recognized, restarting WarpSpeed.
  echo.
  pause
  goto :restart
)

del OutputFolder\ /q 2>nul

if exist "*.pdf" (
echo.
echo FATAL ERROR: .pdf files already in WarpSpeed Folder, process aborted.
echo.
pause
exit
)

setlocal EnableDelayedExpansion
set "startTime=%time: =0%"

rem Copying of all .pdfs from p2 folder to WarpSpeed folder
echo.
echo Copying reports from eci folder . . .
if "%debug%" == "1" (
	robocopy "%cd%\debug\ " %cd% /NFL /NDL /NJH /NJS /nc /ns /np >nul
) else (
	robocopy Z:\eci %cd% /MAXAGE:1 /NFL /NDL /NJH /NJS /nc /ns /np >nul
)
if not exist "*.pdf" (
echo.
echo FATAL ERROR: .pdf files unable to be copied, process aborted.
echo.
pause
exit
)

md "OutputFolder" 2>nul
echo Identifying and renaming reports . . . 
rem Main program loop
for /F %%I in ('dir *.pdf /B') do call :RenamePDF "%%~fI"

rem Cleanup subroutine
echo Further refining report names . . .
call :cleanUp

rem Subroutine for the creation of Audit Pack
echo Creating Audit Pack . . .
call :auditPack

rem Copying of created Audit Pack to destination folder on shared drive
echo Copying Audit Pack to Cloud Drive . . .
if "%debug%" == "1" (
	robocopy "%cd%\OutputFolder\AuditPack\ " "%cd%\debug\Network Shares\Westin File Server\Accounting Public\Westin - Night Audit\%year%\%nameMonth%\%month%-%auditDay%-%year% " /NFL /NDL /NJH /NJS /nc /ns /np >nul
) else (
	robocopy "%cd%\OutputFolder\AuditPack\ " "E:\Network Shares\Westin File Server\Accounting Public\Westin - Night Audit\%year%\%nameMonth%\%month%-%auditDay%-%year% " /NFL /NDL /NJH /NJS /nc /ns /np >nul
)

rem WarpSpeed report backed up to WS_Report.txt and displayed in console before program quits
echo Generating WarpSpeed Report . . .
call :makeReport

pause
exit

:renamePDF
set "filePDF=%~1"
set "fileTXT=%~dpn1.txt"

rem Create text version of .pdf to be read
rem Parameters are set to print only the first page and in raw which condenses the data
tools\xpdf\pdftotext.exe -l 1 -raw "%filePDF%"

rem For loop is used to call find which selects the line that contains the report name
rem Names
for /F "tokens=1 delims=:" %%A in ('%SystemRoot%\System32\find.exe "Westin Medical Ctr" "%fileTXT%"') do set "report=%%A"

rem The text version of the PDF file is no longer needed.
del "%fileTXT%"

rem Trim unneccessary words from file name variable
rem Creates actual fileName variable from user inputted and computer derived variables
set report=%report:Westin Medical Ctr =%
set report=%report: Page Number=%
set "fileName=%report%_%dat%"

rem Move the PDF file to OutputFolder and rename the file while moving it.
rem Also checks to make sure no files are overwritten
:while
if not exist "OutputFolder\%fileName%.pdf" (
move "%filePDF%" "OutputFolder\%fileName%.pdf" >nul 2>nul
set /a count=1
) else (
set "fileName=%report%_%dat% (%count%)"
set /a count=%count%+1
goto :while
)

rem AR reports have the illegal "/" character, this removes it
rem Also checks to make sure no files are overwritten
if errorlevel 1 set report=%report:A/R =%
if errorlevel 1 set "fileName=AR %report%_%dat%"
if errorlevel 1 (
:whileAR
if not exist "OutputFolder\%fileName%.pdf" (
move "%filePDF%" "OutputFolder\%fileName%.pdf" >nul 2>nul
set /a count=1
) else (
set "fileName=AR %report%_%dat% (%count%)"
set /a count=%count%+1
goto :whileAR
)
)
rem Exit the subroutine RenamePDF and continue with FOR loop in main code.
goto :EOF

rem Crudely renames duplicate files based on known contents
:cleanUp
cd OutputFolder
ren "---------- C_%dat%.pdf" "Open Folio System Balancing Report_%dat%.pdf"
ren "---------- C_%dat% (1).pdf" "Open Folio System Balancing Report_%dat% (1).pdf"
ren "---------- C_%dat% (2).pdf" "Payment Register Detail Report_%dat%.pdf"
ren "---------- C_%dat% (3).pdf" "Accounts Receivable Activity Report_%dat%.pdf"
ren "Bank Transaction Report_%dat%.pdf" "Bank Transaction Report_%dat% (All).pdf"
ren "Bank Transaction Report_%dat% (1).pdf" "Bank Transaction Report_%dat% (Dcln).pdf"
ren "Bank Transaction Report_%dat% (2).pdf" "Bank Transaction Report_%dat% (Pend).pdf"
ren "Bank Transaction Report_%dat% (3).pdf" "Bank Transaction Report_%dat% (All)(1).pdf"
ren "Bank Transaction Report_%dat% (4).pdf" "Bank Transaction Report_%dat% (Blt).pdf"
ren "Bank Transaction Report_%dat% (5).pdf" "Bank Transaction Report_%dat% (Blt)(1).pdf"
ren "Complimentary Rooms Report_%dat%.pdf" "Complimentary Rooms Report_%dat% (Inhouse).pdf"
ren "Complimentary Rooms Report_%dat% (1).pdf" "Complimentary Rooms Report_%dat% (All).pdf"
ren "Covers Report_%dat%.pdf" "Covers Report_%dat% (Detail YTD).pdf"
ren "Covers Report_%dat% (1).pdf" "Covers Report_%dat% (Detail).pdf"
ren "Covers Report_%dat% (2).pdf" "Covers Report_%dat% (Summary).pdf"
ren "Daily Revenue Report_%dat% (2).pdf" "Daily Revenue Report_%dat% (Inc YTD Budget).pdf"
ren "Daily Revenue Report_%dat% (3).pdf" "Daily Revenue Report_%dat% (2).pdf"
ren "Daily Revenue Report_%dat% (4).pdf" "Daily Revenue Report_%dat% (3).pdf"
ren "Detail Ticket Report_%dat%.pdf" "Detail Ticket Report_%dat% (Dep_All Sub_All).pdf"
ren "Detail Ticket Report_%dat% (1).pdf" "Detail Ticket Report_%dat% (Dep_98 Sub_1).pdf"
ren "Detail Ticket Report_%dat% (2).pdf" "Detail Ticket Report_%dat% (Dep_98 Sub_1)(1).pdf"
ren "Detail Ticket Report_%dat% (3).pdf" "Detail Ticket Report_%dat% (Dep_98 Sub_24).pdf"
ren "Detail Ticket Report_%dat% (4).pdf" "Detail Ticket Report_%dat% (Dep_98 Sub_24)(1).pdf"
ren "Detail Ticket Report_%dat% (5).pdf" "Detail Ticket Report_%dat% (Dep_All Sub_All)(1).pdf"
ren "Detail Ticket Report_%dat% (6).pdf" "Detail Ticket Report_%dat% (Dep_All Sub_All)(2).pdf"
ren "Detail Ticket Report_%dat% (7).pdf" "Detail Ticket Report_%dat% (Dep_5 Sub_All).pdf"
ren "Detail Ticket Report_%dat% (8).pdf" "Detail Ticket Report_%dat% (Dep_6 Sub_All).pdf"
ren "Detail Ticket Report_%dat% (9).pdf" "Detail Ticket Report_%dat% (Dep_7 Sub_All).pdf"
ren "Detail Ticket Report_%dat% (10).pdf" "Detail Ticket Report_%dat% (Dep_All Sub_All)(3).pdf"
ren "Detail Ticket Report_%dat% (11).pdf" "Detail Ticket Report_%dat% (Dep_1to87 Sub_51to60).pdf"
ren "Detail Ticket Report_%dat% (12).pdf" "Detail Ticket Report_%dat% (Dep_15 Sub_All).pdf"
ren "Detail Ticket Report_%dat% (13).pdf" "Detail Ticket Report_%dat% (Dep_1to87 Sub_51to99).pdf"
ren "Detail Ticket Report_%dat% (14).pdf" "Detail Ticket Report_%dat% (Dep_1 Sub_95).pdf"
ren "Detail Ticket Report_%dat% (15).pdf" "Detail Ticket Report_%dat% (Dep_1to99 Sub_50to99).pdf"
ren "Guest Ledger Summary Report_%dat%.pdf" "Guest Ledger Summary Report_%dat% (By Name).pdf"
ren "Guest Ledger Summary Report_%dat% (1).pdf" "Guest Ledger Summary Report_%dat% (By Room).pdf"
ren "Guest Ledger Summary Report_%dat% (2).pdf" "Guest Ledger Summary Report_%dat% (By Room)(1).pdf"
ren "High Balance Report_%dat%.pdf" "High Balance Report_%dat% (Exceed Or Within 150).pdf"
ren "High Balance Report_%dat% (1).pdf" "High Balance Report_%dat% (All).pdf"
ren "In House Guest Report_%dat%.pdf" "In House Guest Report_%dat% (All Inhouse).pdf"
ren "In House Guest Report_%dat% (1).pdf" "In House Guest Report_%dat% (Reg Only).pdf"
ren "Managers Statistics Report_%dat% (1).pdf" "Managers Statistics Report_%dat% (Summary).pdf"
ren "Room Rate Change Report_%dat% (1).pdf" "Room Rate Change Report_%dat% (Inhouse Only).pdf" 
ren "Special Services Report_%dat%.pdf" "Special Services Report (T5)_%dat% (with Comments).pdf" 
ren "Special Services Report_%dat% (1).pdf" "Special Services Report (T5)_%dat%.pdf" 
ren "Special Services Report_%dat% (2).pdf" "Special Services Report (J8)_%dat%.pdf" 
ren "Special Services Report_%dat% (3).pdf" "Special Services Report (PARK)_%dat%.pdf" 
ren "Special Services Report_%dat% (4).pdf" "Special Services Report (Ts)_%dat%.pdf" 
ren "Special Services Report_%dat% (5).pdf" "Special Services Report (SCRE)_%dat%.pdf" 
ren "Special Services Report_%dat% (6).pdf" "Special Services Report (EFE)_%dat%.pdf" 
cd..
goto :EOF

:auditPack
cd OutputFolder
md "AuditPack" 2>nul
copy "Advance Deposit Balance Sheet_%dat% (1).pdf" "AuditPack/Advance Deposit Balance Sheet_%dat%.pdf" >nul
copy "AR Summary Report_%dat% (1).pdf" "AuditPack/AR Summary Report_%dat%.pdf" >nul
copy "Accounts Receivable Validation Report_%dat%.pdf" "AuditPack/AR Validation Report_%dat%.pdf" >nul
copy "Complimentary Rooms Report_%dat% (All).pdf" "AuditPack/Complimentary Rooms Report_%dat%.pdf" >nul
copy "Covers Report_%dat% (Detail).pdf" "AuditPack/Covers Report_%dat%.pdf" >nul
copy "Daily Cash Out Report_%dat% (3).pdf" "AuditPack/Daily Cash Out Report_%dat%.pdf" >nul
copy "Daily Revenue Report_%dat% (3).pdf" "AuditPack/Daily Revenue Report_%dat%.pdf" >nul
copy "Detail Ticket Report_%dat% (Dep_All Sub_All)(3).pdf" "AuditPack/Detail Ticket Report_%dat%.pdf" >nul
copy "Detail Ticket Report_%dat% (Dep_1to87 Sub_51to99).pdf" "AuditPack/Detail Adjustments Report_%dat%.pdf" >nul
copy "Expected Arrival Report_%dat% (1).pdf" "AuditPack/Expected Arrival Report_%dat%.pdf" >nul
copy "Guest History Exception Report_%dat%.pdf" "AuditPack/Guest History Exception Report_%dat%.pdf" >nul
copy "Guest Ledger Summary Report_%dat% (By Room)(1).pdf" "AuditPack/Guest Ledger Summary Report_%dat%.pdf" >nul
copy "High Balance Report_%dat% (Exceed Or Within 150).pdf" "AuditPack/High Balance Report_%dat%.pdf" >nul
copy "Managers Statistics Report_%dat%.pdf" "AuditPack/Managers Statistics Report_%dat%.pdf" >nul
copy "Market Segment Analysis_%dat% (1).pdf" "AuditPack/Market Segment Analysis_%dat%.pdf" >nul
copy "No Show Report_%dat%.pdf" "AuditPack/No Show Report_%dat%.pdf" >nul
copy "Open Folio System Balancing Report_%dat% (1).pdf" "AuditPack/Open Folio System Balancing Report_%dat%.pdf" >nul
copy "Out Of Order Report_%dat%.pdf" "AuditPack/Out Of Order Report_%dat%.pdf" >nul
copy "Reservation Activity Report_%dat%.pdf" "AuditPack/Reservation Activity Report_%dat%.pdf" >nul
copy "Room Post Audit Report_%dat% (1).pdf" "AuditPack/Room Post Audit Report_%dat%.pdf" >nul
copy "Room Rate Change Report_%dat%.pdf" "AuditPack/Room Rate Change Report_%dat%.pdf" >nul
copy "Special Services Report (EFE)_%dat%.pdf" "AuditPack/Special Services Report (EFE)_%dat%.pdf" >nul
copy "Special Services Report (SCRE)_%dat%.pdf" "AuditPack/Special Services Report (SCRE)_%dat%.pdf" >nul
copy "Special Services Report (T5)_%dat% (with Comments).pdf" "AuditPack/Special Services Report (T5)_%dat%.pdf" >nul
copy "Special Services Report (Ts)_%dat%.pdf" "AuditPack/Special Services Report (Ts)_%dat%.pdf" >nul
copy "VIP Report_%dat%.pdf" "AuditPack/VIP Report_%dat%.pdf" >nul
set "auditPackLoc=%cd%\AuditPack\"
cd..

goto :EOF

:makeReport
rem Pulls Bank Balance Sheet Numbers from Daily Cash Out Report
(
tools\xpdf\pdfinfo.exe "OutputFolder\Daily Cash Out Report_%dat% (3).pdf"
) > temp.txt
for /F "tokens=1-2 delims= " %%A in ('find "Pages" "temp.txt"') do set /a "pageNum=%%B"
set /a "pageNum=%pageNum%-1"
tools\xpdf\pdftotext.exe -raw -f %pageNum% -l %pageNum% "OutputFolder\Daily Cash Out Report_%dat% (3).pdf"
for /F "tokens=1-5 delims= " %%A in ('find "704924 AX American Express" "OutputFolder\Daily Cash Out Report_%dat% (3).txt"') do set "dcoAxSetl=%%E"
set dcoAxSetl=%dcoAxSetl:,=%
if "%dcoAxSetl%" == "REPORT_%dat%" ( set "dcoAxSetl=0.00" )
for /F "tokens=1-4 delims= " %%A in ('find "704924 VI Visa" "OutputFolder\Daily Cash Out Report_%dat% (3).txt"') do set "dcoViSetl=%%D"
set dcoViSetl=%dcoViSetl:,=%
if "%dcoViSetl%" == "OUT" ( set "dcoViSetl=0.00" )
for /F "tokens=1-4 delims= " %%A in ('find "704924 MC MasterCard" "OutputFolder\Daily Cash Out Report_%dat% (3).txt"') do set "dcoMcSetl=%%D"
set dcoMcSetl=%dcoMcSetl:,=%
if "%dcoMcSetl%" == "OUT" ( set "dcoMcSetl=0.00" )
for /F "tokens=1-5 delims= " %%A in ('find "704924 DI Discover" "OutputFolder\Daily Cash Out Report_%dat% (3).txt"') do set "dcoDiSetl=%%E"
set dcoDiSetl=%dcoDiSetl:,=%
if "%dcoDiSetl%" == "REPORT_%dat%" ( set "dcoDiSetl=0.00" )
del "OutputFolder\Daily Cash Out Report_%dat% (3).txt"

rem Pulls Bank Balance Sheet Numbers from Daily Revenue Report
tools\xpdf\pdftotext.exe -raw "OutputFolder/Daily Revenue Report_%dat% (3).pdf"
(
find "Deposit Rcvd" "OutputFolder\Daily Revenue Report_%dat% (3).txt"
) > temp.txt
del "OutputFolder\Daily Revenue Report_%dat% (3).txt"
for /F "tokens=1-5 delims= " %%A in ('find "Deposit Rcvd - AX" "temp.txt"') do set "drrAxDep=%%E"
for /F "tokens=1-5 delims= " %%A in ('find "Deposit Rcvd - VI" "temp.txt"') do set "drrViDep=%%E"
for /F "tokens=1-5 delims= " %%A in ('find "Deposit Rcvd - MC" "temp.txt"') do set "drrMcDep=%%E"
for /F "tokens=1-5 delims= " %%A in ('find "Deposit Rcvd - DI" "temp.txt"') do set "drrDiDep=%%E"

rem Pulls Bank Balance Sheet Numbers from Bank Transaction Report
tools\xpdf\pdftotext.exe -raw "OutputFolder/Bank Transaction Report_%dat% (Blt)(1).pdf"
(
find "*" "OutputFolder\Bank Transaction Report_%dat% (Blt)(1).txt"
) > temp.txt
del "OutputFolder\Bank Transaction Report_%dat% (Blt)(1).txt"
for /F "tokens=1-10 delims= " %%A in ('find "* Total Bank Amount Deposited for AX on " "temp.txt"') do set "btrAxSetl=%%J"
if "%btrAxSetl%" == "" ( set "btrAxSetl=0.00" )
for /F "tokens=1-10 delims= " %%A in ('find "* Total Bank Amount Deposited for VI on " "temp.txt"') do set "btrViSetl=%%J"
if "%btrViSetl%" == "" ( set "btrViSetl=0.00" )
for /F "tokens=1-10 delims= " %%A in ('find "* Total Bank Amount Deposited for MC on " "temp.txt"') do set "btrMcSetl=%%J"
if "%btrMcSetl" == "" ( set "btrMcSetl=0.00" )
for /F "tokens=1-10 delims= " %%A in ('find "* Total Bank Amount Deposited for DI on " "temp.txt"') do set "btrDiSetl=%%J"
if "%btrDiSetl%" == "" ( set "btrDiSetl=0.00" )

rem Pulls Statistics from Managers Statistics Report
tools\xpdf\pdftotext.exe -raw -l 1 "OutputFolder\Managers Statistics Report_%dat% (Summary).pdf" "temp.txt"
setlocal enableextensions enabledelayedexpansion
copy NUL temp2.txt > NUL
set /A maxlines=26
set /A linecount=0
for /F "delims=" %%A in (temp.txt) do ( 
  if !linecount! GEQ %maxlines% goto ExitLoop
  echo %%A >> temp2.txt
  set /A linecount+=1
)
:ExitLoop
sort /R "temp2.txt" /O "temp.txt"
del temp2.txt
for /F "tokens=1-3 delims= " %%A in ('find "Available Rooms " "temp.txt"') do set "avlRooms=%%C"
for /F "tokens=1-4 delims= " %%A in ('find "Rooms Occupied (Paid) " "temp.txt"') do set "occRooms=%%D"
rem set /a occPer=100*(!occRooms! / !avlRooms!)
for /F "tokens=1-3 delims= " %%A in ('find "Rev Par " "temp.txt"') do set "revPar=%%C"
for /F "tokens=1-3 delims= " %%A in ('find "Room Revenue " "temp.txt"') do set "roomRev=%%C"
for /F "tokens=1-3 delims= " %%A in ('find "All Revenue " "temp.txt"') do set "totRev=%%C"
for /F "tokens=1-2 delims= " %%A in ('find "Arrivals " "temp.txt"') do set "arrivals=%%B"
for /F "tokens=1-2 delims= " %%A in ('find "Departures " "temp.txt"') do set "departures=%%B

rem Retrieve IRD Food Total
tools\xpdf\pdftotext.exe -raw -l 1 "OutputFolder/Daily Revenue Report_%dat% (3).pdf"
for /F "tokens=1-4 delims= " %%A in ('find "Food - Breakfast" "OutputFolder/Daily Revenue Report_%dat% (3).txt"') do set "irdBrfkst=%%D"
set irdBrfkst=%irdBrfkst:.=%
for /F "tokens=1-4 delims= " %%A in ('find "Food - Lunch" "OutputFolder/Daily Revenue Report_%dat% (3).txt"') do set "irdLunch=%%D"
set irdLunch=%irdLunch:.=%
for /F "tokens=1-4 delims= " %%A in ('find "Food - Dinner" "OutputFolder/Daily Revenue Report_%dat% (3).txt"') do set "irdDinn=%%D"
set irdDinn=%irdDinn:.=%
for /F "tokens=1-5 delims= " %%A in ('find "Food - Late Night" "OutputFolder/Daily Revenue Report_%dat% (3).txt"') do set "irdLate=%%E"
set irdLate=%irdLate:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Beer" "OutputFolder/Daily Revenue Report_%dat% (3).txt"') do set "irdBeer=%%B"
set irdBeer=%irdBeer:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Wine" "OutputFolder/Daily Revenue Report_%dat% (3).txt"') do set "irdWine=%%B"
set irdWine=%irdWine:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Liquor" "OutputFolder/Daily Revenue Report_%dat% (3).txt"') do set "irdLiquor=%%B"
set irdLiquor=%irdLiquor:.=%
set /A irdTot=!irdBrfkst!+!irdLunch!+!irdDinn!+!irdLate!+!irdBeer!+!irdWine!+!irdLiquor!
set irdTot=%irdTot:~0,-2%.%irdTot:~-2%
del "OutputFolder\Daily Revenue Report_%dat% (3).txt"

rem Retrieve T54 Food Total
tools\xpdf\pdftotext.exe -raw -f 2 -l 2 "OutputFolder/Daily Revenue Report_%dat% (3).pdf" "temp.txt"
setlocal enableextensions enabledelayedexpansion
copy NUL temp2.txt > NUL
copy NUL temp3.txt > NUL
set /A maxlines=36
set /A linecount=0
set /A line19=19
set /A line20=20
set /A line21=21
set /A line22=22
set /A line23=23
set /A line24=24
set /A line31=31
set /A line32=32
set /A line33=33
set /A line34=34
for /F "delims=" %%A in (temp.txt) do ( 
  if !linecount! GEQ %maxlines% goto ExitLoop
  if !linecount! EQU %line19% echo %%A >> temp2.txt 
  if !linecount! EQU %line20% echo %%A >> temp2.txt 
  if !linecount! EQU %line21% echo %%A >> temp2.txt 
  if !linecount! EQU %line22% echo %%A >> temp2.txt 
  if !linecount! EQU %line23% echo %%A >> temp2.txt 
  if !linecount! EQU %line24% echo %%A >> temp2.txt 
  if !linecount! EQU %line31% echo %%A >> temp3.txt 
  if !linecount! EQU %line32% echo %%A >> temp3.txt 
  if !linecount! EQU %line33% echo %%A >> temp3.txt 
  if !linecount! EQU %line34% echo %%A >> temp3.txt 
  set /A linecount+=1
)
:ExitLoop
for /F "tokens=1-4 delims= " %%A in ('find "Food - Breakfast" "temp2.txt"') do set "t54Brfkst=%%D"
set t54Brfkst=%t54Brfkst:.=%
for /F "tokens=1-4 delims= " %%A in ('find "Food - Lunch" "temp2.txt"') do set "t54Lunch=%%D"
set t54Lunch=%t54Lunch:.=%
for /F "tokens=1-4 delims= " %%A in ('find "Food - Dinner" "temp2.txt"') do set "t54Dinner=%%D"
set t54Dinner=%t54Dinner:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Beer" "temp2.txt"') do set "t54Beer=%%B"
set t54Beer=%t54Beer:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Wine" "temp2.txt"') do set "t54Wine=%%B"
set t54Wine=%t54Wine:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Liquor" "temp2.txt"') do set "t54Liquor=%%B"
set t54Liquor=%t54Liquor:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Food" "temp3.txt"') do set "patFood=%%B"
set patFood=%patFood:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Beer" "temp3.txt"') do set "patBeer=%%B"
set patBeer=%patBeer:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Wine" "temp3.txt"') do set "patWine=%%B"
set patWine=%patWine:.=%
for /F "tokens=1-2 delims= " %%A in ('find "Liquor" "temp3.txt"') do set "patLiquor=%%B"
set patLiquor=%patLiquor:.=%
set /A t54Tot=!t54Brfkst!+!t54Lunch!+!t54Dinner!+!t54Beer!+!t54Wine!+!t54Liquor!+!patFood!+!patBeer!+!patWine!+!patLiquor!
set t54Tot=%t54Tot:~0,-2%.%t54Tot:~-2%

del temp.txt
del temp2.txt
del temp3.txt
(
echo WarpSpeed: Night Audit Accelerator by John Dudek v%ver%
echo.
echo WarpSpeed Report:
echo ------------------------------------------- 
echo Operation Completed on %date% at %time% 
echo.
echo Audit Pack Back-Up Location: %auditPackLoc%
echo.
echo WarpSpeed Report Location: %cd%\WS_Report.txt
echo.
echo +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
echo.
echo Bank Balance Sheet Information:
echo =============================== 
echo Daily Cash Out Settlement Totals -
echo -------------------------------------------
echo American Express: %dcoAxSetl%
echo.
echo Visa: %dcoViSetl%
echo.
echo MasterCard: %dcoMcSetl%
echo.
echo Discover: %dcoDiSetl%
echo.
echo.
echo Daily Revenue Report Deposit Totals -
echo -------------------------------------------
echo American Express: %drrAxDep%
echo.
echo Visa: %drrViDep%
echo.
echo MasterCard: %drrMcDep%
echo.
echo Discover: %drrDiDep%
echo.
echo.
echo Bank Transaction Report Settlement Totals -
echo -------------------------------------------
echo American Express: %btrAxSetl%
echo.
echo Visa: %btrViSetl%
echo.
echo MasterCard: %btrMcSetl%
echo.
echo Discover: %btrDiSetl%
echo.
echo.
echo.
echo Ops Report Statistics:
echo ===============================
echo Occupied Rooms: !occRooms!
echo Occupancy: !occRooms! / !avlRooms!    //Not Retrievable by WarpSpeed
echo RevPAR: !revPar!
echo Room Rev: !roomRev!
echo Total Rev: !totRev!
echo Arrivals: !arrivals!
echo Departures: !departures!
echo.
echo. 
echo.
echo Hotel Effectiveness Night Audit Entry:
echo ===============================
echo Occupied Rooms: !occRooms!
echo Room Revenue: !roomRev!
echo Departures: !departures!
echo In Room Dining: !irdTot! 
echo Restaurant Revenue: !t54Tot!
echo.
echo.  
) > WS_Report.txt
echo Complete.

set "endTime=%time: =0%"
rem Get elapsed time:
set "end=!endTime:%time:~8,1%=%%100)*100+1!"  &  set "start=!startTime:%time:~8,1%=%%100)*100+1!"
set /A "elap=((((10!end:%time:~2,1%=%%100)*60+1!%%100)-((((10!start:%time:~2,1%=%%100)*60+1!%%100), elap-=(elap>>31)*24*60*60*100"
rem Convert elapsed time to HH:MM:SS:CC format:
set /A "cc=elap%%100+100,elap/=100,ss=elap%%60+100,elap/=60,mm=elap%%60+100,hh=elap/60+100"

echo.
echo ------------------------------------------- 
echo.
echo Total elapsed time: %mm:~1% minute(s) and %ss:~1%%time:~8,1%%cc:~1% seconds
echo.
echo Audit Pack Back-Up Location: %auditPackLoc%
echo.
echo WarpSpeed Report Location: %cd%\WS_Report.txt
echo.
echo ------------------------------------------- 
echo.
echo Have a great rest of your shift^^!
echo.

start notepad "WS_Report.txt" 

goto :EOF

:getAuditDate
rem Night Audit Date detection
set month=%date:~4,2%
set day=%date:~7,2%
set year=%date:~10,4%
if %day% LEQ 9 (
	set day=%day:~1,1%
)
set /a auditDay="%day%"-1
rem Leap Year Calculation
set /a leapConA = "%year%" %% 4
set /a leapConB = "%year%" %% 100
set /a leapConC = "%year%" %% 400
if "%leapConA%"=="0" (
	if "%leapConB%" NEQ "0" (
		set leap=1
		)
	)
if "%leapConC%"=="0" (
	set leap=1
	)
rem Accounts for the current day being the first day of the month
if "%day%"=="01" (
  if "%month%"=="01" (
    set nameMonth=December
	set month=12
	set auditDay=31
	set /a year=%year%-1
  )
    if "%month%"=="02" (
    set nameMonth=January
	set month=01
	set auditDay=31
  )
    if "%month%"=="03" (
    set nameMonth=February
	set month=02
	set auditDay=28		
	if "%leap%"=="1" (
		set auditDay=29
	)	
  )
    if "%month%"=="04" (
	set nameMonth=March
	set month=03
	set auditDay=31
  )
    if "%month%"=="05" (
    set nameMonth=April
	set month=04
	set auditDay=30
  )
    if "%month%"=="06" (
    set nameMonth=May
	set month=05
	set auditDay=31
  )
    if "%month%"=="07" (
    set nameMonth=June
	set month=06
	set auditDay=30
  )
    if "%month%"=="08" (
    set nameMonth=July
	set month=07
	set auditDay=31
  )
    if "%month%"=="09" (
    set nameMonth=August
	set month=08
	set auditDay=31
  )
    if "%month%"=="10" (
    set nameMonth=September
	set month=09
	set auditDay=30
  )
    if "%month%"=="11" (
    set nameMonth=October
	set month=10
	set auditDay=31
  )
    if "%month%"=="12" (
    set nameMonth=November
	set month=11
	set auditDay=30
  )
) else (
  if "%month%"=="01" set nameMonth=January
  if "%month%"=="02" set nameMonth=February
  if "%month%"=="03" set nameMonth=March
  if "%month%"=="04" set nameMonth=April
  if "%month%"=="05" set nameMonth=May
  if "%month%"=="06" set nameMonth=June
  if "%month%"=="07" set nameMonth=July
  if "%month%"=="08" set nameMonth=August
  if "%month%"=="09" set nameMonth=September
  if "%month%"=="10" set nameMonth=October
  if "%month%"=="11" set nameMonth=November
  if "%month%"=="12" set nameMonth=December
)
rem Removes leading 0 on single digit information
if %month% LEQ 9 (
	set month=%month:~1,1%
)
set dat=%month%-%auditDay%-%year%
goto :EOF

:warpDrive
rem WarpDrive Package Detector
set wdVer=1.2
cls
echo WarpDrive: Room Package Detector by John Dudek
echo ##############################################
echo # __      __             ___      _          #
echo # \ \    / /_ _ _ _ _ __^|   \ _ _^(_^)_ _____  #
echo #  \ \/\/ / _^` ^| ^'_^| ^'_ \ ^|^) ^| ^'_^| \ V / -_^) #
echo #   \_/\_/\__,_^|_^| ^| .__/___/^|_^| ^|_^|\_/\___^| #
echo #                  ^|_^|                 v%wdVer%  #
echo #                                            #
echo ##############################################
echo.
echo Which packages would you like to find?:
echo  1. All Packages
echo  2. Valet Packages
echo.
set /p input=" Please make a selection (1|2): "
if %input% EQU 1 (
	set wdKey="PKG_LIST"
	) else (
	if %input% EQU 2 (
		set wdKey="Valet"
		) else (
		echo.
		echo Invalid Input.
		echo.
		pause
		goto :warpDrive
		)
	)
)
setlocal EnableDelayedExpansion
set "startTime=%time: =0%"

echo.
echo Scanning for Packages . . .                                   
(
echo WarpDrive: Room Package Detector by John Dudek v%wdVer%
echo.
echo WarpDrive Report:
echo ------------------------------------------- 
echo Operation Completed on %date% at %time%
echo.
echo WarpDrive Report Location: %cd%\WD_Report.txt
echo.
echo +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
) > WD_Report.txt
tools\xpdf\pdftotext.exe -raw -nopgbrk "%wdRep%.pdf"
FINDSTR /M /C:"%wdKey%" "tools\packages\*.txt" > temp.txt
for /F "delims=" %%A in (temp.txt) do (
	set packDesc=0
	for /F "skip=1 delims=" %%B IN (%%A) DO if !packDesc! EQU 0 set "packDesc=%%B"
	FINDSTR /i /L /g:"%%A" "%wdRep%.txt" > temp2.txt
	if !errorlevel! EQU 0 echo.>>WD_Report.txt
	if !errorlevel! EQU 0 echo !packDesc!>>WD_Report.txt
	if !errorlevel! EQU 0 echo ===============================>>WD_Report.txt
	for /F "delims=" %%C in (temp2.txt) do (
		set packRoom=%%C
		set packRoom=!packRoom:~0,4!
		set packRoom=!packRoom: =!
		set /a roomCheck=packRoom/1
		if "!roomCheck!" NEQ "!packRoom!" (
		echo Package with No Assigned Room>>WD_Report.txt
		) else (
		echo !packRoom!>>WD_Report.txt)
	)
)
set "endTime=%time: =0%"
rem Get elapsed time:
set "end=!endTime:%time:~8,1%=%%100)*100+1!"  &  set "start=!startTime:%time:~8,1%=%%100)*100+1!"
set /A "elap=((((10!end:%time:~2,1%=%%100)*60+1!%%100)-((((10!start:%time:~2,1%=%%100)*60+1!%%100), elap-=(elap>>31)*24*60*60*100"
rem Convert elapsed time to HH:MM:SS:CC format:
set /A "cc=elap%%100+100,elap/=100,ss=elap%%60+100,elap/=60,mm=elap%%60+100,hh=elap/60+100"

del %wdRep%.txt
del temp.txt
del temp2.txt
start notepad "WD_Report.txt"
echo Complete.
echo.
echo ------------------------------------------- 
echo.
echo Total elapsed time: %mm:~1% minute(s) and %ss:~1%%time:~8,1%%cc:~1% seconds
echo.
echo WarpDrive Report Location: %cd%\WD_Report.txt
echo.
echo ------------------------------------------- 
echo.
echo Have a great rest of your shift^^!
echo.
setlocal DisableDelayedExpansion
goto :EOF

exit
