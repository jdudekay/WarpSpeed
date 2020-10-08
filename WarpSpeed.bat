@echo off
rem    WarpSpeed: Night Audit Accelerator - A tool made with the intention of streamlining the renmaming a large amount of randomly-named .pdf files based on their contents
rem    Copyright (C) 2020 John Dudek
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
echo +-+-+-                                                   v 1.2.0                                                  -+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-                                       Copyright (C) 2020 John Dudek                                        -+-+-+
echo -+-+-+                                                                                                            +-+-+-
echo +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
echo -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
echo +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
pause
cls

setlocal EnableExtensions DisableDelayedExpansion
set /a count=1
set /a totFile=0
set /a dupFile=0
set /a illChar=0

rem Initialization of OutputFolder and main program prompts
md "OutputFolder" 2>nul
echo WarpSpeed: Night Audit Accelerator by John Dudek v1.2.0
echo. 
set /p dat="What is the date you are doing Night Audit for? (MM.DD.YY): "

rem Main program loop
for /F %%I in ('dir *.pdf /B') do call :RenamePDF "%%~fI"
echo Initial renaming and moving sequence complete.
pause

rem Cleanup subroutine
echo Cleaning up file names . . .
cd OutputFolder
call :cleanUp
pause

rem Pause before program exit
cls
set /a totFile=%totFile%-%illChar%
echo WarpSpeed: Night Audit Accelerator by John Dudek v1.2.0
echo. 
echo Operation complete.
echo.
echo Statistics Report:
echo ----------------------------------------
echo Illegal Characters Corrected: %illChar% 
echo.
echo Total Files Renamed and Moved: %totFile%
echo.
pause
goto :EOF

:RenamePDF
set "FilePDF=%~1"
set "FileTXT=%~dpn1.txt"

rem Create text version of .pdf to be read
rem Parameters are set to print only the first page and in raw which condenses the data
pdftotext.exe -l 1 -raw "%FilePDF%"

rem For loop is used to call find which selects the line that contains the report name
rem Names
for /F "tokens=1 delims=:" %%A in ('%SystemRoot%\System32\find.exe "Westin Medical Ctr" "%FileTXT%"') do set "report=%%A"

rem The text version of the PDF file is no longer needed.
del "%FileTXT%"

rem Trim unneccessary words from file name variable
rem Creates actual fileName variable from user inputted and computer derived variables
set report=%report:Westin Medical Ctr =%
set report=%report: Page Number=%
set "fileName=%report%_%dat%"

rem Move the PDF file to OutputFolder and rename the file while moving it.
rem Also checks to make sure no files are overwritten
:while
if not exist "OutputFolder\%fileName%.pdf" (
move "%FilePDF%" "OutputFolder\%fileName%.pdf"
set /a count=1
set /a totFile=%totFile%+1
) else (
set "fileName=%report%_%dat% (%count%)"
set /a count=%count%+1
goto :while
)

rem AR reports have the illegal "/" character, this removes it
rem Also checks to make sure no files are overwritten
if errorlevel 1 set report=%report:A/R =%
if errorlevel 1 set "fileName=AR %report%_%dat%"
if errorlevel 1 echo Illegal Character detected and removed.
if errorlevel 1 set /a illChar=%illChar%+1
if errorlevel 1 (
:whileAR
if not exist "OutputFolder\%fileName%.pdf" (
move "%FilePDF%" "OutputFolder\%fileName%.pdf"
set /a count=1
set /a totFile=%totFile%+1
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
ren "Covers Report_%dat%.pdf" "Covers Report_%dat% (Detail All Outlets).pdf"
ren "Covers Report_%dat% (1).pdf" "Covers Report_%dat% (Detail T54).pdf"
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
rem ren "Expected Arrival Report_%dat%.pdf" "Expected Arrival Report_%dat%.pdf"
rem ren "Expected Arrival Report_%dat% (1).pdf" "Expected Arrival Report_%dat% (1).pdf"
rem ren "Expected Arrival Report_%dat% (2).pdf" "Expected Arrival Report_%dat% (2).pdf"
rem ren "Expected Arrival Report_%dat% (3).pdf" "Expected Arrival Report_%dat% (3).pdf"
rem ren "Expected Arrival Report_%dat% (4).pdf" "Expected Arrival Report_%dat% (4).pdf"
ren "Guest Ledger Summary Report_%dat%.pdf" "Guest Ledger Summary Report_%dat% (By Name).pdf"
ren "Guest Ledger Summary Report_%dat% (1).pdf" "Guest Ledger Summary Report_%dat% (By Room).pdf"
ren "Guest Ledger Summary Report_%dat% (2).pdf" "Guest Ledger Summary Report_%dat% (By Room)(1).pdf"
ren "High Balance Report_%dat%.pdf" "High Balance Report_%dat% (Exceed Or Within 150).pdf"
ren "High Balance Report_%dat% (1).pdf" "High Balance Report_%dat% (All).pdf"
ren "In House Guest Report_%dat%.pdf" "In House Guest Report_%dat% (All Inhouse).pdf"
ren "In House Guest Report_%dat% (1).pdf" "In House Guest Report_%dat% (Reg Only).pdf"
ren "Managers Statistics Report_%dat% (1).pdf" "Managers Statistics Report_%dat% (Summary).pdf"
ren "Room Rate Change Report_%dat% (1).pdf" "Room Rate Change Report_%dat% (Inhouse Only).pdf" 
ren "Special Services Report_%dat%.pdf" "T5 Special Services Report_%dat% (with Comments).pdf" 
ren "Special Services Report_%dat% (1).pdf" "T5 Special Services Report_%dat%.pdf" 
ren "Special Services Report_%dat% (2).pdf" "J8 Special Services Report_%dat%.pdf" 
ren "Special Services Report_%dat% (3).pdf" "PARK Special Services Report_%dat%.pdf" 
ren "Special Services Report_%dat% (4).pdf" "Ts Special Services Report_%dat%.pdf" 
ren "Special Services Report_%dat% (5).pdf" "SCRE Special Services Report_%dat%.pdf" 
ren "Special Services Report_%dat% (6).pdf" "EFE Special Services Report_%dat%.pdf" 
goto :EOF
