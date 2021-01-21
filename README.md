# Welcome to WarpSpeed!
 
**WarpSpeed** was created with the intention of streamlining the Night Audit process by means of automatically identifying and sorting all 98 Night Audit reports and extracting key information to be compiled in a convenient **WarpSpeed** report. 

Using **WarpSpeed** is designed to be as simple as possible, and only requires a few easy steps:

## Step One
### Prep **WarpSpeed** folder

Make sure the **WarpSpeed** folder is empty except for the following files:
* WarpSpeed.bat
* README.md
* LICENSE
* /tools/
 
If there is ANYTHING other than the above listed files/folders in the **WarpSpeed** folder on the Desktop, delete them so that the folder contents match the above

## Step Two
### Run **WarpSpeed**
 
The next step once you’ve gotten the folder ready for processing, run **WarpSpeed**, at the Title Screen it will ask you to press any key to continue after which WarpSpeed will check for updates and then attempt to automatically assess what date you are trying to do Audit for. If this date is incorrect for any reason, type "n". Advanced users can type "manual" and will be able to manually enter the data that way.
 
**WarpSpeed** will run automatically through all of its functions, ending with generating a report. The Audit Pack and WS_Report's locations will be displayed on the Operation Complete screen, and the **WarpSpeed** report itself which contains additional information for the Night Audit will be displayed as the program closes.

## Step Three
### Copy reports to Night Audit folder

As shown in the **WarpSpeed** report screen, the processed files will be placed in a folder called OutputFolder within the **WarpSpeed** folder and the actual Audit Pack will be placed in a folder called AuditPack within the OutputFolder folder. The only additional reports that need to be run for the Audit Pack will be the Occupational Forecast reports which need to be run manually.

## Step Four
### Retrieve Information from the **WarpSpeed** Report
 
The generated report within the **WarpSpeed** folder, WS_Report.txt will currently contain the following data:

- Bank Balance Sheet Totals
- Ops Report Statistics
- Hotel Effectiveness Night Audit Entry Statistics

As it currently stands there are no further statistics which **WarpSpeed** can reliably extract.

## Miscellaneous
### Updating WarpSpeed

At the main prompt asking which date you're doing Night Audit for, enter "update" (Without quotations) and the program will update itself to the latest released version and close itself.

### Error-Checking

WarpSpeed will perform basic error-checking and correcting functions upon startup to help mitigate errors, the complexity of which will increase over time.

If you have ANY questions, feel free to contact John Dudek at ANY time, 24/7 @ (631)745-0820 or jdudek@thewestinhouston.com

*WarpSpeed uses parts of Xpdf to more effectively maniuplate .pdf files.*
*The Xpdf software and documentation are copyright 1996-2019 Glyph & Cog, LLC.*

*WarpSpeed: Night Audit Accelerator Copyright (C) 2020 John Dudek*
