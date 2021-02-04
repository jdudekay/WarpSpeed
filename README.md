# Welcome to WarpSpeed!
 
**WarpSpeed** was created with the intention of streamlining the Night Audit process by means of automatically identifying and sorting all ~~98~~ 99 Night Audit reports and extracting key information to be compiled in a convenient **WarpSpeed** report. 

The scope of **WarpSpeed** has widened over time and is now also designed to analyze a list of guests, compare them against a list of packages, and produce a report of said guests via a cheekily named **WarpDrive** subroutine. (Documentation on how to use this is given in the Miscellaneous section below!)

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
 
**WarpSpeed** will run automatically through the following processes:
- Copying of .pdfs into current folder
- Analysis of .pdfs
- Renaming and copying of .pdfs into Back-Up folder and final Audit Pack folder
- Generation of WS_Report.txt

The Audit Pack Back-Up and WS_Report's locations will be displayed on the Operation Complete screen, and the **WarpSpeed** report itself which contains additional information for the Night Audit will be displayed before the program closes.

## Step Three
### Retrieve Information from the **WarpSpeed** Report
 
The generated report within the **WarpSpeed** folder, WS_Report.txt will currently contain the following data:

- Bank Balance Sheet Totals
- Ops Report Statistics
- Hotel Effectiveness Night Audit Entry Statistics

As it currently stands there are no further statistics which **WarpSpeed** can reliably extract.

## Miscellaneous
### WarpDrive

If **WarpSpeed** is started with a .pdf named *rmrtver.pdf* within the /WarpSpeed/ directory, **WarpDrive** will automatically engage and compare the list of guests contained within the .pdf to user-created Package Description files contained within the /WarpSpeed/tools/packages/ directory and will generate a report indicating which, if any, of these guests had purchased a package that would require additional action from the front desk.

All Package Description files used by **WarpSpeed** will be .txt filetype and have the following format:

PKG_LIST  
NAME AND DESCRIPTION OF PACKAGE  
99RATE  
99RATE  

### Updating WarpSpeed

At the main prompt asking which date you're doing Night Audit for, enter "update" (Without quotations) and the program will update itself to the latest released version and close itself.

### Error-Checking

WarpSpeed will perform basic error-checking and correcting functions upon startup to help mitigate errors, the complexity of which will increase over time.

If you have ANY questions, feel free to contact John Dudek at ANY time, 24/7 @ (631)745-0820 or jdudek@thewestinhouston.com

*WarpSpeed uses parts of Xpdf.*
*The Xpdf software and documentation are copyright 1996-2019 Glyph & Cog, LLC.*

*WarpSpeed: Night Audit Accelerator Copyright (C) 2020-2021 John Dudek*
