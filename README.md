# Welcome to WarpSpeed!

**WarpSpeed** was created with the intention of streamlining the Night Audit process by means of automatically identifying and sorting all ~~98~~ 99 Night Audit reports and extracting key information to be compiled in a convenient **WarpSpeed** report.

The scope of **WarpSpeed** has widened over time and is now also designed to analyze a list of guests, compare them against various lists of packages, and produce a report of said guests via a cheekily named **WarpDrive** subroutine.

Using **WarpSpeed** is designed to be as simple as possible, with a full explanation of main menu prompts given as follows:

## Night Audit Accelerator
**WarpSpeed** will automatically assess what date you are trying to do Audit for. Type "y" and **WarpSpeed** will run automatically through the following processes:
- Copying of .pdfs into OutputFolder (Folder will be created if it does not exist)
- Analysis of .pdfs
- Renaming and copying of .pdfs into AuditPack and AllReports folders
- Generation of NAA_Report.txt, bbsData.csv, and opsData.csv
- Copying of AuditPack folder into appropriate destination on Cloud Drive

The Audit Pack Back-Up and WS_Report's locations will be displayed on the Operation Complete screen, and the **WarpSpeed** report as well as the generated spreadsheets which contains additional information for the completion of Night Audit will be displayed before the program closes.

(Advanced users can type "debug"  at the prompt to use a user-defined group of Night Audit Reports)

## Room Package Detector (WarpDrive)
By generating either the Expected Arrivals or Actual Arrivals report in .csv format and placing it in the /WarpSpeed/ directory, you can use this function to compare the report against user-created Package Description files contained within the /WarpSpeed/tools/packages/ directory and generate a report indicating which, if any, of these guests had purchased a package that would require additional action from the front desk.
The user will be prompted to select either All Packages or just Valet Packages, after which they will be allowed to select a .csv file to analyze. Currently the Room Package Detector is only configured to analyze the Expected Arrivals report.

All Package Description files used by **WarpSpeed** will be .txt filetype and have the following format:

*PKG_LIST*
*NAME AND DESCRIPTION OF PACKAGE*
*99RATE*
*99RATE*
*etc...*

## Help
This menu is pretty self-explanatory, you can either open this ReadMe, or you can run **WarpSpeed**'s Update function.
When updating, **WarpSpeed** will first check to see what is the most updated version and display that information along-side the current version number so the user can choose to update or not.

### Error-Checking

WarpSpeed will perform basic error-checking and correcting functions upon startup to help mitigate errors, the complexity of which will increase over time.

If you have ANY questions, feel free to contact John Dudek at ANY time, 24/7 @ (631)745-0820 or jdudek@thewestinhouston.com

*WarpSpeed uses parts of Xpdf.*
*The Xpdf software and documentation are copyright 1996-2021 Glyph & Cog, LLC.*

*WarpSpeed: Night Audit Accelerator Copyright (C) 2020-2021 John Dudek*
