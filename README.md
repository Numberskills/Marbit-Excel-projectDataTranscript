# Marbit-Excel-projectDataTranscript
Excel add on for transcribing data to PC/PL-files monthly, and automatically creating relations in follow up file.

This Excel add-on has been created as a temporary ad-hoc solution to two manual and laborous activities:
  Creation of relations in the follow-up file
  Setting monthly base data in PL/PC files from financial data files

Scripts intially created by Michael Reinholtz

-------------------------------

v1.0
Functionality
-------------
Creating relations in the follow-up file
* Dialog for user to choose source library
* Setting source file name (file path dialog, PL/PC name in column A, source sheet dependent on ActiveSheet)
* Setting source row relation dependant on Selection.Row
* Setting source relation formulas based on Selection.Address, allowing for Range selection (multiple cells in the same sheet)

Instructions
------------
Instructions for installation and usage can be found at https://youtu.be/ix-hlASL3Dk

Known issues
------------
* Set to actual max number of PC/PL (no scaleability)
* Source path cannot contain single quotes (')
* PL/PC name in column A must correlate exactly to source file name (i.e. "PC Michael Reinholtz" -> "[...]\Projektuppföljning 2022 PC Michael Reinholtz.xlsx")
* PL/PC files must be closed when running the script

-------------------------------
