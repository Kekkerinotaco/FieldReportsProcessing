About all: The excel report formats are dependent of the field operator, and
can change with time, so i hope that the code ideas in this repository will be
useful for you, when you will face the challenge of processing and analyzing
field data.

The main idea of all scripts is to transform the reports into the format, which
can be easily used for analysis.

============================================

DailyProductionReports:

Input: A link to a folder with daily production reports in .xlsx format.

Output:.xlsx file with all the year data collected in one file.

Comments: The code is tuned to work with a certain format, but as they very very
much, and have a tendency to change over time really fast, it's best for you to
tune it to the one you have on your field.
Format_1,2,3,4 - different files for 4 different report format.
main.py - interface + the main function
============================================

DrillingReports

Input: A .xlsx file, containing all the drilling data;

Output: A .xlsx file, with format, ready for the analysis (or to be loaded to Petrel as VOL file);

Comments: The reports i saw had a format such as:
Excel worksheets - days of the months, and all the Drilling data for this day on the worksheet.
The format of output document - wellnames as the names of sheets, and all the Data about the well on the Worksheet.

Format_1, Format_2 - two different modules for processing two different formats os files (Format_2 contains merged cells)
============================================

3_MonthlyProductionReport

Input: Well monthly production report file in .xlsx format;

Output: A .xlsx file, with format, ready for the analysis;

Comments: Most reports i saw had a similar format, so i think most of ideas you see in this script will be useful.
