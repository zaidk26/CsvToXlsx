﻿Usage:

Call the program from the CMD.

Syntax:
"File path > Desired Sheet Name > Date Format > column size > open file > max columns" output file name

column size options : auto , 15 (any valid interger) , custom

=====================================================================================================================================================================

1 file

CsvToXlsx.exe "C:\temp-copies\test.csv>sheet 1>dd-mm-yyyy>auto>Y>10" "C:\temp-copies\combined.xlsx"

=====================================================================================================================================================================

Multiple files - Each on its own sheet

CsvToXlsx.exe "C:\temp-copies\test.csv>sheet 1>dd-mm-yyyy>auto>N>20"  "C:\temp-copies\Bulk Note.csv>sheet 2>dd-mm-yyyy>12>N>100" "C:\temp-copies\combined.xlsx"

=====================================================================================================================================================================
NOTE:
The Parser does not recognize blank lines as a row.

To create a blank row at the word in column !!!BLANK_ROW 
=====================================================================================================================================================================

Auto Formating Of Cells

0.00        => Formated to currency number
(0.00)      => Formated to currency number negative

0.00000     => Formated to decimal
(0.00000)     => Formated to decimal negative

100         => Formated to number
(100)       => Formated to -100

01/12/2019  => Formated to Date (enter desired format in CMD)


* To force a field which has digits to format as text, append to field: !#TEXT#!
*numbers starting with 0 or greater that 15 characters are processed as text

=====================================================================================================================================================================

Styling Cells

Pass values with column data prepend with #$#property=value

e.g Customer Details#$#font-weight=bold#$#text-color=#456879

Available Styles:
-----------------

column-width=120  (if a width is specified for 1 column it should be specified for all. IF no columns specify a width, width will be set to auto.)

column-freeze=true

column-merge=2 (merge the next 2 cells)

column-border=right
column-border=left
column-border=top
column-border=bottom
column-border=all

column-background=#456FFF

font-bold=true
font-italic=true
font-underline=true
font-color=#456FFF   
font-size=26  

column-horizontal-align=Center   (Center,CenterContinuous,Left,Right,Distributed,Fill,Justify)
column-vertical-align=Center     (Center,Top,Bottom,Distributed,Justify)

