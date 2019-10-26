Usage:

Call the program from the CMD.

Syntax:
"File path > Desired Sheet Name > Date Format" output file name

=====================================================================================================================================================================

1 file

CsvToXlsx.exe "C:\temp-copies\test.csv>sheet 1>dd-mm-yyyy" "C:\temp-copies\combined.xlsx"

=====================================================================================================================================================================

Multiple files - Each on its own sheet

CsvToXlsx.exe "C:\temp-copies\test.csv:sheet 1>dd-mm-yyyy"  "C:\temp-copies\Bulk Note.csv:sheet 2>dd-mm-yyyy" "C:\temp-copies\combined.xlsx"

=====================================================================================================================================================================

Auto Formating Of Cells

0.00        => Formated to currency number
(0.00)      => Formated to currency number negative

0.00000     => Formated to decimal
(0.00000)     => Formated to decimal negative

100         => Formated to number
(100)       => Formated to -100

01/12/2019  => Formated to Date (enter desired format in CMD)

*numbers starting with 0 or greater that 15 characters are processed as text

=====================================================================================================================================================================

Styling Cells

Pass values with column data prepend with #$#property=value

e.g Customer Details#$#font-weight=bold#$#text-color=#456879

Available Styles:

column-width=120  (if a width is specified for 1 column it should be specified for all. IF no columns specify a width, width will be set to auto.)

font-bold=true
font-italic=true
font-underline=true
font-color=#456FFF
