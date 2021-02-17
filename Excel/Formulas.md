This is a repo of various useful Excel formulas. =IFERROR(IF(A1="","",FORMULA),"") is recommended to blank out any cells that don't comply. This can be changed if required.

# Giving the Month in Upper Case Format.
=UPPER(TEXT(DATE(YEAR(TODAY()),MONTH(A1),1),"MMM"))
## Variables:
MONTH(A1) Points to the cell that you are getting the month from. A1 should be changed to the relevant input cell.
## Example (Input = Output):
25/01/1995 (DMY) = JAN

# Taking the Month (in "MMM" format) and giving the next month along.
=UPPER(TEXT(DATE(YEAR(TODAY()),MONTH(DATEVALUE(A1&1))+1,1),"MMM"))
## Variables:
MONTH(DATEVALUE(A1&1))+1,1) Points to the cell that you are getting the month from. A1 should be changed to the relevant input cell.
## Example (Input = Output):
JAN = FEB

# Giving the Date of the Monday of the same Week as Input
=IFERROR(IF(B1="","",A1-WEEKDAY(A1,3)),"")
## Variables:
- A1-WEEKDAY(A1,3)) Points to the cell that you are getting the input from. A1 should be changed to the relevant input cell.
- IF(B1=""... gives a cell to look at to see if the row needs to be omitted. It looks for blank cells, and can be the same cell as the input or can be different.
## Example (Input = Output):
25/01/1995 (DMY) = 23/01/1995
