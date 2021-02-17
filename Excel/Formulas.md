This is a repo of various useful Excel formulas. =IFERROR(IF(A1="","",FORMULA),"") is recommended to blank out any cells that don't comply. This can be changed if required.

# Template
## Explanation
This is a template for how each Item should be laid out.
## Example (Input = Output):
Input Value = Output Value
## Formula
=FORMULA
## Variables/Considerations:
- This contains any points that need explaining for the FORMULA to work.

# Upper Case Month
## Explanation
This formula will give the Month in Upper Case Format
## Example (Input = Output):
25/01/1995 (DMY) = JAN
## Formula
=UPPER(TEXT(DATE(YEAR(TODAY()),MONTH(A1),1),"MMM"))
## Variables/Considerations:
- MONTH(A1) Points to the cell that you are getting the month from. A1 should be changed to the relevant input cell.


# Next Month Along ("MMM" Format)
## Explanation
This formula will take the Month (in "MMM" format) and give the next month along
## Example (Input = Output):
JAN = FEB
## Formula
=UPPER(TEXT(DATE(YEAR(TODAY()),MONTH(DATEVALUE(A1&1))+1,1),"MMM"))
## Variables/Considerations:
MONTH(DATEVALUE(A1&1))+1,1) Points to the cell that you are getting the month from. A1 should be changed to the relevant input cell.

# Monday Date
## Explanation
This formula will give the date of the Monday in the same Week as input value.
## Example (Input = Output):
25/01/1995 (DMY) = 23/01/1995
## Formula
=IFERROR(IF(B1="","",A1-WEEKDAY(A1,3)),"")
## Variables/Considerations:
- A1-WEEKDAY(A1,3)) Points to the cell that you are getting the input from. A1 should be changed to the relevant input cell.
- IF(B1=""... gives a cell to look at to see if the row needs to be omitted. It looks for blank cells, and can be the same cell as the input or can be different.

# Stock Suggested Adjustment
## Explanation
This formula will give a suggested Adjustment for your stocks rounded up to the nearest 10. Difficult to explain - but combination of IFERROR, FLOOR/MIN, and CEILING can give required result.
## Example (Input = Output):

## Formula
=IFERROR(FLOOR((MIN((BM3-$K3),BA3+BS3)*-1),-10),CEILING($K3-BM3,10))
## Variables/Considerations:
- A lot.
