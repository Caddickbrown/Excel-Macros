# Formulas
This is a repo of various useful Excel formulas. =IFERROR(IF(A1="","",FORMULA),"") is recommended to blank out any cells that don't comply. This can be changed if required.

# Template
Code: 000
## Explanation:
This is a template for how each item should be laid out.
## Example (Input = Output):
Input Value = Output Value
## Formula:
=FORMULA
## Variables/Considerations:
- This contains any points that need explaining for the FORMULA to work.
- Short-hand Dates should be explained in brackets afterwards as to their format. For instance 25th of January, 1995 can be written as 25/01/1995 (DMY), 1995/01/25 (YMD), 01/25/1995 (MDY). As long as explained, the different formats can be used interchangeably.

# Upper Case Month
Code: 001
## Explanation:
This formula will give the Month in Upper Case Format
## Example (Input = Output):
25/01/1995 (DMY) = JAN
## Formula:
=UPPER(TEXT(DATE(YEAR(TODAY()),MONTH(A1),1),"MMM"))
## Variables/Considerations:
- MONTH(A1) Points to the cell that you are getting the month from. A1 should be changed to the relevant input cell.

# Next Month Along ("MMM" Format)
Code: 002
## Explanation:
This formula will take the Month (in "MMM" format) and give the next month along
## Example (Input = Output):
JAN = FEB
## Formula:
=UPPER(TEXT(DATE(YEAR(TODAY()),MONTH(DATEVALUE(A1&1))+1,1),"MMM"))
## Variables/Considerations:
MONTH(DATEVALUE(A1&1))+1,1) Points to the cell that you are getting the month from. A1 should be changed to the relevant input cell.

# Monday Date
Code: 003
## Explanation:
This formula will give the date of the Monday in the same Week as input value.
## Example (Input = Output):
25/01/1995 (DMY) = 23/01/1995
## Formula:
=IFERROR(IF(B1="","",A1-WEEKDAY(A1,3)),"")
## Variables/Considerations:
- A1-WEEKDAY(A1,3)) Points to the cell that you are getting the input from. A1 should be changed to the relevant input cell.
- IF(B1=""... gives a cell to look at to see if the row needs to be omitted. It looks for blank cells, and can be the same cell as the input or can be different.

# Stock Suggested Adjustment
Code: 004
## Explanation:
This formula will give a suggested Adjustment for your stocks rounded up to the nearest 10. Difficult to explain - but combination of IFERROR, FLOOR/MIN, and CEILING can give required result.
## Formula:
=IFERROR(FLOOR((MIN((BM3-$K3),BA3+BS3)*-1),-10),CEILING($K3-BM3,10))
## Variables/Considerations:
- A lot.

# Suggested Production Plan (Alt.)
Code: 005
## Explanation:
This is a simple Formula that has promise and could supplant 004. Needs work and may be too simple. Potential for deletion.
## Formula:
=IF(R4>$G4,0,$G4-R4)
## Variables/Considerations:
- This formula does not account for cancelling orders that have already been raised.

# Consistency Check
Code: 006
## Explanation:
This formula checks for consistency in a value (for instance pricing) for data that contains multiple versions of the same product. If there is a difference, it will flag up as "DIFFERENCE"
## Formula:
=IF(B2=B3,IF(E2=E3,"","DIFFERENCE"),"")
## Variables/Considerations:
- B2/B3 refers to the Product Code and E2/E3 refers to the "Pricing" Value.

# ISDATE
Code: 007
## Explanation:
This formula checks if a cell is of a date format and if so, uses that cell - otherwise it pulls the date from the previous cell. It's good for translating data that has been formatted badly into a better format, and is of a more transitional use to get to where you want. It will likely be deleted soon after.
## Example (Input = Output):
Input Value = Output Value
## Formula:
=IF(LEFT(CELL("format",B4))="D",B4,A4)
## Variables/Considerations:
- This Formula would be put into A5
- B4 would contain data and dates above them. formatted into sections
- A4 is the previous date

# Pack Size Conversion
Code: 008
## Explanation:
Useful for conversion based on pack sizes from a built up spreadsheet.
## Example (Input = Output):
144 = 96
## Formula:
=ROUNDUP((F28*VLOOKUP(C28,$C:$M,4,FALSE))/VLOOKUP(C28,$C:$M,8,FALSE),0)
## Variables/Considerations:
- F28 is the number of packs that requires conversion
- VLOOKUP looks for the initial pack size and divides it against the nw pack size.
- C28 is the value within the data table

# Upper Case Month & Year
Code: 009
## Explanation:
This formula will give the Month in Upper Case Format
## Example (Input = Output):
25/01/1995 (DMY) = JAN 1995
## Formula:
=CONCATENATE(UPPER(TEXT(DATE(YEAR(TODAY()),MONTH(G2),1),"MMM"))," ",YEAR(G2))
## Variables/Considerations:
- G2 Points to the cell that you are getting the month from. G2 should be changed to the relevant input cell.

# Year-Month Numeric
Code: 010
## Explanation:
This formula will give the Month in Upper Case Format
## Example (Input = Output):
25/01/1995 (DMY) = 1995-01
## Formula:
=CONCATENATE(YEAR(G2),"-",TEXT(MONTH(G2),"00"))
## Variables/Considerations:
- G2 Points to the cell that you are getting the month from. G2 should be changed to the relevant input cell.
- This is useful for getting things into date order

# If Character In String
Code: 011
## Explanation:
This formula will give a value based on if a specific singular/set of characters is found in a string.
## Example (Input = Output):
/=2, R=1
## Formula:
=IF(ISNUMBER(SEARCH("/",AF2)),2,1)
## Variables/Considerations:
- AF2 is the cell to look at.
