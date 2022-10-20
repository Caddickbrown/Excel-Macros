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
- Short-hand Dates should be explained in brackets afterwards as to their format. For instance 25th of January, 1995 can be written as 25/01/1995 (DD/MM/YYYY), 1995/01/25 (YYYY/MM/DD), 01/25/1995 (MM/DD/YYYY). As long as explained, the different formats can be used interchangeably.
- If a new line is needed to be shown ";" will be used to show it.

# Upper Case Month
Code: 001
## Explanation:
This formula will give the Month in Upper Case Format
## Example (Input = Output):
25/01/1995 (DD/MM/YYYY) = JAN
## Formula:
=UPPER(TEXT(A1,"MMM"))
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
- MONTH(DATEVALUE(A1&1))+1,1) Points to the cell that you are getting the month from. A1 should be changed to the relevant input cell.

# Monday Date
Code: 003
## Explanation:
This formula will give the date of the Monday in the same Week as input value.
## Example (Input = Output):
25/01/1995 (DD/MM/YYYY) = 23/01/1995
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
This formula checks if a cell is of a date format and if so, uses that cell - otherwise it pulls the date from the previous cell. It's good for translating data that has been formatted badly into a better format, and is of a more transitional use to get to where you want. It will likely be deleted/pasted as values soon after.
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
25/01/1995 (DD/MM/YYYY) = JAN 1995
## Formula:
=CONCATENATE(UPPER(TEXT(DATE(YEAR(TODAY()),MONTH(G2),1),"MMM"))," ",YEAR(G2))
## Variables/Considerations:
- G2 Points to the cell that you are getting the month from. G2 should be changed to the relevant input cell.

# Year-Month Numeric
Code: 010
## Explanation:
This formula will give the Month in Upper Case Format
## Example (Input = Output):
25/01/1995 (DD/MM/YYYY) = 1995-01
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

# Conformance to Plan (12 Week Average)
Code: 012
## Explanation:
Looks at a table of data
## Example (Input = Output):
(10+15+12)/(18+20+15) = 69.8%
## Formula:
=SUM(SUMIFS('CTP DATA'!$G:$G,'CTP DATA'!$B:$B,"<="&TODAY(),'CTP DATA'!$B:$B,">"&TODAY()-83)/SUMIFS('CTP DATA'!$D:$D,'CTP DATA'!$B:$B,"<="&TODAY(),'CTP DATA'!$B:$B,">"&TODAY()-83))
## Variables/Considerations:
- 'CTP DATA'! - needs to point to the relevant sheet for the data.
- B = Date
- D = Total Plan
- G = Total Achieved
- This can be used to look at last week by replacing TODAY() and TODAY()-83 with TODAY()-7 and TODAY()-90

# Schedule Adherence (12 Week Average)
Code: 013
## Explanation:
Looks at a table of data
## Example (Input = Output):
(10+15+12)/(18+20+15) = 69.8%
## Formula:
=SUM(SUMIFS('CTP DATA'!$E:$E,'CTP DATA'!$B:$B,"<="&TODAY(),'CTP DATA'!$B:$B,">"&TODAY()-83)/SUMIFS('CTP DATA'!$D:$D,'CTP DATA'!$B:$B,"<="&TODAY(),'CTP DATA'!$B:$B,">"&TODAY()-83))
## Variables/Considerations:
- 'CTP DATA'! - needs to point to the relevant sheet for the data.
- B = Date
- D = Total Plan
- E = Completed Planned
- This can be used to look at last week by replacing TODAY() and TODAY()-83 with TODAY()-7 and TODAY()-90

# Text Join Across Multiple Lines (Multiple Line Concatenate)
Code: 014
## Explanation:
This formula concatenates all values that match the input value from range of data.
## Example (Input = Output):
John's Sales = "£455, £245, £120, £150, £310"
## Formula:
=TEXTJOIN(", ",TRUE,IF(Sheet2!A:A=G2,IF(MATCH(Sheet2!H:H,Sheet2!H:H, 0)=MATCH(ROW(Sheet2!H:H),ROW(Sheet2!H:H)),Sheet2!H:H,""),""))
## Variables/Considerations:
- This is an INCREDIBLY resource intensive formula as it checks every line multiple times.

# Variable Column SUMIFS
Code: 015
## Explanation:
Used to look through a data dump file to find the right column and add up values.
## Example (Input = Output):
Data Dump into Released Orders
## Formula:
=SUMIFS(INDEX('Data Dump'!$A:$CA,0,MATCH("Lot Size",'Data Dump'!$A$1:$CA$1,0)),INDEX('Data Dump'!$A:$CA,0,MATCH("Part No",'Data Dump'!$A$1:$CA$1,0)),$A2)
## Variables/Considerations:
- 'Data Dump' - the "Data Dump" tab within Excel. The Column range may need extending depending on how large the data dump file is
- $A2 - Refers to the "Lookup Value"
- This Needs to MATCH on the headers of the Data Dump File - Only Row 1
- "Lot Size" and "Part Number" are the columns that are looked up and referred to.

# Variable Column Layout Lookup
Code: 016
## Explanation:
Pulls through the required column based on searching for the Columns name. 
## Example (Input = Output):
|Something|DesCol|Else |
|:-------:|:----:|:---:|
|    1    |  2   |  3  |
|    1    |  2   |  3  |
|    1    |  2   |  3  |
= "2";"2";"2"
## Formula:
=INDEX('Source Data'!A:DZ,0,MATCH("DesCol",'Source Data'!A$1:DZ$1,0))
## Variables/Considerations:
- A:DZ are the Columns your data looks upto.
- A$:DZ$1 is the row of headers to look through.
- 'Source Data'! refers to the data to search through to find what you want.
- This tends to be used in conjunction with a Sumif, lookup, or something else.


# Percentage Through Day
Code: 017
## Explanation:
This will give you a percentage of how far through the day you are. This can be used to calculate rough ideas of where you should be within a daily plan etc.
## Example (Input = Output):
07/06/2022 13:27 = 0.55 (55%)
## Formula:
=NOW()-TODAY()
## Variables/Considerations:
- This formula is volatile and will update at various intervals. It should keep relatively upto date though and give an idea of how far through the day you are.

# Quarter and Year
Code: 018
## Explanation:
This gives you the relevant Quarter and Year for a date.
## Example (Input = Output):
"01/01/2022" = "Q1 - 2022"
## Formula:
=CONCATENATE("Q",ROUNDUP(MONTH(A2)/3,0)," - ",YEAR(A2))
## Variables/Considerations:
- A2 is the cell with the date. The formula will need changing to account for whatever cell you're looking up.

# Component List
Code: 019
## Explanation:
This formula will give you a list of component parts for a specific parent part from a list of all Manufacturing Structures. It will "Spill" down into cells below.
## Example (Input = Output):
590389 = [; refers to a new line] 60300328;60300483;60300396;28151BID;28033BID;585775;585752;585736
## Formula:
=IFERROR(FILTER(Structure!N:N,B1=Structure!C:C),"-")
## Variables/Considerations:
- Structure!N:N Refers to Component Part List
- B1 Refers to the Parent Part you need to Lookup
- Structure!C:C Refers to Parent Part List

# Closest Value
Code: 020
## Explanation:
This gives you the closest value to an inputted value in a cell.
## Example (Input = Output):
14 & 20,20,16,15,18,19,17,54,12 = 15
## Formula:
=INDEX(B3:B22,MATCH(MIN(ABS(B3:B22-E2)),ABS(B3:B22-E2),0))
## Variables/Considerations:
- B3:B22 is your range of values to search within
- E2 is your search value

# Highest Value From Another Column
Code: 021
## Explanation:
This gives you the highest value of an array of "Lookup"s.
## Example (Input = Output):
J=12
J=15
MAX J = 15
## Formula:
=MAX(FILTER(L:L,A:A=A2,0))
## Variables/Considerations:
- L:L is the "Data" column that will pull info from.
- A:A is the "Lookup" Column.
- A2 is what you're looking up.

# Retrieve the Last entered value from a column
Code: 022
## Explanation:

## Example (Input = Output):

## Formula:
=LOOKUP(2,1/('This Week Tracker'!$F:$F<>""),'This Week Tracker'!$F:$F)
## Variables/Considerations:

