# Excel tips and tricks

## Cleaning a dataset

First fix the column heads in place
1. Go to `View`.
2. Click on `Freeze Top Row`.

### Step 1 Remove duplicates

1. Go to `Data` > `Sort & Filter` > `Advanced`
2. Then select `filter the list in place`
3. Check `Unique records only`
4. Click `OK`
5. To remove duplicates entirely, go to `Data` > `Remove Duplicates`.
6. To clear the filter
7. Go to `Data` > `Filter` > `Clear`

### Step 2 Replace empty cells

1. Make sure to only select the area containing data.
2. Use the `Find` function (magnify glass, top right). Click on `Replace`.
3. Leave `Find` blank and type _null_ in Replace. 
4. Click on `Replace All`.

### Step 3 Remove corrupted data

1. Go to `Data` > `Auto-Filter` to display the filter arrows in the column heads
2. Click on the filter arrow in the column head.
3. In the menu select the data that is corrupted.
4. Then delete those rows.
5. Remove the filter

### Step 4 Check Spelling

1. Go to `Tools` > `Spelling`
2. Follow the prompts
3. Click `OK` when done

### Step 5 To change the character case

1. Create a new column
2. In the new columnn type:
3. For lower case `=LOWER(cell)` 
4. For UPPER case `=UPPER(cell)`
5. For Proper Case `=PROPER(cell)`

### Step 6 To trim extra white spaces from text

1. Create a new column
2. In the new columnn type:
3. For lower case `=TRIM(cell)`

### Step 7 Split and Join data

1. Create a new column
2. In the new columnn type:
3. If text such as names, use `LEFT`, `MID`, `RIGHT`, `SEARCH`, and `LEN`
4. EG. To split "Joe Bloggs" use `=LEFT(cell,SEARCH(" ",cell))` and then `RIGHT(cell,SEARCH(" ",cell))`
5. If date and time in the same cell, format the column to either Date or Time 
6. Then in the first column type `=INT(A2)` This produces just the date
7. In the next column, subtract the new date from the original cell, eg `=A2-C2`
8. Or use `=TEXT(A2,"mm/dd/yyyy")` and `=TEXT(A2,"hh:mm")` - The `TEXT` function changes the format of the cell to text.
9. Or use `Text to Columns` under Data in the menu, make sure to specify the cell for the first part of the data.
10. To join cells, use `&` EG `=D3&D4`
11. To add a space or comma, do this `=D3&", "&D4` - note the use of & after the first cell and before the second cell.
12. To join date and time, try this `=C28& " " &TEXT(D28, "DD,MM,YY")` 
13. More complex, `=TEXT(C28, "HH/MM AM/PM")& " " &TEXT(D28, "DD,MM,YY")`

### Step 8 Join data into one column

1. Create a new column
2. In the new columnn type: `=CONCAT(A2," ",B2)`

### Step 9 To change from long to wide data

1. In a new cell, type `=TRANSPOSE(A1:B4)`

### Step 10 VLOOKUP

1. Syntax is `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`
2. EG. `=VLOOKUP(A2,'Sheet 2'!A2:B31,2,false)`

## Formulas

### The Basics

To Add, select cell F3, type `=C3+C4`, then press Return. 

To Subtract, select cell F4, type `=C3-C4`, then press Return. 

To Multiply, select cell F5, type `=C3*C4`, then press Return.

To Divide, select cell F6, type `=C3/C4`, then press Return.

To raise a value to a power, select cell F6, type `=C3^F6`, then press Return.

To add a range of cells, type `=SUM(A1:A6)`, then press Return. Additional arguments use a comma, eg `=SUM(A1:A6,B1:B6)`

To count the number of cells, type `=COUNT(A1:A6)`, then press Return.

`=MEDIAN(A1:A6)` gives the middle number in a list.

`=MODE(A1:A6)` gives the most common number in a list.

Use the `MIN` function to get the smallest number in a range of cells. EG `=MIN(A1:A6)`

Use the `MAX` function to get the largest number in a range of cells. EG `=MAX(A1:A6)`

Additional arguments use a comma, eg `=MAX(A1:A6,B1:B6)`

### Time Functions

To get today's date, type `=TODAY()`, then press Return.

To get the time, type `=NOW()`, then press Return.

Add and subtract times, EG `=((D35-D32)-(D34-D33))*24` - the 24 converts into hours. Make sure that this cell is formatted to Number and not Time.

### IF Statements

`IF` statements enable you to make logical comparisons between conditions, eg `=IF(A12="Apple",TRUE,FALSE)` If the condition is met, ie A12 equals Apple then True is returned, if the condition isn't met then False is returned.

`TRUE` and `FASLE` don't need to be in quotes but `"Yes"` and `"no"` would need quotes. 

Numbers also don't need quotes. `=IF(A12>=100,"Yes","No")` 

`IF` can also force another calcuation if a conditin is met, eg `=IF(A12>=100,"Yes","No")` or `=IF(C2>B2,”Over Budget”,”Within Budget”)`

In this formula, `=IF(C2>B2,C2-B2,0)`, the result of a calculation is returned `C2-B2`, else return 0.

`=IF(E7=”Yes”,F5*0.2,0)` This multiplies E7 by 0.2, if the condition is `No` then it returns 0.

Or `=IF(A12="Yes",A11*Shipping,0)` This multiplies A11 by the Shipping costs, if the condition is `No` then it returns 0.

Shipping is a Named Range and this is set using `Formulas` > `Define Name` This means that you only need to change it once for the entire workbook.

<img width="259" alt="image" src="https://user-images.githubusercontent.com/117950069/214031314-f4f8021c-5a28-4876-b748-b1a6337a4188.png">

Use the `IFS` function to check whether one or more conditions are met and returns a value that corresponds to the first TRUE condition.

eg `=IFS(B2>=90,"A",B2>=70,"B",B2>=50,"C",B2<50,"D")

### VLOOKUP Statements

`VLOOKUP` lets you look up a value in a column on the left, then returns information in another column to the right if it finds a match.

eg `=VLOOKUP (A12,B12:C22,2,FALSE)` This looks for the data in `A12` and tries to find a match in the data selected from B12 and C22, if it finds a match it moves to the column next to it (signified by the `2`) and returns the data in that cell. `FALSE` here means an exact match.

If there isn't a match then `N/A` will be returned. To hide this, use an `IF` statement, eg `=IF(A12="","",VLOOKUP (A12,B12:C22,2,FALSE)`

You can also use an `IFERROR` statement, eg `=IFERROR(VLOOKUP(A12,B12:C22,2,FALSE),"")`

In both cases `""` is used to display nothing, but you can use text, eg `"Nothing to show"`.

### XLOOKUP Statements



### Conditional Statements

eg `SUMIF` `SUMIFS` `COUNTIF` `COUNTIFS`

`SUMIF` lets you sum in one range based on a specific criteria you look for in another range.

eg `=SUMIF(C3:C14,C17,D3:D14)` The first part is the range to look at, `C17` is the value or text to find, then the sum that match in `D3:D14`

eg `=SUMIF(C3:C14,"Apples",D3:D14)` The first part is the range to look at, `"Apples"` is the value or text to find, then the sum that match in `D3:D14`

eg `=SUMIF(D118:D122,">=50")` This is is different type of criteria.

`SUMIFS` is the same as SUMIF, but it lets you use multiple criteria.

Format is `SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

eg `=SUMIFS(H3:H14,F3:F14,F17,G3:G14,G17)` `H3:H14` is the range to sum, next `F3:F14` is the range to look for matches, `F17` is the criteria, then `G3:G14` is the second range to look for matches, and `G17` is the criteria for the second match.

`COUNTIF` and `COUNTIFS` let you count values in a range based on a criteria you specify. 

`COUNTIF` counts the number of instances in a column/row. Be sure to enclose the criteria argument in quotes unless its a cell reference. Criteria aren't case sensitive. Wildcard characters —the question mark (?) and asterisk (*)—can be used in criteria. If you want to find an actual question mark or asterisk, type a tilde (~) in front of the character.

eg `COUNTIF(C50:C61,C64)` first is the range to look at, then the value to look for.

eg `COUNTIF(C50:C61,"SpecificWord")` counts the cells that contain SpecificWord.

eg `=COUNTIF(B2:B5,">55")` counts the cells with values greater than 55.

eg `=COUNTIF(A2:A5,"*")` the `*` looks for any cells with text.

eg `=COUNTIF(A2:A5,"?????es")` looks for 7 character words that end in es.

`COUNTIFS` counts the number of instances but with more criteria.

eg `=COUNTIFS(F50:F61,F64,G50:G61,G64)` so `F50:F61` is the range to count, `F64` is the criteria to match, `G50:G61` is the second range and `G64` is the second criteria to match.

Others include `AVERAGEIF` `AVERAGEIFS` `MAXIFS` and `MINIFS`

### Problems with formulas

If there is an error, use `Error-checking` option under `Formulas` to help with finding a solution. 

## Pivot tables
