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

### Step 7 Split data into separate columns

1. Create a new column
2. In the new columnn type:
3. If text such as names, use `LEFT`, `MID`, `RIGHT`, `SEARCH`, and `LEN`
4. EG. To split "Joe Bloggs" use `=LEFT(cell,SEARCH(" ",cell))` and then `RIGHT(cell,SEARCH(" ",cell))`
5. If date and time in the same cell, format the column to either Date or Time 
6. Then in the first column type `=INT(A2)` This produces just the date
7. In the next column, subtract the new date from the original cell, eg `=A2-C2`
8. Or use `=TEXT(A2,"mm/dd/yyyy")` and `=TEXT(A2,"hh:mm")`
9. Or use `Text to Columns` under Data in the menu, make sure to specify the cell for the first part of the data.

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

To Add, select cell F3, type =C3+C4, then press Return. 

To Subtract, select cell F4, type =C3-C4, then press Return. 

To Multiply, select cell F5, type =C3*C4, then press Return.

To Divide, select cell F6, type =C3/C4, then press Return.

To raise a value to a power, select cell F6, type =C3^F6, then press Return.

To add a range of cells, type =SUM(A1:A6), then press Return. Additional arguments use a comma, eg =SUM(A1:A6,B1:B6)

To count the number of cells, type =COUNT(A1:A6), then press Return.

=MEDIAN(A1:A6) gives the middle number in a list.

=MODE(A1:A6) gives the most common number in a list.

Use the MIN function to get the smallest number in a range of cells. EG =MIN(A1:A6)

Use the MAX function to get the largest number in a range of cells. EG =MAX(A1:A6)

Additional arguments use a comma, eg =MAX(A1:A6,B1:B6)

### Time Functions

To get today's date, type =TODAY(), then press Return.

To get the time, type =NOW(), then press Return.

Add and subtract times, EG =((D35-D32)-(D34-D33))*24 ----- the 24 converts into hours. Make sure that this cell is formatted to Number and not Time.

### Joining and splitting cells



## Pivot tables
