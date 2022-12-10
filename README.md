# Excel tips and tricks

## Cleaning a dataset

The basic steps for cleaning data are as follows:
	1.	Import the data from an external data source. 
	2.	Create a backup copy of the original data in a separate workbook. 
	3.	Ensure that the data is in a tabular format of rows and columns with: similar data in each column, all columns and rows visible, and no blank rows within the range. For best results, use an Excel table. 
	4.	Do tasks that don't require column manipulation first, such as spell-checking or using the Find and Replace dialog box. 
	5.	Next, do tasks that do require column manipulation. The general steps for manipulating a column are:
	a.	Insert a new column (B) next to the original column (A) that needs cleaning. 
	b.	Add a formula that will transform the data at the top of the new column (B). 
	c.	Fill down the formula in the new column (B). In an Excel table, a calculated column is automatically created with values filled down. 
	d.	Select the new column (B), copy it, and then paste as values into the new column (B). 
	e.	Remove the original column (A), which converts the new column from B to A. 
Efficiency tools

Conditional formatting is a spreadsheet tool that changes how cells appear when values meet specific conditions.

Remove duplicates" is a tool that automatically searches for and eliminates duplicate entries from a spreadsheet. 

A text string is a group of characters within a cell, most often composed of letters.

An important characteristic of a text string is its length, which is the number of characters in it. You'll learn more about that soon. For now, it's also useful to know that a substring is a smaller subset of a text string.

Split is a tool that divides a text string around the specified character and puts each fragment into a new and separate cell. Split is helpful when you have more than one piece of data in a cell and you want to separate them out.

CONCATENATE is a function that joins multiple text strings into a single string. 

A function is a set of instructions that performs a specific calculation using the data in a spreadsheet.

COUNTIF is a function that returns the number of cells that match a specified value. Basically, it counts the number of times a value appears in a range of cells. 

=COUNTIF(range:range,”condition”)

Eg COUNTIF(A2:A7,”>100”)

Syntax is a predetermined structure that includes all required information and its proper placement.

LEN is a function that tells you the length of the text string by counting the number of characters it contains.

=LEN(range)
=LEN(A2)


LEFT is a function that gives you a set number of characters from the left side of a text string. RIGHT is a function that gives you a set number of characters from the right side of a text string.

=LEFT(range, number of characters)

=RIGHT(range, number of characters)

Eg    =LEFT(A2,4)


MID is a function that gives you a segment from the middle of a text string. 

=MID(range, reference starting point,number of middle characters)

Eg    =MID(D2,4,2)

CONCATENATE, which is a function that joins together two or more text strings. 

= CONCATENATE(item 1,item 2)

EG =CONCATENATE(H3,I3)

TRIM is a function that removes leading, trailing, and repeated spaces in data. Sometimes when you import data, your cells have extra spaces, which can get in the way of your analysis.

=TRIM(range)

EG =TRIM(A2)
