# VBA-Challenge
## Description
These are the files that compose my Module 2 Challenge for the WGU Data Abalytics Boot Camp.

This is definitely not the most sophisticated script poassible, but I was able to piece it with what I learned from in class 
lectures and examples, along with online sources and study groups

I had to use multiple resources to write the VBA script, I will credit/source them below. Since I formatted it as 3 separate
sub-routines I will go through them each in order.

### Sub Mainloop

This sub-routine is how I was able to excecute my script to all the sheets in the work book. 
I found a portion of the code at: https://excelchamps.com/vba/loop-sheets/ . I was still struggling with the loop,
I joined a study group where my classmate Bryan Jones suggested I try the Worksheet.Activate method and this seemed
to address the issue. https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.activate(method)

### Sub TickerSummary

The [] method that was used to set the cell/column names was shared by another classmate during a study session. It is easier than setting a cell or range equal to the name so I went with it.  

I used the range.autofit method to format the columns in the sheet. https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

To set the last row in my for loop I used 'lastrow = Cells(Rows.Count, 1).End(xlUp).Row'
This is a piece of script we learned in class doing the "Star Counter Example"

My main 'if' statement I  used a portion of the script we learned during the "Credit Charges Example" in class.

I used 'FormatPercent' to format my "Percent Change" results. https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/formatpercent-function

### Sub Greatest_Increase_Decrease

Since I was struggling to come up with the correct syntax to include the "Greatest % Increase" "Greatest % Decrease" and "Greatest Total Volume" script into the existing
previous subroutine, I decided that the best way to proceed would be to create a third sub-routine that would work with the results from the "MainLoop" sub-routine to calculate
my results.



