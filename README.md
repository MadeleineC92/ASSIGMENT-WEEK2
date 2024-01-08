# ASSIGMENT-WEEK2

**Recivied help from TA**
[NOTE: use the provided codes just as an example and guidance and may not suit your need precisely!]
Variable Naming Consistency:
The variable names in your code are not consistent. For example, you used both percentage_change and percentage for different variables. To avoid confusion, it's better to stick with one naming convention.
Dim percentage_change As Double 
' ... 
percentage_change = (closing - opening) / opening * 100 
2. Undefined Variables:
you used some variables without declaring them. For instance, yearly_change is used, but it's not explicitly declared. It's good practice to declare all variables before using them.
Dim yearly_change As Double 
' ... 
yearly_change = closing - opening 
3. Incorrect Assignment of Ticker Code:
The assignment of the ticker code seems to be incorrect. The ticker code should be associated with each row, and it should reset when a new ticker is encountered. The code for updating the ticker code and writing the ticker in column "I" needs to be adjusted.
'ticker code 
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then 
    tickercode = tickercode + 1 
    ws.Cells(i, 9).Value = ticker 
End If 
4. Volume Calculation:
The variable Volume is declared but not used. Instead, Stock_Volume is used for volume calculations. Make sure to use consistent variable names.
Dim Stock_Volume As Double 
' ... 
Stock_Volume = Stock_Volume + Cells(i, 7).Value 
5. Incorrect Column Reference for Ticker Output:
When writing the ticker in column "I," it should be in the same row as the current data row, not in the row corresponding to the ticker code.
ws.Cells(i, 9).Value = ticker 
After making these adjustments, your code should work more effectively for calculating the yearly change, percentage change, and total stock volume for each stock ticker.

recivied help from askbcs learning assistant for interior colour and selescting worksheets
recivied help from Tutour simion (Sunday 8.01.2024)

