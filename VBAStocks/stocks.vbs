Sub stocks()

'A. Create a script that will loop through all the stocks for one year for each run and take the following information.

'    1. The ticker symbol.
'    2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'    3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'    4. The total stock volume of the stock.

' B. You should also have conditional formatting that will highlight positive change in green and negative change in red.


' --------------------------------------------
' SET THE VARIABLES
' --------------------------------------------

'Set a variable for the ticker
Dim ticker As String
        
' Set a variable for the row of the summary table
Dim summary_table_row As Integer
summary_table_row = 2
    
' Set a variable for the yearly change
Dim YoY_abs As Double
    
'Set variables to help calculate the yearly change
Dim year_begin As Double
Dim count_begin As Double
    
' Set a variable for the percent change
Dim YoY_pct As Double
    
' Set a variable for the total stock volume
Dim stock_volume As Double
    
       
' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
    
For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Columns("J:L").AutoFit

' Determine the Last Row
LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row

'MsgBox (LastRow)

' Filter zeros out from Column C <open>

Range("A1", "G" & LastRow).AutoFilter Field:=3, Criteria1:="<>0"

' Determine the Last Row after filtering

LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row

'MsgBox (LastRow)

' Set variables to initial value

stock_volume = 0
year_begin = ws.Cells(2, 3).Value

' Loop through rows

For i = 2 To LastRow
         
  If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
      
    stock_volume = stock_volume + ws.Cells(i, 7).Value
    
  ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
     ticker = ws.Cells(i, 1)
     YoY_abs = ws.Cells(i, 6).Value - year_begin
     YoY_pct = YoY_abs / year_begin
     ws.Cells(summary_table_row, 9).Value = ticker
     ws.Cells(summary_table_row, 10).NumberFormat = "0.000000"
     ws.Cells(summary_table_row, 10).Value = YoY_abs
     ws.Cells(summary_table_row, 11).Value = YoY_pct
     ws.Cells(summary_table_row, 12).Value = stock_volume
          
     summary_table_row = summary_table_row + 1
    
    stock_volume = 0
    count_begin = i + 1
    year_begin = ws.Cells(count_begin, 3).Value
      
   End If

Next i


' ------------------------------------------------
' APPLY THE RIGHT FORMAT TO PERCENTAGE CHANGE
' ------------------------------------------------

' Determine the Last Row in the summary table
       
LastRow = ws.Cells(Rows.count, 11).End(xlUp).Row

' Define variables for conditional formatting
       
Dim rng As Range
Dim condition1 As FormatCondition
Dim condition2 As FormatCondition
        
'Determine the range to apply the formatting on
        
Set rng = ws.Range("K2", "K" & LastRow)
Set rng = Range("K2", "K" & LastRow)

'Delete any preexisting formatting
        
rng.FormatConditions.Delete
        
'Apply the percentage formatting
        
For i = 2 To LastRow

    ws.Cells(i, 11).NumberFormat = "0.00%"

Next i
        
'Define conditional formatting criteria
        
Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
  
'Define and set the format to be applied for each condition
        
With condition1
 .Interior.Color = vbGreen
End With

With condition2
 .Interior.Color = vbRed
End With


Next ws

' --------------------------------------------
' HOMEWORK COMPLETE
' --------------------------------------------

MsgBox ("Homework Complete!")


End Sub