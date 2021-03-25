Attribute VB_Name = "Module1"
Sub stock()

' Loop through all sheets
For Each ws In Worksheets

' Set an initial variable for holding the Tickers Symbols
Dim Ticker_Symbols As String
  
' Set an initial variable for holding the Total Volume per ticker symbol
Dim Total_Volume As Double
Total_Volume = 0

' Set an initial variable for holding the Yearly Change per ticker symbol
Dim Yearly_Change As Double
Yearly_Change = 0

' Set variable for the percent change
Dim Percent_Change As Double
Percent_Change = 0
ws.Columns("K").NumberFormat = "0.00%"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

' Change Column Sizes
ws.Cells(2, 15).ColumnWidth = 25
ws.Cells(2, 12).ColumnWidth = 20

' Define column names
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

' Define column names for bonus questions
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

' Keep track of the location for each Tickers Symbols in the summary table
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

' Determine the Last Row variable
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Define Start Row
Dim Start_Row As Long
Start_Row = 2

' Loop through all the Tickers Symbols
For i = 2 To LastRow

' Check for a Ticker Symbol change
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
' Set the Tickers names
      Ticker_Symbols = ws.Cells(i, 1).Value
    
' Add to the Volume Total
Total_Volume = Total_Volume + ws.Cells(i, 7).Value

' Define End Row
Dim End_Row As Long
End_Row = i
       
' Set Yearly Change
Yearly_Change = ws.Cells(End_Row, 6).Value - ws.Cells(Start_Row, 3).Value

' Set Percent Change with a loop to avoid #DIV/0!
If ws.Cells(Start_Row, 3).Value = 0 Then

Percent_Change = 0

        Else
        
Percent_Change = (Yearly_Change / ws.Cells(Start_Row, 3).Value)

End If

' Print Ticker Symbols in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbols
      
' Print Volume Total in the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_Volume
      
' Print Yearly Change in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
' Print Percent Change in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      
' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
' Reset the Volume Total
Total_Volume = 0

' Reset Start_Row and End_Row
Start_Row = i + 1
End_Row = 0


' If the cell following a row is the same ticker
Else

' Add to the Volume Total
Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    
    End If
    
Next i


'Colour

'Define Last Row of Summary Chart
Dim Last_Summary_Row As Long
Last_Summary_Row = Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To Last_Summary_Row
    If ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
        
    Else
    ws.Cells(i, 10).Interior.ColorIndex = 4
        
    End If
        
Next i



' Bonus



' Find the greatest percent increase

' Set an initial variable for holding the maximum percent change
Dim max_change As Double
max_change = 0

' Loop through Last summary row data
For i = 2 To Last_Summary_Row
        
' Comparing each value in the column to the maximum value
    If ws.Cells(i, 11).Value > max_change Then
        
' Then max_change becomes the new highest value
    max_change = ws.Cells(i, 11).Value
    
' Print greatest percent increase
    ws.Cells(2, 17).Value = max_change
        
' Print ticker value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
    Else
        
        
    End If
    
        
Next i



' Find the greatest percent decrease

' Set an initial variable for holding the minimum percent change
Dim min_change As Double
min_change = 0

' Loop through Last Summary Row
For i = 2 To Last_Summary_Row

' Comparing each value in the column to the minimum value
    If ws.Cells(i, 11).Value < min_change Then
    
' Then min_change becomes the new lowest value
    min_change = ws.Cells(i, 11).Value
    
' Print greatest percent decrease
    ws.Cells(3, 17).Value = min_change
          
' Print ticker value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    
    Else
        

        
    End If
        
Next i


' Set maximum total volume

' Set an initial variable for holding the maximum total volume
Dim max_volume As Double
max_volume = 0

' Loop through Last Summary Row
For i = 2 To Last_Summary_Row
        
' Comparing each value in the column to the maximum volume
    If ws.Cells(i, 12).Value > max_volume Then
        
' Then max_volume becomes the new highest value
    max_volume = ws.Cells(i, 12).Value
        
' Print greatest total volume
    ws.Cells(4, 17).Value = max_volume
        
' Print the ticker value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        
    Else
        
        
    End If
        
Next i




Next

End Sub









