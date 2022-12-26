Sub MultipleYearStockData()

' -------------------------------------------------
' begin Loop through each worksheet
' -------------------------------------------------
For Each ws In Worksheets

        ' --------------------------------------------
        ' create variables
        ' --------------------------------------------
        
        Dim WorksheetName As String
        
        ' i represents the current row
        Dim i As Long
        ' j represents the start row the the Ticker, will need this to calculate changes
        Dim j As Long
        
        ' create variable for the last row in column A
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' create variable for tickercount, this will keep track of what row to input our ticker in
        Dim TickerCount As Long
        
        ' create variables for greatest % increase, decrease, and total volume for Summary Table
        Dim GrePerIncrease As Double
        Dim GrePerDecrease As Double
        Dim GreTotalVol As Double
        
        ' grab the worksheetname
        WorksheetName = ws.Name
        
        ' msgbox to test worksheetname
        ' msgbox(WorksheetName) works, removed.
        
        ' ---------------------------------------
        ' Create Headers on each sheet
        ' ---------------------------------------
        
        ' create column headers for each worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' create headers for the summary table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' set TickerCount to the second row
        TickerCount = 2
        
        ' set the start row to 2 for the ticker
        j = 2
        
        ' ----------------------------------------------------------------
        ' Begin For Loop
        ' ----------------------------------------------------------------
        
        For i = 2 To LastRow
        
            ' search column A for the first ticker that doesnt match the next (name change)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' write the name of the ticker before name change in column I. using the TickerCount to sequencially move down column I
            ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
            
            ' calc and write the Yearly Change in column J
            ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
            ' calc and write the Percent change in Column K
            ws.Cells(TickerCount, 11).Value = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
            
            ' calc and write the Total Stock Volume in column L
            ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
           ' increase TickerCount by 1 to move to the next row for data input
           TickerCount = TickerCount + 1
           
           ' set new start row for the Ticker. Once a change in the ticker is found, i will represent the last constant Ticker, need to add 1.
           j = i + 1
           
           
            End If
        
        Next i
        
        ' ------------------------------------------------
        ' calculating summary table
        ' ------------------------------------------------
        
        ' create variable for the last rows in column K and L. Will be used in the Summary Table
        Dim LastRowK As Long
        LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        ' assign values to each of the variables in the summary table. Using the max function to find the max value in each range
        GrePerIncrease = WorksheetFunction.Max(Range(ws.Cells(2, 11), ws.Cells(LastRowK, 11)))
        GrePerDecrease = WorksheetFunction.Min(Range(ws.Cells(2, 11), ws.Cells(LastRowK, 11)))
        GreTotalVol = WorksheetFunction.Max(Range(ws.Cells(2, 12), ws.Cells(LastRowK, 12)))
        
        ' write assigned variables to the summary table
        ws.Cells(2, 17).Value = GrePerIncrease
        ws.Cells(3, 17).Value = GrePerDecrease
        ws.Cells(4, 17).Value = GreTotalVol
        
        ' find the ticker that is associated with each value in the summary table
        For i = 2 To LastRowK
        
            ' search for the value in summary table that matches the total data
            If ws.Cells(i, 11).Value = GrePerIncrease Then
            
            ' find the ticker associated with that value and in put the ticker in the summary table
            ws.Cells(2, 16).Value = ws.Cells(i, 9)
            
            ' find the ticker associated with that value and in put the ticker in the summary table
            ElseIf ws.Cells(i, 11).Value = GrePerDecrease Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9)
            
            ' find the ticker associated with that value and in put the ticker in the summary table
            ElseIf ws.Cells(i, 12).Value = GreTotalVol Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9)
            
            End If
            
        Next i
        
        ' -------------------------------
        ' Formatting Cells
        ' -------------------------------
        
        ' change the Percent Change values to a %
        Range(ws.Cells(2, 11), ws.Cells(LastRowK, 11)).NumberFormat = "0.00%"
        Range(ws.Cells(2, 17), ws.Cells(3, 17)).NumberFormat = "0.00%"
        
        ' apply color formatting to the Yearly Change column
        For i = 2 To LastRowK
            
            'search for values less than 0
            If ws.Cells(i, 10).Value < 0 Then
            
            'change the interior of the cell to Red if the value is less than 0
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
            ' if the value is greater than 0 then change the interior color of the cell to Green
            Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
            End If
        
        Next i

Next ws

End Sub
