Attribute VB_Name = "Module1"
Sub MultipleYearStockData()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickerCounter As Long
        Dim LastRowInA As Long
        Dim LastRowInI As Long
        Dim PercentageChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        
        ' Store the worksheet name
        WorksheetName = ws.Name
        
        ' Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Initialize ticker counter
        TickerCounter = 2
        
        ' Set the starting row
        j = 2
        
        ' Determine the last non-blank cell in column A
        LastRowInA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all rows
        For i = 2 To LastRowInA
            
            ' Check if the stock ticker has changed
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Write the ticker symbol to column I
                ws.Cells(TickerCounter, 9).Value = ws.Cells(i, 1).Value
                
                ' Calculate and write Quarterly Change to column J
                ws.Cells(TickerCounter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                ' Apply conditional formatting based on the Quarterly Change
                If ws.Cells(TickerCounter, 10).Value < 0 Then
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3  ' Red for negative change
                Else
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4  ' Green for positive change
                End If
                
                ' Calculate and write Percent Change to column K
                If ws.Cells(j, 3).Value <> 0 Then
                    PercentageChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(TickerCounter, 11).Value = Format(PercentageChange, "Percent")
                Else
                    ws.Cells(TickerCounter, 11).Value = Format(0, "Percent")
                End If
                
                ' Calculate and write Total Volume to column L
                ws.Cells(TickerCounter, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                ' Increment the ticker counter
                TickerCounter = TickerCounter + 1
                
                ' Update the start row for the next ticker block
                j = i + 1
                
            End If
        
        Next i
        
        ' Find the last non-blank cell in column I
        LastRowInI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Initialize summary variables
        GreatestVolume = ws.Cells(2, 12).Value
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        
        ' Compute summary statistics
        For i = 2 To LastRowInI
            
            ' Determine the highest total volume
            If ws.Cells(i, 12).Value > GreatestVolume Then
                GreatestVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
            
            ' Determine the highest percentage increase
            If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
            
            ' Determine the highest percentage decrease
            If ws.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        ' Update summary results
        ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
        ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
        ws.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
        
        ' Adjust column widths automatically
        ws.Columns("A:Z").AutoFit
        
    Next ws
    
End Sub

