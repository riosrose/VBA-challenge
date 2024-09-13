Attribute VB_Name = "Module2"
Sub GreatestValuesWithPopUpAndWriteToSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim outputRow As Integer
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        outputRow = 1
        
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        
        For i = 2 To lastRow
            If ws.Cells(i, "L").Value > maxIncrease Then
                maxIncrease = ws.Cells(i, "L").Value
                maxIncreaseTicker = ws.Cells(i, "J").Value
            End If
            
            If ws.Cells(i, "L").Value < maxDecrease Then
                maxDecrease = ws.Cells(i, "L").Value
                maxDecreaseTicker = ws.Cells(i, "J").Value
            End If
            
            If ws.Cells(i, "M").Value > maxVolume Then
                maxVolume = ws.Cells(i, "M").Value
                maxVolumeTicker = ws.Cells(i, "J").Value
            End If
        Next i
        
        ' Output the results for this worksheet without blank rows
        ws.Cells(outputRow, "P").Value = "Metric"
        ws.Cells(outputRow, "Q").Value = "Ticker"
        ws.Cells(outputRow, "R").Value = "Value"
        
        outputRow = outputRow + 1
        
        ws.Cells(outputRow, "P").Value = "Greatest % Increase"
        ws.Cells(outputRow, "Q").Value = maxIncreaseTicker
        ws.Cells(outputRow, "R").Value = Format(maxIncrease, "0.00%")
        
        outputRow = outputRow + 1
        
        ws.Cells(outputRow, "P").Value = "Greatest % Decrease"
        ws.Cells(outputRow, "Q").Value = maxDecreaseTicker
        ws.Cells(outputRow, "R").Value = Format(maxDecrease, "0.00%")
        
        outputRow = outputRow + 1
        
        ws.Cells(outputRow, "P").Value = "Greatest Total Volume"
        ws.Cells(outputRow, "Q").Value = maxVolumeTicker
        ws.Cells(outputRow, "R").Value = maxVolume ' No formatting applied
        
        outputRow = outputRow + 1
    Next ws
End Sub

